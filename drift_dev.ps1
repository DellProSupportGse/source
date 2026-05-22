<#
.SYNOPSIS
    DriFT 2.0 - Driver and Firmware Tool.

.DESCRIPTION
    DriFT 2.0 analyzes Dell support collections and compares installed firmware,
    drivers, BIOS, operating system updates, and platform-specific compatibility data
    against the appropriate Dell catalog source.

    The public Invoke-RunDriFT entry point is preserved so DriFT can still be launched
    directly from GitHub one-liners, while the internal engine is organized into
    documented functions for extraction, platform detection, inventory normalization,
    catalog loading, matching, reporting, telemetry, debug export, and cleanup.

    Supported collection and platform paths include:
      - 17G SupportAssist / TSR collections using Redfish inventory data.
      - Legacy 16G and older TSR collections using DCIM SoftwareIdentity XML.
      - DSET collections that predate TSR, including password-protected DSET ZIPs.
      - PowerEdge server catalog matching through Dell Catalog.xml.
      - AX / Azure Stack HCI systems using ASHCI-Catalog.xml first, with Catalog.xml fallback.
      - VMware ESXi and vSAN systems with Broadcom Compatibility Guide driver links.
      - Precision workstation first-pass support through Precision-aware platform detection.
      - Cluster comparison reporting with one installed-version column per node.

    DriFT 2.0 normalizes every inventory source into a common object model before
    matching so report generation is consistent across generations and collection
    formats.

.NOTES
    Temporary working files, extracted collections, catalog files, logs, and debug
    exports are stored under %TEMP%\DriFT by default.

    Use -ExportDebugData to export stage-level CSV files such as normalized inventory,
    catalog matches, PCI identity data, and VMware/Broadcom compatibility lookup input.

    Set-StrictMode is intentionally not enabled because TSR, DSET, Redfish, catalog,
    and metadata objects all contain optional or generation-specific fields. Treating
    missing optional properties as terminating errors would make parser behavior less
    reliable across mixed Dell collection formats.
#>
Set-StrictMode -Off

#region Constants / Types

$script:DriFTVersion = 'DriFT_v2.00DEV'
$script:DriFTDellDownloadRoot = 'https://downloads.dell.com/'
$script:DriFTDefaultWorkRoot = Join-Path $env:TEMP 'DriFT'

#endregion Constants / Types

#region Public Entry Point

function Invoke-RunDriFT {
<#
.SYNOPSIS
    Public DriFT entry point.

.DESCRIPTION
    Preserves the existing public function name used by GitHub one-liners.
    This function should remain small and only orchestrate the run:
      1. Initialize context/logging.
      2. Resolve input files.
      3. Prepare catalogs.
      4. Import each SupportAssist collection.
      5. Normalize inventory.
      6. Compare to catalog.
      7. Add supplemental checks.
      8. Write reports.
      9. Cleanup.

.PARAMETER InputPath
    One or more SupportAssist Collection ZIP files. If omitted, a file picker is used.

.PARAMETER CluChk
    Enables CluChk-compatible behavior and supplemental XML outputs.

.PARAMETER FileNameGuid
    Optional file name GUID used by CluChk output naming.

.PARAMETER NoTelemetry
    Disables telemetry upload.

.PARAMETER KeepTemp
    Keeps temporary extracted files for troubleshooting.

.PARAMETER ExportDebugData
    Enables debug CSV exports for normalized data, matches, and unmatched records.

.EXAMPLE
    Invoke-RunDriFT

.EXAMPLE
    Invoke-RunDriFT -InputPath C:\Temp\TSR.zip -ExportDebugData

.EXAMPLE
    Invoke-RunDriFT -InputPath @('C:\Temp\node1.zip','C:\Temp\node2.zip') -CluChk -FileNameGuid 1234
#>
    [CmdletBinding()]
    param(
        [Parameter(Position = 0)]
        [string[]]$InputPath,

        [switch]$CluChk,

        [string]$FileNameGuid,

        [switch]$NoTelemetry,

        [switch]$KeepTemp,

        [switch]$ExportDebugData
    )

    $ctx = $null

    try {
        $ctx = Initialize-DriFTRunContext `
            -InputPath $InputPath `
            -CluChk:$CluChk `
            -FileNameGuid $FileNameGuid `
            -NoTelemetry:$NoTelemetry `
            -KeepTemp:$KeepTemp `
            -ExportDebugData:$ExportDebugData

        Write-DriFTBanner -Context $ctx

        if (-not $ctx.NoTelemetry) {
            Write-DriFTTelemetry -Context $ctx
        }

        $catalogSet = Initialize-DriFTCatalogSet -Context $ctx

        $collections = Import-DriFTSupportAssistCollections -Context $ctx

        $allReportRows = @()
        $allBiosConfigRows = @()
        $allSwitchMapRows = @()
        $allSelRows = @()

        foreach ($collection in $collections) {

            # Track active collection for downstream report generation.
            if ($ctx.PSObject.Properties.Name -notcontains 'CurrentCollection') {
                $ctx | Add-Member -MemberType NoteProperty -Name CurrentCollection -Value $collection -Force
            }
            else {
                $ctx.CurrentCollection = $collection
            }

            # Write the HTML report beside the TSR/SupportAssist ZIP being processed.
            $sourceFolder = Split-Path -Parent $collection.SourcePath
            if ($sourceFolder -and (Test-Path -LiteralPath $sourceFolder)) {
                $ctx.OutputRoot = $sourceFolder
            }
            Write-DriFTLog -Context $ctx -Message "Processing collection: $($collection.SourcePath)" -Level Info

            $system = Get-DriFTSystemIdentity -Collection $collection -Context $ctx
            $os = Get-DriFTOperatingSystem -Collection $collection -Context $ctx
            $inventory = @(Get-DriFTInstalledInventory -Collection $collection -Context $ctx | Where-Object { $null -ne $_ })
            Write-DriFTLog -Context $ctx -Message "Normalized inventory rows: $(@($inventory).Count)" -Level Info -Indent 1

            if (-not $inventory -or @($inventory).Count -eq 0) {
                Write-DriFTLog -Context $ctx -Message "WARNING: No installed inventory rows were returned for this collection." -Level Warn -Indent 1
            }

            Export-DriFTDebugData -Context $ctx -Name "$($system.ServiceTag)_NormalizedInventory.csv" -InputObject $inventory

            $filteredCatalogs = Get-DriFTApplicableCatalogRows `
                -CatalogSet $catalogSet `
                -System $system `
                -OperatingSystem $os `
                -Context $ctx

            $catalogIndex = New-DriFTCatalogIndex -CatalogRows $filteredCatalogs.AllRows -Context $ctx

            $matches = Compare-DriFTInventoryToCatalog `
                -Inventory $inventory `
                -CatalogIndex $catalogIndex `
                -CatalogRows $filteredCatalogs `
                -System $system `
                -OperatingSystem $os `
                -Context $ctx

            Export-DriFTDebugData -Context $ctx -Name "$($system.ServiceTag)_CatalogMatches.csv" -InputObject $matches

            $reportRows = New-DriFTReportRows `
                -Matches $matches `
                -System $system `
                -OperatingSystem $os `
                -Context $ctx

            foreach ($row in @($reportRows)) { if ($null -ne $row) { $allReportRows += @($row) } }

            $unmatchedPciRows = New-DriFTUnmatchedPciReportRows `
                -Matches $matches `
                -System $system `
                -OperatingSystem $os `
                -Context $ctx

            foreach ($row in @($unmatchedPciRows)) { if ($null -ne $row) { $allReportRows += @($row) } }

            if ($os.Family -eq 'Windows') {
                $driverRows = Add-DriFTWindowsDriverRows `
                    -Inventory $inventory `
                    -CatalogRows $filteredCatalogs `
                    -System $system `
                    -OperatingSystem $os `
                    -Context $ctx

                foreach ($row in @($driverRows)) { if ($null -ne $row) { $allReportRows += @($row) } }

                $osRows = Add-DriFTWindowsUpdateRows `
                    -System $system `
                    -OperatingSystem $os `
                    -Context $ctx

                foreach ($row in @($osRows | Where-Object { $null -ne $_ })) { $allReportRows += @($row) }
            }

            if ($os.Family -eq 'VMware' -or $os.RawName -imatch 'vSAN|VMware|ESXi' -or $os.DisplayName -imatch 'vSAN|VMware|ESXi') {
                $vmwareRows = Add-DriFTVmwareCompatibilityRows `
                    -Inventory $inventory `
                    -System $system `
                    -OperatingSystem $os `
                    -Context $ctx

                foreach ($row in @($vmwareRows)) { if ($null -ne $row) { $allReportRows += @($row) } }
            }

            $selRows = Get-DriFTSelHealthRows -Collection $collection -System $system -Context $ctx
            foreach ($row in @($selRows)) { if ($null -ne $row) { $allSelRows += @($row) } }

            if ($ctx.CluChk) {
                $biosRows = Get-DriFTBiosAndIdracConfigRows `
                    -Collection $collection `
                    -System $system `
                    -OperatingSystem $os `
                    -Context $ctx

                foreach ($row in @($biosRows)) { if ($null -ne $row) { $allBiosConfigRows += @($row) } }

                $switchRows = Get-DriFTSwitchPortMapRows `
                    -Collection $collection `
                    -System $system `
                    -Context $ctx

                foreach ($row in @($switchRows)) { if ($null -ne $row) { $allSwitchMapRows += @($row) } }
            }
        }

        $reportPath = Write-DriFTHtmlReport `
            -Rows @($allReportRows) `
            -Context $ctx `
            -CatalogSet $catalogSet

        if ($ctx.CluChk) {
            Write-DriFTCluChkOutputs `
                -Context $ctx `
                -BiosConfigRows @($allBiosConfigRows) `
                -SwitchMapRows @($allSwitchMapRows) `
                -SelRows @($allSelRows)
        }

        Write-DriFTLog -Context $ctx -Message "Report written to: $reportPath" -Level Success

        if ($reportPath -and (Test-Path -LiteralPath $reportPath -PathType Leaf)) {
            try {
                Invoke-Item -LiteralPath $reportPath
            }
            catch {
                Write-DriFTLog -Context $ctx -Message "Report was written but could not be opened automatically: $($_.Exception.Message)" -Level Warn
            }
        }

        return $reportPath
    }
    catch {
        $err = $_
        $line = $err.InvocationInfo.ScriptLineNumber
        $cmd  = $err.InvocationInfo.Line
        $msg  = $err.Exception.Message
        $stack = $err.ScriptStackTrace

        Write-Host ""
        Write-Host "========== DriFT DEBUG ==========" -ForegroundColor Red
        Write-Host ("Line    : {0}" -f $line) -ForegroundColor Yellow
        Write-Host ("Message : {0}" -f $msg) -ForegroundColor Yellow
        Write-Host ("Command : {0}" -f $cmd) -ForegroundColor Cyan
        Write-Host ""
        Write-Host "Stack Trace:" -ForegroundColor Magenta
        Write-Host $stack
        Write-Host "=================================" -ForegroundColor Red

        throw
    }
    finally {
        if ($ctx) {
            Complete-DriFTRun -Context $ctx
        }
    }
}

#endregion Public Entry Point

#region Run Context / Logging / UI

function Initialize-DriFTRunContext {
<#
.SYNOPSIS
    Creates the per-run state object.

.DESCRIPTION
    Centralizes run-specific paths, switches, log path, report ID, and input paths.
    Avoids global variables and prevents accidental session-scope cleanup.
#>
    [CmdletBinding()]
    param(
        [string[]]$InputPath,
        [switch]$CluChk,
        [string]$FileNameGuid,
        [switch]$NoTelemetry,
        [switch]$KeepTemp,
        [switch]$ExportDebugData
    )

    $dateStamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $workRoot = $script:DriFTDefaultWorkRoot
    $logRoot = Join-Path $workRoot 'Logs'
    $extractRoot = Join-Path $workRoot 'Extract'
    $redfishRoot = Join-Path $workRoot 'Redfish'
    $catalogRoot = Join-Path $workRoot 'Catalog'
    $debugRoot = Join-Path $workRoot 'Debug'
    $runRoot = Join-Path $extractRoot ('Run_' + ([guid]::NewGuid().Guid.Substring(0, 8)))

    foreach ($path in @($workRoot, $extractRoot, $redfishRoot, $catalogRoot, $debugRoot, $runRoot)) {
        if (-not (Test-Path -Path $path -PathType Container)) {
            New-Item -ItemType Directory -Force -Path $path | Out-Null
        }
    }

    if (-not (Test-Path -Path $logRoot -PathType Container)) {
        New-Item -ItemType Directory -Force -Path $logRoot | Out-Null
    }

    $logPath = Join-Path $logRoot "DriFT_$dateStamp.log"
    try { Start-Transcript -Path $logPath -NoClobber -ErrorAction Stop | Out-Null } catch { }

    [PSCustomObject]@{
        Version       = $script:DriFTVersion
        DisplayVer    = ($script:DriFTVersion -replace '^.*_v', '')
        RunId         = [guid]::NewGuid().Guid
        DateStamp     = $dateStamp
        Started       = Get-Date
        InputPath     = $InputPath
        CluChk        = [bool]$CluChk
        FileNameGuid  = $FileNameGuid
        NoTelemetry   = [bool]$NoTelemetry
        KeepTemp      = [bool]$KeepTemp
        ExportDebugData = [bool]$ExportDebugData
        WorkRoot      = $workRoot
        RunRoot       = $runRoot
        RedfishRoot   = $redfishRoot
        CatalogRoot   = $catalogRoot
        DebugRoot     = $debugRoot
        ExportDebugDataRoot = $debugRoot
        LogPath       = $logPath
        OutputRoot    = $null
        ServiceTags   = New-Object System.Collections.Generic.List[string]
    }
}

function Write-DriFTBanner {
<#
.SYNOPSIS
    Writes the startup banner.

.DESCRIPTION
    Keeps presentation code out of the orchestrator.
#>
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Context)

$VerLine = "|v$($Context.DisplayVer)"+" "*(43-$($Context.DisplayVer).length)+"|"

$text = @"
+--------------------------------------------+
$VerLine 
|          __   __     ___ ___               |
|         |  \ |__) | |__   |                |
|         |__/ |  \ | |     |                |
|                                            |
|             Driver & Firmware Tool         |
|                         By: Jim Gandy      |
|                                            |
+--------------------------------------------+

"@

    Write-Host $text
    Write-DriFTLog -Context $Context -Message "Starting log: $($Context.LogPath)" -Level Info
}

function Write-DriFTLog {
<#
.SYNOPSIS
    Writes a consistent status message.

.DESCRIPTION
    Simple wrapper for Write-Host today. Can later be replaced with structured logging.
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Context,
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('Info','Warn','Error','Success')]
        [string]$Level = 'Info',
        [int]$Indent = 0
    )

    $prefix = '  ' * $Indent
    $color = switch ($Level) {
        'Warn'    { 'Yellow' }
        'Error'   { 'Red' }
        'Success' { 'Green' }
        default   { 'Gray' }
    }

    Write-Host "$prefix$Message" -ForegroundColor $color
}

function Export-DriFTDebugData {
<#
.SYNOPSIS
    Exports stage debug data when -ExportDebugData is enabled.

.DESCRIPTION
    Debug export should never break a DriFT run. If ExportDebugDataRoot is missing,
    it falls back to DebugRoot and then %TEMP%\DriFT\Debug.
#>
    [CmdletBinding()]
    param(
        [AllowNull()]$Context,
        [Parameter(Mandatory)][string]$Name,
        [AllowNull()][object]$InputObject
    )

    try {
        if ($null -eq $Context) { return $null }

        $enabled = $false
        foreach ($propName in @('ExportDebugData','Debug','EnableDebugExports')) {
            if ($Context.PSObject.Properties.Name -contains $propName -and [bool]$Context.$propName) {
                $enabled = $true
                break
            }
        }

        if (-not $enabled) { return $null }

        $debugRoot = Get-DriFTFirstNonEmpty `
            $Context.ExportDebugDataRoot `
            $Context.DebugRoot `
            (Join-Path $env:TEMP 'DriFT\Debug')

        if ($Context.PSObject.Properties.Name -notcontains 'ExportDebugDataRoot') {
            $Context | Add-Member -MemberType NoteProperty -Name ExportDebugDataRoot -Value $debugRoot -Force
        }
        elseif ([string]::IsNullOrWhiteSpace([string]$Context.ExportDebugDataRoot)) {
            $Context.ExportDebugDataRoot = $debugRoot
        }

        if (-not (Test-Path -LiteralPath $debugRoot -PathType Container)) {
            New-Item -ItemType Directory -Force -Path $debugRoot | Out-Null
        }

        $safeName = ([string]$Name) -replace '[<>:"/\\|?*]', '_'
        if ($safeName -notmatch '\.(csv|json|txt)$') {
            $safeName = "$safeName.csv"
        }

        $path = Join-Path $debugRoot $safeName

        if ($null -eq $InputObject) {
            '' | Out-File -LiteralPath $path -Encoding UTF8
        }
        elseif ($safeName -match '\.json$') {
            $InputObject | ConvertTo-Json -Depth 20 | Out-File -LiteralPath $path -Encoding UTF8
        }
        elseif ($safeName -match '\.csv$') {
            @($InputObject) | Export-Csv -NoTypeInformation -LiteralPath $path -Encoding UTF8
        }
        else {
            $InputObject | Out-File -LiteralPath $path -Encoding UTF8
        }

        Write-DriFTLog -Context $Context -Message "Debug export written: $path" -Level Info -Indent 1

        return $path
    }
    catch {
        try {
            Write-DriFTLog -Context $Context -Message "Debug export failed for '$Name': $($_.Exception.Message)" -Level Warn -Indent 1
        }
        catch { }

        return $null
    }
}

function Complete-DriFTRun {
<#
.SYNOPSIS
    Performs final cleanup.

.DESCRIPTION
    Stops transcript and removes temporary extraction folders unless -KeepTemp was used.
#>
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Context)

    if (-not $Context.KeepTemp) {
        foreach ($path in @($Context.RunRoot, $Context.RedfishRoot)) {
            if ($path -and (Test-Path $path)) {
                try { Remove-Item -Path $path -Recurse -Force -ErrorAction Stop }
                catch { Write-Warning "Failed to remove temp path $path. It can be deleted manually." }
            }
        }
    }

    try { Stop-Transcript | Out-Null } catch {}
}

#endregion Run Context / Logging / UI

#region Telemetry

function Write-DriFTTelemetry {
<#
.SYNOPSIS
    Records optional DriFT telemetry.

.DESCRIPTION
    This function intentionally does not embed SAS tokens. The SAS URL should be provided
    through an environment variable or external configuration. This avoids committing a
    writable storage token to GitHub.

    Expected environment variable:
      DRIFT_TABLE_SAS_URL
#>
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Context)

    $sasUrl = [Environment]::GetEnvironmentVariable('DRIFT_TABLE_SAS_URL', 'User')
    if (-not $sasUrl) {
        $sasUrl = [Environment]::GetEnvironmentVariable('DRIFT_TABLE_SAS_URL', 'Machine')
    }

    if (-not $sasUrl) {
        Write-DriFTLog -Context $Context -Message 'Telemetry skipped: DRIFT_TABLE_SAS_URL is not configured.' -Level Warn -Indent 1
        return
    }

    # Intentionally minimal in the foundation. Port existing Azure Table write logic here.
    Write-DriFTLog -Context $Context -Message 'Telemetry configured. Upload logic should be ported into Write-DriFTTelemetry.' -Level Info -Indent 1
}

#endregion Telemetry

#region Input / Extraction

function Resolve-DriFTInputPath {
<#
.SYNOPSIS
    Resolves SupportAssist input files.

.DESCRIPTION
    Uses the supplied -InputPath when present. If omitted, opens the same ZIP picker
    behavior used by the legacy script.
#>
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Context)

    if ($Context.InputPath -and @($Context.InputPath).Count -gt 0) {
        $resolved = @()

        foreach ($item in @($Context.InputPath)) {
            if ([string]::IsNullOrWhiteSpace([string]$item)) { continue }

            $candidate = ([string]$item).Trim().Trim('"').Trim("'")

            if (Test-Path -LiteralPath $candidate -PathType Leaf) {
                $resolved += @((Resolve-Path -LiteralPath $candidate).Path)
                continue
            }

            # Last-resort DSET friendliness: if PowerShell wildcard parsing or
            # pasted paths with brackets caused issues, try a filename search in
            # the parent folder using -Filter, not wildcard path matching.
            $parent = Split-Path -Path $candidate -Parent
            $leaf = Split-Path -Path $candidate -Leaf

            if ($parent -and (Test-Path -LiteralPath $parent -PathType Container)) {
                $found = Get-ChildItem -LiteralPath $parent -Filter $leaf -File -ErrorAction SilentlyContinue | Select-Object -First 1
                if ($found) {
                    $resolved += @($found.FullName)
                    continue
                }
            }

            $resolved += @($candidate)
        }

        return @($resolved)
    }

    Add-Type -AssemblyName System.Windows.Forms
    $dialog = New-Object System.Windows.Forms.OpenFileDialog -Property @{
        Multiselect      = $true
        Title            = 'Please Select One or More SupportAssist File(s).'
        InitialDirectory = $env:USERPROFILE
        Filter           = 'ZIP (*.zip)|*.zip'
    }

    $dialog.ShowDialog((New-Object System.Windows.Forms.Form -Property @{ TopMost = $true })) | Out-Null
    return @($dialog.FileNames)
}



function Get-DriFT7ZipPath {
<#
.SYNOPSIS
    Finds a usable 7-Zip executable.

.DESCRIPTION
    Password-protected DSET archives cannot be extracted by .NET ZipArchive.
    7-Zip supports encrypted ZIP entries and can extract DSET archives with the
    legacy password.
#>
    [CmdletBinding()]
    param()

    $candidates = @(
        "$env:ProgramFiles\7-Zip\7z.exe",
        "${env:ProgramFiles(x86)}\7-Zip\7z.exe",
        "$env:ProgramData\chocolatey\bin\7z.exe",
        "$env:ProgramData\chocolatey\bin\7za.exe",
        "7z.exe",
        "7za.exe"
    )

    foreach ($candidate in $candidates) {
        if ([string]::IsNullOrWhiteSpace($candidate)) { continue }

        try {
            $cmd = Get-Command $candidate -ErrorAction SilentlyContinue
            if ($cmd) { return $cmd.Source }
        }
        catch { }

        if (Test-Path -LiteralPath $candidate -PathType Leaf) {
            return $candidate
        }
    }

    return $null
}

function Expand-DriFTZipFileWith7Zip {
<#
.SYNOPSIS
    Extracts ZIP files using 7-Zip.

.DESCRIPTION
    Used for password-protected DSET ZIP archives. DSET default password is "dell".
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Destination,
        [AllowNull()][string]$Password,
        [Parameter(Mandatory)]$Context
    )

    $sevenZip = Get-DriFT7ZipPath
    if ([string]::IsNullOrWhiteSpace($sevenZip)) {
        throw "7-Zip was not found. Password-protected DSET ZIP extraction requires 7-Zip. Install 7-Zip or manually extract the DSET ZIP with password 'dell' and run DriFT against the extracted folder."
    }

    if (-not (Test-Path -LiteralPath $Destination -PathType Container)) {
        New-Item -ItemType Directory -Force -Path $Destination | Out-Null
    }

    $args = @(
        'x',
        '-y',
        "-o$Destination"
    )

    if (-not [string]::IsNullOrWhiteSpace($Password)) {
        $args += "-p$Password"
    }

    $args += $Path

    Write-DriFTLog -Context $Context -Message "Trying 7-Zip extraction for ZIP: $Path" -Level Info -Indent 1

    $process = Start-Process `
        -FilePath $sevenZip `
        -ArgumentList $args `
        -NoNewWindow `
        -Wait `
        -PassThru `
        -RedirectStandardOutput (Join-Path $Destination 'DriFT_7zip_stdout.txt') `
        -RedirectStandardError (Join-Path $Destination 'DriFT_7zip_stderr.txt')

    if ($process.ExitCode -ne 0) {
        $stderr = ''
        $stderrPath = Join-Path $Destination 'DriFT_7zip_stderr.txt'
        if (Test-Path -LiteralPath $stderrPath -PathType Leaf) {
            $stderr = Get-Content -Raw -LiteralPath $stderrPath -ErrorAction SilentlyContinue
        }

        throw "7-Zip extraction failed with exit code $($process.ExitCode). $stderr"
    }

    $anyExtracted = Get-ChildItem -LiteralPath $Destination -Recurse -Force -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -notmatch '^DriFT_7zip_(stdout|stderr)\.txt$' } |
        Select-Object -First 1

    if (-not $anyExtracted) {
        throw "7-Zip extraction completed but no files were extracted."
    }

    Write-DriFTLog -Context $Context -Message "7-Zip extraction succeeded." -Level Success -Indent 1

    return $Destination
}

function Test-DriFTZipLooksLikeDset {
<#
.SYNOPSIS
    Determines whether a ZIP filename/path appears to be a DSET archive.
#>
    [CmdletBinding()]
    param([AllowNull()][string]$Path)

    return ([string]$Path -imatch 'DSET|DSET_Report|DSET Report')
}


function Expand-DriFTZipFileWithShell {
<#
.SYNOPSIS
    Extracts ZIP files using Windows Shell.Application fallback.

.DESCRIPTION
    Older DSET ZIP archives can contain entries that .NET ZipArchive fails to
    decode with "Found invalid data while decoding." Windows Explorer/Shell can
    often extract those same archives successfully, so this is used as a fallback.
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Destination
    )

    if (-not (Test-Path -LiteralPath $Destination -PathType Container)) {
        New-Item -ItemType Directory -Force -Path $Destination | Out-Null
    }

    $shell = New-Object -ComObject Shell.Application
    $zipNs = $shell.NameSpace((Resolve-Path -LiteralPath $Path).Path)
    $dstNs = $shell.NameSpace((Resolve-Path -LiteralPath $Destination).Path)

    if (-not $zipNs -or -not $dstNs) {
        throw "Shell.Application could not open ZIP or destination. ZIP='$Path'; Destination='$Destination'"
    }

    # 0x10 = Yes to all / no UI prompts where possible
    # 0x400 = Do not display progress dialog
    # 0x4 = No UI
    $copyFlags = 0x10 -bor 0x400 -bor 0x4
    $dstNs.CopyHere($zipNs.Items(), $copyFlags)

    # Shell copy is async. Wait briefly until extracted content appears.
    $deadline = (Get-Date).AddSeconds(30)
    do {
        Start-Sleep -Milliseconds 300
        $items = Get-ChildItem -LiteralPath $Destination -Recurse -Force -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($items) { break }
    } while ((Get-Date) -lt $deadline)

    return $Destination
}

function Expand-DriFTZipFileEntryByEntry {
<#
.SYNOPSIS
    Extracts ZIP entries one by one and skips unreadable DSET entries.

.DESCRIPTION
    This is used before Shell fallback so good entries from partially problematic
    DSET archives are still extracted. Failures are logged and extraction continues.
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Destination,
        $Context
    )

    if ($null -eq $Context) {
        $Context = [PSCustomObject]@{
            QuietExtractionWarnings = $true
        }
    }

    Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue

    if (-not (Test-Path -LiteralPath $Destination -PathType Container)) {
        New-Item -ItemType Directory -Force -Path $Destination | Out-Null
    }

    $zip = [System.IO.Compression.ZipFile]::OpenRead($Path)
    $extracted = 0
    $failed = 0

    try {
        foreach ($entry in $zip.Entries) {
            if ([string]::IsNullOrWhiteSpace($entry.FullName)) { continue }

            $relative = $entry.FullName -replace '/', '\'
            if ($relative.EndsWith('\')) {
                $dirPath = Join-Path $Destination $relative
                if (-not (Test-Path -LiteralPath $dirPath -PathType Container)) {
                    New-Item -ItemType Directory -Force -Path $dirPath | Out-Null
                }
                continue
            }

            $targetPath = Join-Path $Destination $relative
            $targetDir = Split-Path -Parent $targetPath
            if (-not (Test-Path -LiteralPath $targetDir -PathType Container)) {
                New-Item -ItemType Directory -Force -Path $targetDir | Out-Null
            }

            try {
                [System.IO.Compression.ZipFileExtensions]::ExtractToFile($entry, [string]$targetPath, $true)
                $extracted++
            }
            catch {
                $failed++

                $isLowValueDsetLog = ($entry.FullName -imatch '\.(log|txt)$')
                $showExtractionWarning = $true

                if ($isLowValueDsetLog -and $Context.QuietExtractionWarnings -ne $false) {
                    $showExtractionWarning = $false
                }

                if ($showExtractionWarning) {
                    Write-DriFTLog -Context $Context -Message "ZIP entry skipped during extraction: $($entry.FullName) - $($_.Exception.Message)" -Level Warn -Indent 1
                }
            }
        }
    }
    finally {
        if ($zip) { $zip.Dispose() }
    }

    Write-DriFTLog -Context $Context -Message "ZIP extraction completed. Extracted=$extracted; Skipped=$failed" -Level Info -Indent 1

    return $Destination
}


function Expand-DriFTZipFile {
<#
.SYNOPSIS
    Extracts a ZIP file to a destination folder.

.DESCRIPTION
    Uses entry-by-entry .NET extraction first. If .NET cannot decode an older DSET
    archive, it falls back to Windows Shell.Application extraction, which handles
    some legacy ZIP structures better.
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Destination,
        $Context
    )

    if ($null -eq $Context) {
        $Context = [PSCustomObject]@{
            QuietExtractionWarnings = $true
        }
    }

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Input file was not found: $Path"
    }

    if (-not (Test-Path -LiteralPath $Destination -PathType Container)) {
        New-Item -ItemType Directory -Force -Path $Destination | Out-Null
    }

    # DSET archives are commonly password-protected with password 'dell'.
    # .NET ZipArchive cannot extract encrypted ZIP entries, so try 7-Zip first
    # for files that look like DSET reports.
    if (Test-DriFTZipLooksLikeDset -Path $Path) {
        try {
            Expand-DriFTZipFileWith7Zip -Path $Path -Destination $Destination -Password 'dell' -Context $Context | Out-Null
            return $Destination
        }
        catch {
            Write-DriFTLog -Context $Context -Message "7-Zip/DSET password extraction did not complete: $($_.Exception.Message)" -Level Warn -Indent 1
            Write-DriFTLog -Context $Context -Message "Continuing with standard extraction fallback..." -Level Warn -Indent 1
        }
    }

    try {
        Expand-DriFTZipFileEntryByEntry -Path $Path -Destination $Destination -Context $Context | Out-Null

        $anyExtracted = Get-ChildItem -LiteralPath $Destination -Recurse -Force -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($anyExtracted) {
            return $Destination
        }

        throw "No files were extracted by .NET ZipArchive."
    }
    catch {
        Write-DriFTLog -Context $Context -Message "Primary ZIP extraction failed: $($_.Exception.Message)" -Level Warn -Indent 1
        Write-DriFTLog -Context $Context -Message "Trying Shell.Application ZIP extraction fallback..." -Level Warn -Indent 1

        try {
            Expand-DriFTZipFileWithShell -Path $Path -Destination $Destination | Out-Null

            $anyExtracted = Get-ChildItem -LiteralPath $Destination -Recurse -Force -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($anyExtracted) {
                Write-DriFTLog -Context $Context -Message "Shell.Application extraction fallback succeeded." -Level Success -Indent 1
                return $Destination
            }

            throw "Shell.Application did not extract any files."
        }
        catch {
            throw "Unable to extract ZIP '$Path'. Primary and fallback extraction failed. Last error: $($_.Exception.Message)"
        }
    }
}

function Import-DriFTSupportAssistCollections {
<#
.SYNOPSIS
    Extracts and classifies all input SupportAssist collections.

.DESCRIPTION
    Expands the outer TSR ZIP and any inner inventory ZIPs, then calls
    Get-DriFTCollectionType to decide which parser path applies:
      - LegacyTSR for 16G and older XML inventory
      - TSR17G for 17G viewer.html/redfishidracwalk
      - SAEXml for SupportAssist Enterprise XML
      - SAEJson for unsupported SAE JSON detection
#>
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Context)

    $inputFiles = Resolve-DriFTInputPath -Context $Context
    if (-not $inputFiles -or @($inputFiles).Count -eq 0) {
        throw 'No SupportAssist input file was selected or provided.'
    }

    $collections = @()

    foreach ($input in $inputFiles) {
        if (-not (Test-Path -LiteralPath $input -PathType Leaf)) {
            throw "Input file was not found: $input"
        }

        $input = [string]$input
        $input = $input.Trim().Trim('"').Trim("'")

        $baseName = [IO.Path]::GetFileNameWithoutExtension($input)
        $dest = Join-Path $Context.RunRoot $baseName
        Write-DriFTLog -Context $Context -Message "Extracting $input" -Level Info

        Expand-DriFTZipFile -Path $input -Destination $dest -Context $Context | Out-Null
        Expand-DriFTNestedInventoryZips -Root $dest -Context $Context

        Write-DriFTLog -Context $Context -Message "Classifying extracted collection..." -Level Info -Indent 1
        $collection = Get-DriFTCollectionType -Root ([string]$dest) -SourcePath ([string]$input) -Context $Context
        Write-DriFTLog -Context $Context -Message "Collection type detected: $($collection.Type)" -Level Success -Indent 1
        $collections += @($collection)
    }

    if (@($collections).Count -gt 0) {
        $Context.OutputRoot = Split-Path -Path $inputFiles[0] -Parent
    }

    return @($collections)
}

function Expand-DriFTNestedInventoryZips {
<#
.SYNOPSIS
    Extracts inner SupportAssist ZIPs.

.DESCRIPTION
    Preserves legacy behavior of extracting nested ZIP files while excluding obvious
    non-inventory payloads such as thermal and dump logs.
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Root,
        [Parameter(Mandatory)]$Context
    )

    $innerZips = Get-ChildItem -Path $Root -Filter '*.zip' -File -Recurse -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -notmatch 'thermal|dumplog' } |
        Sort-Object FullName -Unique

    foreach ($zip in $innerZips) {
        try {
            Write-DriFTLog -Context $Context -Message "Extracting nested ZIP: $($zip.Name)" -Level Info -Indent 1
            Expand-DriFTZipFile -Path ([string]$zip.FullName) -Destination ([string]$Root) -Context $Context | Out-Null
        }
        catch {
            Write-DriFTLog -Context $Context -Message "Failed to extract nested ZIP $($zip.FullName): $($_.Exception.Message)" -Level Warn -Indent 1
        }
    }
}

function Get-DriFTCollectionType {
<#
.SYNOPSIS
    Classifies an extracted collection.

.DESCRIPTION
    Determines the parser path without parsing the entire collection.
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Root,
        [Parameter(Mandatory)][string]$SourcePath,
        [Parameter(Mandatory)]$Context
    )

    $inventoryDir = Get-ChildItem -Path $Root -Filter inventory -Directory -Recurse -Force -ErrorAction SilentlyContinue |
        Select-Object -First 1

    $hasLegacySoftwareIdentity = $false
    if ($inventoryDir) {
        $hasLegacySoftwareIdentity = Test-Path (Join-Path $inventoryDir.FullName 'sysinfo_DCIM_SoftwareIdentity.xml') -PathType Leaf
    }

    $viewerHtml = Get-ChildItem -Path $Root -Filter 'viewer.html' -File -Recurse -Force -ErrorAction SilentlyContinue |
        Select-Object -First 1

    $redfishWalk = Get-ChildItem -Path $Root -Filter 'redfishidracwalk.tar.gz' -File -Recurse -Force -ErrorAction SilentlyContinue |
        Select-Object -First 1

    $saeXml = Get-ChildItem -Path $Root -Include 'MaserInfo.xml','Inventory.xml' -File -Recurse -Force -ErrorAction SilentlyContinue |
        Select-Object -First 1

    $saeJson = Get-ChildItem -Path $Root -Filter 'supportassist_output.json' -File -Recurse -Force -ErrorAction SilentlyContinue |
        Select-Object -First 1

    $dsetInventory = Get-ChildItem -Path $Root -Filter 'Inventory.xml' -File -Recurse -Force -ErrorAction SilentlyContinue |
        Where-Object { $_.FullName -match 'OM Server Administrator' -or $_.FullName -match 'logs' } |
        Select-Object -First 1

    $dsetSyssum = Get-ChildItem -Path $Root -Filter 'syssum.xml' -File -Recurse -Force -ErrorAction SilentlyContinue |
        Where-Object { $_.FullName -match 'rawxml|xml|tmpreport' } |
        Select-Object -First 1

    $dsetFwView = Get-ChildItem -Path $Root -Filter 'fwview.xml' -File -Recurse -Force -ErrorAction SilentlyContinue |
        Where-Object { $_.FullName -match 'rawxml|xml|tmpreport' } |
        Select-Object -First 1

    $dsetBiosView = Get-ChildItem -Path $Root -Filter 'biosview.xml' -File -Recurse -Force -ErrorAction SilentlyContinue |
        Where-Object { $_.FullName -match 'rawxml|xml|tmpreport' } |
        Select-Object -First 1

    $type = if ($dsetInventory -or $dsetSyssum -or $dsetFwView -or $dsetBiosView) {
        'DSET'
    }
    elseif ($inventoryDir -and $hasLegacySoftwareIdentity) {
        'LegacyTSR'
    }
    elseif ($viewerHtml -or $redfishWalk) {
        'TSR17G'
    }
    elseif ($saeXml) {
        'SAEXml'
    }
    elseif ($saeJson) {
        'SAEJson'
    }
    else {
        'Unknown'
    }

    [PSCustomObject]@{
        SourcePath       = $SourcePath
        Root             = $Root
        Type             = $type
        InventoryPath    = if ($inventoryDir) { $inventoryDir.FullName } else { $null }
        ViewerHtmlPath   = if ($viewerHtml) { $viewerHtml.FullName } else { $null }
        RedfishWalkPath  = if ($redfishWalk) { $redfishWalk.FullName } else { $null }
        SaeXmlPath       = if ($saeXml) { $saeXml.FullName } else { $null }
        SaeJsonPath      = if ($saeJson) { $saeJson.FullName } else { $null }
        DsetInventoryPath= if ($dsetInventory) { $dsetInventory.FullName } else { $null }
        DsetSyssumPath   = if ($dsetSyssum) { $dsetSyssum.FullName } else { $null }
        DsetFwViewPath   = if ($dsetFwView) { $dsetFwView.FullName } else { $null }
        DsetBiosViewPath = if ($dsetBiosView) { $dsetBiosView.FullName } else { $null }
    }
}

#endregion Input / Extraction

#region Normalized Object Constructors

function New-DriFTSystemInfo {
<#
.SYNOPSIS
    Creates a normalized system identity object.
#>
    [CmdletBinding()]
    param(
        [string]$ServiceTag,
        [string]$PowerEdge,
        [string]$OS,
        [string]$HostName,
        [string]$SystemID,
        [string]$SourceType,
        [string]$SpecialCatalogNeeded = 'NO',
        [string]$S2DCatalogNeeded = 'NO'
    )

    [PSCustomObject]@{
        ServiceTag           = $ServiceTag
        PowerEdge            = $PowerEdge
        OS                   = $OS
        HostName             = $HostName
        SystemID             = $SystemID
        SourceType           = $SourceType
        SpecialCatalogNeeded = $SpecialCatalogNeeded
        S2DCatalogNeeded     = $S2DCatalogNeeded
    }
}

function New-DriFTInventoryItem {
<#
.SYNOPSIS
    Creates a normalized inventory row.

.DESCRIPTION
    Every parser must emit this shape. This lets the matcher work the same way
    for legacy XML, SAE XML, and 17G Redfish.
#>
    [CmdletBinding()]
    param(
        [string]$SourceGeneration,
        [string]$ComponentType,
        [string]$ComponentID,
        [string]$VendorID,
        [string]$DeviceID,
        [string]$SubVendorID,
        [string]$SubDeviceID,
        [string]$Version,
        [string]$Display,
        [string]$ElementName,
        [string]$RelatedItem,
        [string]$Source
    )

    [PSCustomObject]@{
        SourceGeneration = $SourceGeneration
        ComponentType    = $ComponentType
        ComponentID      = $ComponentID
        VendorID         = Convert-DriFTHexId $VendorID
        DeviceID         = Convert-DriFTHexId $DeviceID
        SubVendorID      = Convert-DriFTHexId $SubVendorID
        SubDeviceID      = Convert-DriFTHexId $SubDeviceID
        Version          = $Version
        Display          = $Display
        ElementName      = $ElementName
        RelatedItem      = $RelatedItem
        Source           = $Source
    }
}

function New-DriFTOperatingSystemInfo {
<#
.SYNOPSIS
    Creates a normalized OS object.
#>
    [CmdletBinding()]
    param(
        [string]$RawName,
        [string]$Family,
        [string]$DisplayName,
        [string]$Version,
        [string]$CatalogPackageType = 'LW64',
        [string]$MajorVersion,
        [string]$MinorVersion,
        [string]$Build,
        [bool]$DriverSupport = $true
    )

    [PSCustomObject]@{
        RawName            = $RawName
        Family             = $Family
        DisplayName        = $DisplayName
        Version            = $Version
        CatalogPackageType = $CatalogPackageType
        MajorVersion       = $MajorVersion
        MinorVersion       = $MinorVersion
        Build              = $Build
        DriverSupport      = $DriverSupport
    }
}

function New-DriFTReportRow {
<#
.SYNOPSIS
    Creates a normalized report row.
#>
    [CmdletBinding()]
    param(
        [string]$ServiceTag,
        [string]$PowerEdge,
        [string]$OS,
        [string]$Type,
        [string]$Category,
        [string]$Name,
        [string]$InstalledVersion,
        [string]$AvailableVersion,
        [string]$CatalogInfo,
        [string]$Criticality,
        [string]$ReleaseDate,
        [string]$URL,
        [string]$Details,
        [string]$SourceType = 'Unknown'
    )

    # If a supplemental/non-Dell row has no direct download URL but does have
    # documentation/details, reuse that link in the Download Link column so the
    # HTML report remains consistent with legacy DriFT behavior.
    if ([string]::IsNullOrWhiteSpace([string]$URL) -and
        -not [string]::IsNullOrWhiteSpace([string]$Details)) {
        $URL = $Details
    }

    [PSCustomObject]@{
        ServiceTag       = $ServiceTag
        PowerEdge        = $PowerEdge
        OS               = $OS
        Type             = $Type
        Category         = $Category
        Name             = $Name
        InstalledVersion = $InstalledVersion
        AvailableVersion = $AvailableVersion
        CatalogInfo      = $CatalogInfo
        Criticality      = $Criticality
        ReleaseDate      = $ReleaseDate
        URL              = $URL
        Details          = $Details
        SourceType       = $SourceType
    }
}

#endregion Normalized Object Constructors

#region Parser Dispatch

function Get-DriFTSystemIdentity {
<#
.SYNOPSIS
    Returns normalized system identity.

.DESCRIPTION
    Dispatches to parser-specific identity functions based on collection type.
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Collection,
        [Parameter(Mandatory)]$Context
    )

    switch ($Collection.Type) {
        'LegacyTSR' { return Get-DriFTLegacyTsrSystemIdentity -Collection $Collection -Context $Context }
        'DSET'      { return Get-DriFTDsetSystemIdentity -Collection $Collection -Context $Context }
        'TSR17G'    { return Get-DriFT17GSystemIdentity -Collection $Collection -Context $Context }
        'SAEXml'    { return Get-DriFTSaeXmlSystemIdentity -Collection $Collection -Context $Context }
        default     { throw "Unsupported or unknown SupportAssist collection type: $($Collection.Type)" }
    }
}

function Get-DriFTOperatingSystem {
<#
.SYNOPSIS
    Returns normalized operating system information.

.DESCRIPTION
    Dispatches to parser-specific OS functions and normalizes catalog OS values.
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Collection,
        [Parameter(Mandatory)]$Context
    )

    switch ($Collection.Type) {
        'LegacyTSR' { return Get-DriFTLegacyTsrOperatingSystem -Collection $Collection -Context $Context }
        'DSET'      { return Get-DriFTDsetOperatingSystem -Collection $Collection -Context $Context }
        'TSR17G'    { return Get-DriFT17GOperatingSystem -Collection $Collection -Context $Context }
        'SAEXml'    { return Get-DriFTSaeXmlOperatingSystem -Collection $Collection -Context $Context }
        default     { throw "Unsupported or unknown SupportAssist collection type: $($Collection.Type)" }
    }
}

function Get-DriFTInstalledInventory {
<#
.SYNOPSIS
    Returns normalized installed inventory.

.DESCRIPTION
    Dispatches to parser-specific inventory importers. Every parser returns the same
    New-DriFTInventoryItem shape.
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Collection,
        [Parameter(Mandatory)]$Context
    )

    switch ($Collection.Type) {
        'LegacyTSR' { return Import-DriFTLegacyTsrInventory -Collection $Collection -Context $Context }
        'DSET'      { return Import-DriFTDsetInventory -Collection $Collection -Context $Context }
        'TSR17G'    { return Import-DriFT17GInventory -Collection $Collection -Context $Context }
        'SAEXml'    { return Import-DriFTSaeXmlInventory -Collection $Collection -Context $Context }
        default     { throw "Unsupported or unknown SupportAssist collection type: $($Collection.Type)" }
    }
}

#endregion Parser Dispatch


#region DSET Parser

function Get-DriFTDsetInventoryXml {
<#
.SYNOPSIS
    Loads DSET OM Server Administrator Inventory.xml.

.DESCRIPTION
    DSET predates TSR and often stores the useful inventory in
    logs\OM Server Administrator\Inventory.xml as UTF-16 XML.
#>
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Collection)

    $path = $Collection.DsetInventoryPath
    if ([string]::IsNullOrWhiteSpace($path) -or -not (Test-Path -LiteralPath $path -PathType Leaf)) {
        $path = Get-ChildItem -Path $Collection.Root -Filter 'Inventory.xml' -File -Recurse -Force -ErrorAction SilentlyContinue |
            Where-Object { $_.FullName -match 'OM Server Administrator' -or $_.FullName -match 'logs' } |
            Select-Object -First 1 -ExpandProperty FullName
    }

    if ([string]::IsNullOrWhiteSpace($path) -or -not (Test-Path -LiteralPath $path -PathType Leaf)) {
        return $null
    }

    $raw = [System.IO.File]::ReadAllText($path)
    if ([string]::IsNullOrWhiteSpace($raw)) {
        $raw = [System.IO.File]::ReadAllText($path, [System.Text.Encoding]::Unicode)
    }

    # Some DSET files are UTF-16 and can include null characters when read with the
    # wrong default encoding. Strip nulls as a safety net.
    $raw = $raw -replace [char]0, ''

    $xml = New-Object System.Xml.XmlDocument
    $xml.PreserveWhitespace = $false
    $xml.LoadXml($raw.Trim())

    return $xml
}

function Get-DriFTDsetRawXml {
<#
.SYNOPSIS
    Loads a DSET rawxml/xml file by name.
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Collection,
        [Parameter(Mandatory)][string]$FileName
    )

    $file = Get-ChildItem -Path $Collection.Root -Filter $FileName -File -Recurse -Force -ErrorAction SilentlyContinue |
        Where-Object { $_.FullName -match 'rawxml|gui\\xml|gui/xml' } |
        Select-Object -First 1

    if (-not $file) { return $null }

    try {
        return [xml](Get-Content -Raw -LiteralPath $file.FullName)
    }
    catch {
        try {
            $raw = [System.IO.File]::ReadAllText($file.FullName) -replace [char]0, ''
            return [xml]$raw
        }
        catch {
            return $null
        }
    }
}


function Get-DriFTDsetXmlTextValues {
<#
.SYNOPSIS
    Returns text values from XML nodes whose name matches a pattern.
#>
    [CmdletBinding()]
    param(
        [AllowNull()]$Xml,
        [Parameter(Mandatory)][string]$NamePattern
    )

    if ($null -eq $Xml) { return @() }

    $values = @()

    try {
        foreach ($node in @($Xml.SelectNodes("//*"))) {
            if ($node.LocalName -imatch $NamePattern -or $node.Name -imatch $NamePattern) {
                $text = ([string]$node.InnerText).Trim()
                if (-not [string]::IsNullOrWhiteSpace($text)) {
                    $values += @($text)
                }
            }
        }
    }
    catch { }

    return @($values | Sort-Object -Unique)
}

function Get-DriFTDsetFirstXmlTextValue {
<#
.SYNOPSIS
    Returns the first matching text value from one or more DSET XML files.
#>
    [CmdletBinding()]
    param(
        [AllowNull()][object[]]$Xml,
        [Parameter(Mandatory)][string]$NamePattern
    )

    foreach ($doc in @($Xml | Where-Object { $null -ne $_ })) {
        $value = Get-DriFTFirstNonEmpty (Get-DriFTDsetXmlTextValues -Xml $doc -NamePattern $NamePattern)
        if (-not [string]::IsNullOrWhiteSpace($value)) { return $value }
    }

    return ''
}

function New-DriFTDsetInventoryItemFromText {
<#
.SYNOPSIS
    Creates a best-effort DSET inventory item from text-only XML fallback data.
#>
    [CmdletBinding()]
    param(
        [AllowNull()][string]$ComponentType,
        [AllowNull()][string]$Name,
        [AllowNull()][string]$Version,
        [AllowNull()][string]$ComponentID
    )

    if ([string]::IsNullOrWhiteSpace($Name) -and [string]::IsNullOrWhiteSpace($Version)) {
        return $null
    }

    return New-DriFTInventoryItem `
        -SourceGeneration 'DSET' `
        -ComponentType $ComponentType `
        -ComponentID $ComponentID `
        -Version $Version `
        -Display $Name `
        -ElementName $Name `
        -RelatedItem '' `
        -Source 'DSET XML fallback'
}


function Get-DriFTDsetSystemIdentity {
<#
.SYNOPSIS
    Gets normalized system identity from a DSET report.
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Collection,
        [Parameter(Mandatory)]$Context
    )

    $inventoryXml = Get-DriFTDsetInventoryXml -Collection $Collection
    $syssumXml = Get-DriFTDsetRawXml -Collection $Collection -FileName 'syssum.xml'
    $chasInfoXml = Get-DriFTDsetRawXml -Collection $Collection -FileName 'chasinfo.xml'

    $systemId = Get-DriFTFirstNonEmpty `
        $inventoryXml.SVMInventory.System.systemID `
        (Get-DriFTDsetFirstXmlTextValue -Xml @($syssumXml,$chasInfoXml) -NamePattern 'SystemID|SystemId')

    $model = Get-DriFTFirstNonEmpty `
        $syssumXml.OMA.OMA.ChassisList.Chassis.ChassisInfo.ChassisProps1.ChassModel `
        $syssumXml.OMA.ChassisList.Chassis.ChassisInfo.ChassisProps1.ChassModel `
        (Get-DriFTDsetFirstXmlTextValue -Xml @($syssumXml,$chasInfoXml) -NamePattern 'ChassModel|ChassisModel|Model|SystemModel') `
        ''

    $serviceTag = Get-DriFTFirstNonEmpty `
        $syssumXml.OMA.OMA.ChassisList.Chassis.ChassisInfo.ChassisProps2.ServiceTag `
        $syssumXml.OMA.ChassisList.Chassis.ChassisInfo.ChassisProps2.ServiceTag `
        (Get-DriFTDsetFirstXmlTextValue -Xml @($syssumXml,$chasInfoXml) -NamePattern 'ServiceTag|SvcTag|AssetTag') `
        ''

    $hostName = Get-DriFTFirstNonEmpty `
        $syssumXml.OMA.OMA.ChassisList.Chassis.ChassisInfo.SystemInfo.SystemName `
        $syssumXml.OMA.ChassisList.Chassis.ChassisInfo.SystemInfo.SystemName `
        (Get-DriFTDsetFirstXmlTextValue -Xml @($syssumXml,$chasInfoXml) -NamePattern 'SystemName|HostName|Hostname') `
        ''

    if ([string]::IsNullOrWhiteSpace($serviceTag)) {
        $serviceTag = Get-DriFTFirstNonEmpty ([IO.Path]::GetFileNameWithoutExtension($Collection.SourcePath) -replace '^.*?([A-Z0-9]{7}).*$', '$1')
    }

    $normalizedModel = ConvertTo-DriFTServerModel -Model $model
    if ([string]::IsNullOrWhiteSpace($normalizedModel)) { $normalizedModel = $model }

    $catalogFlags = Get-DriFTCatalogNeed -Model $model -NormalizedModel $normalizedModel

    if ($catalogFlags.S2DCatalogNeeded -eq 'YES') {
        $serviceTag = "$serviceTag***"
    }

    if ($Context.ServiceTags) {
        [void]$Context.ServiceTags.Add(($serviceTag -replace '\*',''))
    }

    Write-DriFTLog -Context $Context -Message "DSET identity: ServiceTag=$serviceTag; Model=$normalizedModel; SystemID=$systemId; Host=$hostName" -Level Info -Indent 1

    return New-DriFTSystemInfo `
        -ServiceTag $serviceTag `
        -PowerEdge $normalizedModel `
        -HostName $hostName `
        -SystemID $systemId `
        -SourceType 'DSET' `
        -SpecialCatalogNeeded $catalogFlags.SpecialCatalogNeeded `
        -S2DCatalogNeeded $catalogFlags.S2DCatalogNeeded
}

function Get-DriFTDsetOperatingSystem {
<#
.SYNOPSIS
    Gets operating system data from DSET Inventory.xml.
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Collection,
        [Parameter(Mandatory)]$Context
    )

    $inventoryXml = Get-DriFTDsetInventoryXml -Collection $Collection
    $osNode = $inventoryXml.SVMInventory.OperatingSystem

    $osSumXml = Get-DriFTDsetRawXml -Collection $Collection -FileName 'OSsum.xml'
    $unameXml = Get-DriFTDsetRawXml -Collection $Collection -FileName 'uname.xml'

    $major = Get-DriFTFirstNonEmpty $osNode.majorVersion (Get-DriFTDsetFirstXmlTextValue -Xml @($osSumXml,$unameXml) -NamePattern 'Major')
    $minor = Get-DriFTFirstNonEmpty $osNode.minorVersion (Get-DriFTDsetFirstXmlTextValue -Xml @($osSumXml,$unameXml) -NamePattern 'Minor')
    $vendor = Get-DriFTFirstNonEmpty $osNode.osVendor (Get-DriFTDsetFirstXmlTextValue -Xml @($osSumXml,$unameXml) -NamePattern 'Vendor|Distributor|Name') 'Microsoft'
    $arch = Get-DriFTFirstNonEmpty $osNode.osArch (Get-DriFTDsetFirstXmlTextValue -Xml @($osSumXml,$unameXml) -NamePattern 'Arch|Architecture') 'x64'

    if ([string]::IsNullOrWhiteSpace($major)) { $major = '6' }
    if ([string]::IsNullOrWhiteSpace($minor)) { $minor = '3' }

    $display = if ($vendor -imatch 'Microsoft' -and $major -eq '6' -and $minor -eq '3') {
        'Microsoft Windows Server 2012 R2'
    }
    elseif ($vendor -imatch 'Microsoft' -and $major -eq '6' -and $minor -eq '2') {
        'Microsoft Windows Server 2012'
    }
    elseif ($vendor -imatch 'Microsoft' -and $major -eq '6' -and $minor -eq '1') {
        'Microsoft Windows Server 2008 R2'
    }
    elseif ($vendor -imatch 'Microsoft') {
        "Microsoft Windows $major.$minor"
    }
    else {
        "$vendor $major.$minor"
    }

    Write-DriFTLog -Context $Context -Message "DSET OS detected: $display ($arch)" -Level Info -Indent 1

    return ConvertTo-DriFTOperatingSystemInfo -OSName $display -OSVersion "$major.$minor"
}

function Import-DriFTDsetInventory {
<#
.SYNOPSIS
    Imports installed inventory from DSET.

.DESCRIPTION
    Uses OM Server Administrator Inventory.xml as the primary source because it
    exposes DSET-era componentID, componentType, version, display, and PCI identity.
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Collection,
        [Parameter(Mandatory)]$Context
    )

    $inventoryXml = Get-DriFTDsetInventoryXml -Collection $Collection

    $items = @()

    if ($inventoryXml) {
        foreach ($device in @($inventoryXml.SVMInventory.Device | Where-Object { $null -ne $_ })) {
        $componentId = Get-DriFTFirstNonEmpty $device.componentID
        $vendorId = Get-DriFTFirstNonEmpty $device.vendorID
        $deviceId = Get-DriFTFirstNonEmpty $device.deviceID
        $subVendorId = Get-DriFTFirstNonEmpty $device.subVendorID
        $subDeviceId = Get-DriFTFirstNonEmpty $device.subDeviceID
        $deviceDisplay = Get-DriFTFirstNonEmpty $device.display

        foreach ($app in @($device.Application | Where-Object { $null -ne $_ })) {
            $componentType = Get-DriFTFirstNonEmpty $app.componentType $app.componenttype
            $version = Get-DriFTFirstNonEmpty $app.version
            $display = Get-DriFTFirstNonEmpty $app.display $deviceDisplay

            $item = New-DriFTInventoryItem `
                -SourceGeneration 'DSET' `
                -ComponentType $componentType `
                -ComponentID $componentId `
                -VendorID $vendorId `
                -DeviceID $deviceId `
                -SubVendorID $subVendorId `
                -SubDeviceID $subDeviceId `
                -Version $version `
                -Display $deviceDisplay `
                -ElementName $display `
                -RelatedItem '' `
                -Source 'DSET Inventory.xml'

            if ($null -ne $item -and ($item.ComponentID -or $item.DeviceID -or $item.ElementName -or $item.Display)) {
                $items += @($item)
            }
        }
        }
    }
    else {
        Write-DriFTLog -Context $Context -Message 'DSET Inventory.xml was not found. Using legacy DSET XML fallback parser.' -Level Warn -Indent 1
    }

    if (@($items).Count -eq 0) {
        $fwViewXml = Get-DriFTDsetRawXml -Collection $Collection -FileName 'fwview.xml'
        $biosViewXml = Get-DriFTDsetRawXml -Collection $Collection -FileName 'biosview.xml'
        $driverListXml = Get-DriFTDsetRawXml -Collection $Collection -FileName 'driverlist.xml'
        if (-not $driverListXml) { $driverListXml = Get-DriFTDsetRawXml -Collection $Collection -FileName 'getdriverlist.xml' }

        $biosVersion = Get-DriFTDsetFirstXmlTextValue -Xml @($biosViewXml,$fwViewXml) -NamePattern 'Version|BIOSVersion'
        if (-not [string]::IsNullOrWhiteSpace($biosVersion)) {
            $items += @(New-DriFTInventoryItem `
                -SourceGeneration 'DSET' `
                -ComponentType 'BIOS' `
                -ComponentID '159' `
                -Version $biosVersion `
                -Display 'BIOS' `
                -ElementName 'BIOS' `
                -Source 'DSET BIOS/FW fallback')
        }

        # Best-effort generic firmware/driver extraction. This will not be perfect
        # for every DSET schema, but it lets catalog matching work when componentID
        # is present and keeps older reports from returning no inventory.
        foreach ($docInfo in @(
            [PSCustomObject]@{ Xml = $fwViewXml;     Type = 'FRMW' },
            [PSCustomObject]@{ Xml = $driverListXml; Type = 'DRVR' }
        )) {
            if (-not $docInfo.Xml) { continue }

            foreach ($node in @($docInfo.Xml.SelectNodes('//*') | Where-Object { $_.ChildNodes.Count -gt 0 })) {
                $name = Get-DriFTFirstNonEmpty `
                    (Get-DriFTDsetFirstXmlTextValue -Xml @($node) -NamePattern 'Name|Display|Description|Device')
                $version = Get-DriFTFirstNonEmpty `
                    (Get-DriFTDsetFirstXmlTextValue -Xml @($node) -NamePattern 'Version|FirmwareVersion|DriverVersion')
                $componentId = Get-DriFTFirstNonEmpty `
                    (Get-DriFTDsetFirstXmlTextValue -Xml @($node) -NamePattern 'ComponentID|ComponentId')

                if (-not [string]::IsNullOrWhiteSpace($name) -or -not [string]::IsNullOrWhiteSpace($version)) {
                    $fallbackItem = New-DriFTDsetInventoryItemFromText `
                        -ComponentType $docInfo.Type `
                        -Name $name `
                        -Version $version `
                        -ComponentID $componentId

                    if ($fallbackItem) { $items += @($fallbackItem) }
                }
            }
        }
    }

    # Fallback for older DSET reports where BIOS/FWView entries are present but
    # missing from Inventory.xml.
    if (-not (@($items | Where-Object { $_.ComponentID -eq '159' -or $_.ComponentType -eq 'BIOS' }).Count)) {
        $biosXml = Get-DriFTDsetRawXml -Collection $Collection -FileName 'BIOSView.xml'
        $biosVersion = Get-DriFTFirstNonEmpty $biosXml.OMA.BIOSView1.SystemBIOS.Version
        if (-not [string]::IsNullOrWhiteSpace($biosVersion)) {
            $items += @(New-DriFTInventoryItem `
                -SourceGeneration 'DSET' `
                -ComponentType 'BIOS' `
                -ComponentID '159' `
                -Version $biosVersion `
                -Display 'BIOS' `
                -ElementName 'BIOS' `
                -Source 'DSET BIOSView.xml')
        }
    }

    $items = @($items |
        Where-Object { $null -ne $_ } |
        Sort-Object ComponentType,ComponentID,VendorID,DeviceID,SubVendorID,SubDeviceID,Version,Display,ElementName -Unique)

    Write-DriFTLog -Context $Context -Message "DSET inventory rows found: $(@($items).Count)" -Level Info -Indent 1

    return @($items)
}

#endregion DSET Parser


#region Legacy TSR Parser

function Get-DriFTLegacyTsrXml {
<#
.SYNOPSIS
    Loads a legacy TSR inventory XML file.

.DESCRIPTION
    Centralizes XML loading so parser functions do not duplicate Get-Content/[xml] logic.
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Collection,
        [Parameter(Mandatory)][string]$FileName
    )

    $path = Join-Path $Collection.InventoryPath $FileName
    if (-not (Test-Path $path -PathType Leaf)) { return $null }

    return [xml](Get-Content -Raw -Path $path)
}

function ConvertFrom-DriFTCimNamedInstances {
<#
.SYNOPSIS
    Converts CIM VALUE.NAMEDINSTANCE XML nodes to easier PowerShell objects.

.DESCRIPTION
    Legacy TSR XML stores data as generic Property nodes. This function flattens
    the properties into note properties for faster and easier lookups.
#>
    [CmdletBinding()]
    param([AllowNull()]$NamedInstances)

    $rows = @()

    foreach ($instance in @($NamedInstances.INSTANCE | Where-Object { $null -ne $_ })) {
        $obj = [ordered]@{}
        foreach ($prop in @($instance.Property | Where-Object { $null -ne $_ })) {
            if ($prop.Name) { $obj[$prop.Name] = $prop.InnerText }
        }

        if ($obj.Count -gt 0) {
            $rows += @([PSCustomObject]$obj)
        }
    }

    return @($rows)
}

function Get-DriFTLegacyTsrSystemIdentity {
<#
.SYNOPSIS
    Gets system identity from 16G and older TSR XML.

.DESCRIPTION
    Ports the existing DCIM_SystemView and BIOSAttribute logic into a contained function.
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Collection,
        [Parameter(Mandatory)]$Context
    )

    $biosXml = Get-DriFTLegacyTsrXml -Collection $Collection -FileName 'sysinfo_CIM_BIOSAttribute.xml'
    $viewXml = Get-DriFTLegacyTsrXml -Collection $Collection -FileName 'sysinfo_DCIM_View.xml'

    if (-not $biosXml -or -not $viewXml) {
        throw 'Legacy TSR identity XML files were not found.'
    }

    $biosInstances = $biosXml.CIM.MESSAGE.SIMPLEREQ.'VALUE.NAMEDINSTANCE'.INSTANCE
    $viewInstances = $viewXml.CIM.MESSAGE.SIMPLEREQ.'VALUE.NAMEDINSTANCE'.INSTANCE

    $systemView = $viewInstances | Where-Object { $_.CLASSNAME -eq 'DCIM_SystemView' } | Select-Object -First 1

    $serviceTag = Get-DriFTCimPropertyValue -Instance $systemView -Name 'ServiceTag'
    $model = Get-DriFTCimPropertyValue -Instance $systemView -Name 'Model'

    $sysIdNode = $biosInstances |
        Where-Object { $_.CLASSNAME -eq 'DCIM_LCString' -and $_.PROPERTY.VALUE -eq 'SYSID' } |
        Select-Object -First 1

    $systemId = Get-DriFTCimArrayCurrentValue -Instance $sysIdNode

    $hostNode = $biosInstances |
        Where-Object { $_.CLASSNAME -eq 'DCIM_SystemString' -and $_.PROPERTY.VALUE -eq 'HostName' } |
        Select-Object -First 1

    $hostName = Get-DriFTCimArrayCurrentValue -Instance $hostNode
    $normalizedModel = ConvertTo-DriFTServerModel -Model $model

    $catalogFlags = Get-DriFTCatalogNeed -Model $model -NormalizedModel $normalizedModel

    if ($catalogFlags.S2DCatalogNeeded -eq 'YES') {
        $serviceTag = "$serviceTag***"
    }

    [void]$Context.ServiceTags.Add(($serviceTag -replace '\*',''))

    return New-DriFTSystemInfo `
        -ServiceTag $serviceTag `
        -PowerEdge $normalizedModel `
        -HostName $hostName `
        -SystemID $systemId `
        -SourceType 'LegacyTSR' `
        -SpecialCatalogNeeded $catalogFlags.SpecialCatalogNeeded `
        -S2DCatalogNeeded $catalogFlags.S2DCatalogNeeded
}


function Get-DriFTLegacyTsrSystemStringValues {
<#
.SYNOPSIS
    Gets all CurrentValue entries for a legacy TSR DCIM_SystemString name.

.DESCRIPTION
    VMware TSRs can contain multiple OSVersion CurrentValue values. The old DriFT
    logic inspected all of them and selected the value that contained the actual
    ESXi release, such as 8.0 U3. The rewrite previously took only the first value,
    which could be Dell-ESXi and caused Broadcom Compatibility Guide lookups to fail.
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Instances,
        [Parameter(Mandatory)][string]$Name
    )

    $values = @()

    $nodes = @($Instances |
        Where-Object {
            $_.CLASSNAME -eq 'DCIM_SystemString' -and
            $_.PROPERTY.Value -match $Name
        })

    foreach ($node in $nodes) {
        foreach ($currentNode in @($node.ChildNodes | Where-Object { $_.Name -match 'CurrentValue' })) {
            foreach ($value in @($currentNode.InnerText)) {
                if (-not [string]::IsNullOrWhiteSpace([string]$value)) {
                    $values += @(([string]$value).Trim())
                }
            }
        }

        # Keep fallback for XML shapes where CurrentValue is exposed as PROPERTY.ARRAY.
        $fallback = Get-DriFTCimArrayCurrentValue -Instance $node
        if (-not [string]::IsNullOrWhiteSpace([string]$fallback)) {
            $values += @(([string]$fallback).Trim())
        }
    }

    return @($values | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique)
}

function Select-DriFTBestVmwareOsVersion {
<#
.SYNOPSIS
    Selects the actual ESXi/vSAN release from legacy TSR OSVersion values.
#>
    [CmdletBinding()]
    param([AllowEmptyCollection()][object[]]$Values)

    foreach ($value in @($Values)) {
        $parsed = Get-DriFTEsxiVersionFromText -Text ([string]$value)
        if (-not [string]::IsNullOrWhiteSpace($parsed)) {
            return $parsed
        }
    }

    foreach ($value in @($Values)) {
        if ([string]$value -imatch 'build|patch|GA|Update|U[123]') {
            $parsed = ConvertTo-DriFTVmwareVersion -OSName ([string]$value) -OSVersion ''
            if (-not [string]::IsNullOrWhiteSpace($parsed)) {
                return $parsed
            }
        }
    }

    return Get-DriFTFirstNonEmpty $Values
}


function Get-DriFTLegacyTsrOperatingSystem {
<#
.SYNOPSIS
    Gets operating system data from 16G and older TSR XML.

.DESCRIPTION
    Reads all legacy DCIM_SystemString OSName/OSVersion CurrentValue values. This
    preserves old DriFT VMware behavior where OSVersion can include both Dell image
    profile text and the real ESXi release.
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Collection,
        [Parameter(Mandatory)]$Context
    )

    $biosXml = Get-DriFTLegacyTsrXml -Collection $Collection -FileName 'sysinfo_CIM_BIOSAttribute.xml'
    if (-not $biosXml) {
        return New-DriFTOperatingSystemInfo -RawName '' -Family 'Windows' -DisplayName 'NO OS Detected in TSR Data: Assuming Windows 64bit' -Version '' -MajorVersion 6 -MinorVersion 3
    }

    $instances = $biosXml.CIM.MESSAGE.SIMPLEREQ.'VALUE.NAMEDINSTANCE'.INSTANCE

    $osNameValues = @(Get-DriFTLegacyTsrSystemStringValues -Instances $instances -Name 'OSName')
    $osVersionValues = @(Get-DriFTLegacyTsrSystemStringValues -Instances $instances -Name 'OSVersion')

    $osName = Get-DriFTFirstNonEmpty $osNameValues
    $osVersion = Get-DriFTFirstNonEmpty $osVersionValues

    if ("$osName $($osVersionValues -join ' ')" -imatch 'VMware|ESXi|vSAN') {
        $bestVmwareVersion = Select-DriFTBestVmwareOsVersion -Values $osVersionValues

        Write-DriFTLog -Context $Context -Message "VMware OS detected. OSName='$osName'; OSVersion candidates='$($osVersionValues -join ' | ')'; selected='$bestVmwareVersion'" -Level Info -Indent 1

        return ConvertTo-DriFTOperatingSystemInfo -OSName $osName -OSVersion $bestVmwareVersion
    }

    return ConvertTo-DriFTOperatingSystemInfo -OSName $osName -OSVersion $osVersion
}

function Import-DriFTLegacyTsrInventory {
<#
.SYNOPSIS
    Imports installed hardware from legacy DCIM_SoftwareIdentity XML.

.DESCRIPTION
    Preserves the 16G-and-older inventory source and normalizes rows into
    New-DriFTInventoryItem. Uses plain PowerShell arrays instead of generic lists
    to avoid Windows PowerShell 5.1 XML-node type conversion issues.
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Collection,
        [Parameter(Mandatory)]$Context
    )

    $xml = Get-DriFTLegacyTsrXml -Collection $Collection -FileName 'sysinfo_DCIM_SoftwareIdentity.xml'
    if (-not $xml) {
        Write-DriFTLog -Context $Context -Message 'Legacy SoftwareIdentity XML was not found.' -Level Warn -Indent 1
        return @()
    }

    $namedInstances = @($xml.CIM.MESSAGE.SIMPLEREQ.'VALUE.NAMEDINSTANCE')

    $installed = @($namedInstances | Where-Object {
        try {
            $_.INSTANCENAME.KEYBINDING.KEYVALUE.'#text' -match 'DCIM:INSTALLED'
        }
        catch { $false }
    })

    $items = @()

    foreach ($prop in @($installed.INSTANCE | Where-Object { $null -ne $_ })) {
        $item = New-DriFTInventoryItem `
            -SourceGeneration 'Legacy' `
            -ComponentType (Get-DriFTCimPropertyValue -Instance $prop -Name 'ComponentType') `
            -ComponentID (Get-DriFTCimPropertyValue -Instance $prop -Name 'ComponentID') `
            -VendorID (Get-DriFTCimPropertyValue -Instance $prop -Name 'VendorID') `
            -DeviceID (Get-DriFTCimPropertyValue -Instance $prop -Name 'DeviceID') `
            -SubVendorID (Get-DriFTCimPropertyValue -Instance $prop -Name 'SubVendorID') `
            -SubDeviceID (Get-DriFTCimPropertyValue -Instance $prop -Name 'SubDeviceID') `
            -Version (Get-DriFTCimPropertyValue -Instance $prop -Name 'VersionString') `
            -Display (Get-DriFTCimPropertyValue -Instance $prop -Name 'FQDD') `
            -ElementName (Get-DriFTCimPropertyValue -Instance $prop -Name 'ElementName') `
            -RelatedItem '' `
            -Source 'DCIM_SoftwareIdentity'

        if ($null -ne $item -and ($item.ComponentID -or $item.DeviceID -or $item.ElementName -or $item.Display)) {
            $items += @($item)
        }
    }

    Write-DriFTLog -Context $Context -Message "Legacy installed SoftwareIdentity rows found: $(@($items).Count)" -Level Info -Indent 1

    return @($items |
        Where-Object { $null -ne $_ } |
        Sort-Object ComponentType,ComponentID,VendorID,DeviceID,SubDeviceID,SubVendorID,Version,Display -Unique)
}

#endregion Legacy TSR Parser

#region 17G Redfish Parser

function Import-DriFT17GInventory {
<#
.SYNOPSIS
    Imports 17G inventory from viewer.html or redfishidracwalk.tar.gz.

.DESCRIPTION
    17G SupportAssist collections may omit legacy SoftwareIdentity XML. This parser
    reads normalized Redfish data, builds firmware inventory, builds PCI identity maps,
    then enriches firmware rows with catalog-ready PCI IDs.
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Collection,
        [Parameter(Mandatory)]$Context
    )

    $redfishRoot = if ($Collection.ViewerHtmlPath) {
        Expand-DriFT17GViewerHtmlRedfishWalk -ViewerHtmlPath $Collection.ViewerHtmlPath -Context $Context
    }
    elseif ($Collection.RedfishWalkPath) {
        Expand-DriFT17GRedfishWalk -TarGzPath $Collection.RedfishWalkPath -Context $Context
    }
    else {
        throw '17G collection did not include viewer.html or redfishidracwalk.tar.gz.'
    }

    $allJsonFiles = @(Get-ChildItem -Path $redfishRoot -Filter '*.json' -File -Recurse -Force -ErrorAction SilentlyContinue)
    Export-DriFTDebugData -Context $Context -Name '17G_AllJsonFiles.csv' -InputObject ($allJsonFiles | Select-Object FullName, Length)

    $firmwareRows = @(Get-DriFT17GFirmwareRows -RedfishRoot $redfishRoot -AllJsonFiles $allJsonFiles -Context $Context)
    Write-DriFTLog -Context $Context -Message "17G firmware inventory rows found: $(@($firmwareRows).Count)" -Level Info -Indent 1

    $pciRows = @(Get-DriFT17GPciIdentityRows -AllJsonFiles $allJsonFiles -Context $Context)
    Write-DriFTLog -Context $Context -Message "17G PCI identity rows found: $(@($pciRows).Count)" -Level Info -Indent 1

    $enriched = @(Add-DriFT17GPciIdentityToFirmwareRows -FirmwareRows $firmwareRows -PciRows $pciRows -Context $Context)
    Write-DriFTLog -Context $Context -Message "17G enriched inventory rows returned: $(@($enriched).Count)" -Level Info -Indent 1

    Export-DriFTDebugData -Context $Context -Name '17G_PciIdentity.csv' -InputObject $pciRows
    Export-DriFTDebugData -Context $Context -Name '17G_EnrichedFirmware.csv' -InputObject $enriched

    return @($enriched |
        Sort-Object ComponentType,ComponentID,VendorID,DeviceID,SubVendorID,SubDeviceID,Version,Display -Unique)
}

function Get-DriFT17GSystemIdentity {
<#
.SYNOPSIS
    Gets 17G system identity from metadata.json and/or Redfish System object.

.DESCRIPTION
    Uses metadata.json first, then Redfish System.Embedded.1 fallback.
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Collection,
        [Parameter(Mandatory)]$Context
    )

    $metadata = Get-DriFTMetadataJson -Root $Collection.Root

    $redfishRoot = $null
    try {
        if ($Collection.ViewerHtmlPath) {
            $redfishRoot = Expand-DriFT17GViewerHtmlRedfishWalk -ViewerHtmlPath $Collection.ViewerHtmlPath -Context $Context
        }
        elseif ($Collection.RedfishWalkPath) {
            $redfishRoot = Expand-DriFT17GRedfishWalk -TarGzPath $Collection.RedfishWalkPath -Context $Context
        }
    }
    catch {
        Write-DriFTLog -Context $Context -Message "Unable to expand 17G Redfish for identity: $($_.Exception.Message)" -Level Warn -Indent 1
    }

    $systemJson = if ($redfishRoot) {
        Get-DriFT17GJsonFile -RedfishRoot $redfishRoot -RelativePath 'redfish/v1/Systems/System.Embedded.1/index.json'
    }

    $model = Get-DriFTFirstNonEmpty (Get-DriFTObjectProperty -InputObject $metadata -PropertyName 'Model') $systemJson.Model
    $normalizedModel = ConvertTo-DriFTServerModel -Model $model
    $catalogFlags = Get-DriFTCatalogNeed -Model $model -NormalizedModel $normalizedModel

    $serviceTag = Get-DriFTFirstNonEmpty (Get-DriFTObjectProperty -InputObject $metadata -PropertyName 'ServiceTag') $systemJson.SerialNumber $systemJson.SKU
    if ($catalogFlags.S2DCatalogNeeded -eq 'YES') {
        $serviceTag = "$serviceTag***"
    }

    [void]$Context.ServiceTags.Add(($serviceTag -replace '\*',''))

    return New-DriFTSystemInfo `
        -ServiceTag $serviceTag `
        -PowerEdge $normalizedModel `
        -HostName (Get-DriFTFirstNonEmpty (Get-DriFTObjectProperty -InputObject $metadata -PropertyName 'HostName') $systemJson.HostName $systemJson.Name) `
        -SystemID (Get-DriFTFirstNonEmpty (Get-DriFTObjectProperty -InputObject $metadata -PropertyName 'DeviceSystemId') (Get-DriFTObjectProperty -InputObject $metadata -PropertyName 'SystemID') $systemJson.Oem.Dell.DellSystem.SystemID) `
        -SourceType 'TSR17G' `
        -SpecialCatalogNeeded $catalogFlags.SpecialCatalogNeeded `
        -S2DCatalogNeeded $catalogFlags.S2DCatalogNeeded
}

function Get-DriFT17GOperatingSystem {

    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Collection,
        [Parameter(Mandatory)]$Context
    )

    $metadata = Get-DriFTMetadataJson -Root $Collection.Root

    $osName = Get-DriFTFirstNonEmpty `
        (Get-DriFTObjectProperty -InputObject $metadata -PropertyName 'OSName') `
        (Get-DriFTObjectProperty -InputObject $metadata -PropertyName 'OperatingSystem')

    #
    # IMPORTANT:
    # If 17G metadata.json does not contain OS information,
    # suppress all OS-dependent matching (drivers, VMware, Microsoft Update).
    #
    if ([string]::IsNullOrWhiteSpace($osName)) {

        Write-DriFTLog `
            -Context $Context `
            -Message '17G metadata.json did not contain OS information. Driver and OS update matching will be skipped.' `
            -Level Warn `
            -Indent 1

        return New-DriFTOperatingSystemInfo `
            -RawName '' `
            -Family '' `
            -DisplayName 'NO OS Detected in TSR Data' `
            -Version '' `
            -CatalogPackageType 'LW64' `
            -MajorVersion 0 `
            -MinorVersion 0 `
            -DriverSupport $false
    }

    return ConvertTo-DriFTOperatingSystemInfo `
        -OSName $osName `
        -OSVersion ''
}

function Expand-DriFT17GRedfishWalk {
<#
.SYNOPSIS
    Extracts redfishidracwalk.tar.gz to a short temp path.

.DESCRIPTION
    Uses a short extraction path to reduce MAX_PATH issues in Windows PowerShell 5.1.
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$TarGzPath,
        [Parameter(Mandatory)]$Context
    )

    $dest = Join-Path $Context.RedfishRoot ([guid]::NewGuid().Guid.Substring(0, 8))
    New-Item -ItemType Directory -Force -Path $dest | Out-Null

    $tar = Get-Command tar.exe -ErrorAction SilentlyContinue
    if (-not $tar) { $tar = Get-Command tar -ErrorAction SilentlyContinue }
    if (-not $tar) { throw 'tar.exe was not found. Cannot extract redfishidracwalk.tar.gz.' }

    & $tar.Source -xzf $TarGzPath -C $dest 2>$null
    if ($LASTEXITCODE -ne 0) {
        throw "tar extraction failed for $TarGzPath"
    }

    return $dest
}

function Expand-DriFT17GViewerHtmlRedfishWalk {
<#
.SYNOPSIS
    Extracts embedded Redfish ZIP from 17G viewer.html.

.DESCRIPTION
    Handles invalid Windows path characters in ZIP entries by sanitizing extracted paths.
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ViewerHtmlPath,
        [Parameter(Mandatory)]$Context
    )

    $dest = Join-Path $Context.RedfishRoot ([guid]::NewGuid().Guid.Substring(0, 8))
    New-Item -ItemType Directory -Force -Path $dest | Out-Null

    $raw = Get-Content -Raw -Path $ViewerHtmlPath
    $match = [regex]::Match(
        $raw,
        '<script\s+[^>]*content=["'']redfish["''][^>]*>(?<body>.*?)</script>',
        [System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [System.Text.RegularExpressions.RegexOptions]::Singleline
    )

    if (-not $match.Success) {
        throw "viewer.html did not contain a script tag with content='redfish'."
    }

    $zipPath = Join-Path $dest 'viewer_redfishwalk.zip'
    [System.IO.File]::WriteAllBytes($zipPath, [Convert]::FromBase64String($match.Groups['body'].Value.Trim()))

    Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue

    $invalidCharsPattern = '[<>:"|?*]'
    $zip = [System.IO.Compression.ZipFile]::OpenRead($zipPath)

    try {
        foreach ($entry in $zip.Entries) {
            if ([string]::IsNullOrWhiteSpace($entry.FullName)) { continue }

            $safeRelativePath = ($entry.FullName -replace '/', [IO.Path]::DirectorySeparatorChar) -replace $invalidCharsPattern, '_'
            $safeRelativePath = $safeRelativePath.TrimStart([IO.Path]::DirectorySeparatorChar)
            if ([string]::IsNullOrWhiteSpace($safeRelativePath)) { continue }

            $destinationPath = Join-Path $dest $safeRelativePath

            if ($entry.FullName.EndsWith('/') -or [string]::IsNullOrWhiteSpace($entry.Name)) {
                New-Item -ItemType Directory -Force -Path $destinationPath | Out-Null
                continue
            }

            $destinationDirectory = Split-Path -Path $destinationPath -Parent
            New-Item -ItemType Directory -Force -Path $destinationDirectory | Out-Null

            [System.IO.Compression.ZipFileExtensions]::ExtractToFile($entry, $destinationPath, $true)
        }
    }
    finally {
        $zip.Dispose()
    }

    return $dest
}

function Get-DriFT17GJsonFile {
<#
.SYNOPSIS
    Finds and loads a Redfish JSON file.

.DESCRIPTION
    Tries the expected Redfish path first, then falls back to suffix search.
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$RedfishRoot,
        [Parameter(Mandatory)][string]$RelativePath
    )

    $cleanPath = $RelativePath.TrimStart('/').Replace('/', [IO.Path]::DirectorySeparatorChar)
    $jsonPath = Join-Path $RedfishRoot $cleanPath

    if (-not (Test-Path $jsonPath -PathType Leaf)) {
        $suffix = [regex]::Escape($RelativePath.TrimStart('/').Replace('/', [IO.Path]::DirectorySeparatorChar))
        $jsonPath = Get-ChildItem -Path $RedfishRoot -Filter 'index.json' -File -Recurse -Force -ErrorAction SilentlyContinue |
            Where-Object { $_.FullName -match "$suffix$" } |
            Select-Object -First 1 -ExpandProperty FullName
    }

    if ($jsonPath -and (Test-Path $jsonPath -PathType Leaf)) {
        try { return Get-Content -Raw -Path $jsonPath | ConvertFrom-Json }
        catch { return $null }
    }

    return $null
}

function Get-DriFT17GFirmwareRows {
<#
.SYNOPSIS
    Builds normalized firmware inventory rows from 17G Redfish.

.DESCRIPTION
    Finds FirmwareInventory members using the collection index when present and a
    recursive fallback when not present.
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$RedfishRoot,
        [Parameter(Mandatory)][object[]]$AllJsonFiles,
        [Parameter(Mandatory)]$Context
    )

    $fwIndex = Get-DriFT17GJsonFile -RedfishRoot $RedfishRoot -RelativePath 'redfish/v1/UpdateService/FirmwareInventory/index.json'
    $firmwareFiles = @()

    if ($fwIndex -and $fwIndex.Members) {
        foreach ($member in @($fwIndex.Members)) {
            $memberPath = [string]$member.'@odata.id'
            if ([string]::IsNullOrWhiteSpace($memberPath)) { continue }
            $memberPath = $memberPath.TrimStart('/')
            if ($memberPath -notmatch 'index\.json$') { $memberPath = $memberPath.TrimEnd('/') + '/index.json' }

            $candidate = Join-Path $RedfishRoot ($memberPath.Replace('/', [IO.Path]::DirectorySeparatorChar))
            if (Test-Path $candidate -PathType Leaf) {
                $firmwareFiles += @(Get-Item $candidate)
            }
        }
    }

    if (@($firmwareFiles).Count -eq 0) {
        $AllJsonFiles |
            Where-Object {
                $_.FullName -imatch 'UpdateService.*FirmwareInventory' -and
                $_.DirectoryName -notmatch '[\\/]FirmwareInventory$'
            } |
            ForEach-Object { $firmwareFiles += @($_) }
    }

    if (@($firmwareFiles).Count -eq 0) {
        foreach ($jsonFile in $AllJsonFiles) {
            try {
                $raw = Get-Content -Raw -Path $jsonFile.FullName
                if (($raw -imatch '"@odata\.type"\s*:\s*".*SoftwareInventory') -or
                    ($raw -imatch '"SoftwareId"\s*:') -or
                    ($raw -imatch '"Updateable"\s*:')) {
                    $firmwareFiles += @($jsonFile)
                }
            }
            catch {}
        }
    }

    $rows = New-Object System.Collections.Generic.List[object]

    foreach ($file in $firmwareFiles) {
        try { $fw = Get-Content -Raw -Path $file.FullName | ConvertFrom-Json }
        catch { continue }

        if (-not $fw) { continue }

        $fwIdText = [string]$fw.Id
        $fwOdataText = [string]$fw.'@odata.id'
        $fwFileText = [string]$file.FullName

        # 17G viewer/Redfish data can include previous firmware inventory entries.
        # Legacy DriFT reports current installed versions only, so skip previous entries.
        if ($fwIdText -imatch '^Previous-|DCIM[_:]PREVIOUS|PREVIOUS#') { continue }
        if ($fwOdataText -imatch '/Previous-|DCIM[_:]PREVIOUS|PREVIOUS#') { continue }
        if ($fwFileText -imatch 'Previous-|DCIM[_:]PREVIOUS|PREVIOUS#') { continue }

        if ($fw.Members -and -not $fw.Version -and -not $fw.SoftwareId) { continue }

        $odataType = [string]$fw.'@odata.type'
        if (($odataType -and $odataType -notmatch 'SoftwareInventory|FirmwareInventory') -and
            (-not $fw.SoftwareId) -and (-not $fw.Updateable)) {
            continue
        }

        $relatedPath = $null
        if ($fw.RelatedItem) {
            $relatedPath = @($fw.RelatedItem | ForEach-Object { $_.'@odata.id' } | Where-Object { $_ })[0]
        }

        $display = if ($relatedPath) { (($relatedPath.TrimEnd('/') -split '/')[-1]) } else { Get-DriFTFirstNonEmpty $fw.Id $fw.Name }
        $elementName = ([string](Get-DriFTFirstNonEmpty $fw.Name $display)) -replace '\s+Firmware Inventory$', ''
        $componentId = Get-DriFTFirstNonEmpty $fw.SoftwareId $fw.Id

        $item = New-DriFTInventoryItem `
            -SourceGeneration '17G' `
            -ComponentType 'FRMW' `
            -ComponentID ([string]$componentId) `
            -VendorID '' `
            -DeviceID '' `
            -SubVendorID '' `
            -SubDeviceID '' `
            -Version ([string]$fw.Version) `
            -Display ([string]$display) `
            -ElementName ([string]$elementName) `
            -RelatedItem ([string]$relatedPath) `
            -Source 'RedfishFirmwareInventory'

        if ($item.ComponentID -or $item.Version) {
            $rows += @($item)
        }
    }

    return @($rows)
}

function Get-DriFT17GPciIdentityRows {
<#
.SYNOPSIS
    Builds the 17G PCI identity map.

.DESCRIPTION
    Searches PCIeFunction, Network, Storage, and DellSoftwareInventory records.
    DellSoftwareInventory IdentityInfoValue is preferred because it contains catalog-ready
    hex IDs such as VendorID:DeviceID:SubVendorID:SubDeviceID.
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object[]]$AllJsonFiles,
        [Parameter(Mandatory)]$Context
    )

    $rows = @()

    foreach ($jsonFile in $AllJsonFiles) {
        try { $json = Get-Content -Raw -Path $jsonFile.FullName | ConvertFrom-Json }
        catch { continue }

        if (-not $json) { continue }

        $odataType = [string]$json.'@odata.type'
        $pathText = [string]$jsonFile.FullName

        $dellNic = $json.Oem.Dell.DellNIC
        $dellPcieFunction = $json.Oem.Dell.DellPCIeFunction

        $vendorId = Get-DriFTFirstNonEmpty `
            $json.VendorId $json.VendorID $json.PciVendorId $json.PCIVendorID $json.PCIVendorId $json.PCIVendorID `
            $json.Oem.Dell.VendorId $json.Oem.Dell.VendorID $json.Oem.Dell.PCIVendorID $json.Oem.Dell.PCIVendorId `
            $dellNic.PCIVendorID $dellNic.PCIVendorId $dellPcieFunction.PCIVendorID $dellPcieFunction.PCIVendorId

        $deviceId = Get-DriFTFirstNonEmpty `
            $json.DeviceId $json.DeviceID $json.PciDeviceId $json.PCIDeviceID $json.PCIDeviceId $json.PCIDeviceID `
            $json.Oem.Dell.DeviceId $json.Oem.Dell.DeviceID $json.Oem.Dell.PCIDeviceID $json.Oem.Dell.PCIDeviceId `
            $dellNic.PCIDeviceID $dellNic.PCIDeviceId $dellPcieFunction.PCIDeviceID $dellPcieFunction.PCIDeviceId

        $subVendorId = Get-DriFTFirstNonEmpty `
            $json.SubsystemVendorId $json.SubsystemVendorID $json.SubVendorId $json.SubVendorID $json.PciSubVendorId $json.PCISubVendorID $json.PCISubVendorId $json.PCISubVendorID `
            $json.Oem.Dell.SubsystemVendorId $json.Oem.Dell.SubsystemVendorID $json.Oem.Dell.SubVendorID $json.Oem.Dell.PCISubVendorID $json.Oem.Dell.PCISubVendorId `
            $dellNic.PCISubVendorID $dellNic.PCISubVendorId $dellPcieFunction.PCISubVendorID $dellPcieFunction.PCISubVendorId

        $subDeviceId = Get-DriFTFirstNonEmpty `
            $json.SubsystemId $json.SubsystemID $json.SubsystemDeviceId $json.SubsystemDeviceID $json.SubDeviceId $json.SubDeviceID $json.PciSubDeviceId $json.PCISubDeviceID $json.PCISubDeviceId $json.PCISubDeviceID `
            $json.Oem.Dell.SubsystemId $json.Oem.Dell.SubsystemID $json.Oem.Dell.SubDeviceID $json.Oem.Dell.PCISubDeviceID $json.Oem.Dell.PCISubDeviceId `
            $dellNic.PCISubDeviceID $dellNic.PCISubDeviceId $dellPcieFunction.PCISubDeviceID $dellPcieFunction.PCISubDeviceId

        $identityMatchFound = $false
        if ($json.IdentityInfoType -and $json.IdentityInfoValue) {
            $identityTypes = @($json.IdentityInfoType)
            $identityValues = @($json.IdentityInfoValue)

            for ($i = 0; $i -lt @($identityTypes).Count; $i++) {
                $typeText = [string]$identityTypes[$i]
                $valueText = [string]$identityValues[[Math]::Min($i, @($identityValues).Count - 1)]

                if ($typeText -imatch 'VendorID:DeviceID:SubVendorID:SubDeviceID' -and
                    -not [string]::IsNullOrWhiteSpace($valueText)) {

                    $typeParts = $typeText -split ':'
                    $valueParts = $valueText -split ':'
                    $map = @{}

                    for ($p = 0; $p -lt @($typeParts).Count -and $p -lt @($valueParts).Count; $p++) {
                        $map[$typeParts[$p]] = $valueParts[$p]
                    }

                    $vendorId = $map['VendorID']
                    $deviceId = $map['DeviceID']
                    $subVendorId = $map['SubVendorID']
                    $subDeviceId = $map['SubDeviceID']
                    $identityMatchFound = $true
                    break
                }
            }
        }

        if (-not $vendorId -and -not $deviceId -and -not $subVendorId -and -not $subDeviceId) { continue }

        if ((-not $identityMatchFound) -and
            ($odataType -and $odataType -notmatch 'PCIeFunction|PCIeDevice|DellPCIeFunction|NetworkDeviceFunction|NetworkAdapter|Storage|SoftwareInventory') -and
            ($pathText -notmatch 'PCIe|DellPCIe|NetworkAdapters|NetworkDeviceFunctions|Storage|DellSoftwareInventory')) {
            continue
        }

        $keys = Get-DriFT17GObjectKeys -JsonObject $json -FilePath $jsonFile.FullName
        $fqdd = Get-DriFT17GPreferredFqdd -JsonObject $json -FilePath $jsonFile.FullName

        $rows += @([PSCustomObject]@{
            OdataId     = [string]$json.'@odata.id'
            Id          = [string]$json.Id
            Name        = [string]$json.Name
            Description = [string](Get-DriFTFirstNonEmpty `
                $json.ElementName `
                $json.Description `
                $json.DeviceDescription `
                $json.Model `
                $json.ProductName `
                $json.Oem.Dell.ElementName `
                $json.Oem.Dell.Description `
                $json.Oem.Dell.DeviceDescription `
                $json.Oem.Dell.ProductName `
                $json.Name)
            IdentityInfoValue = [string](@($json.IdentityInfoValue | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | Select-Object -First 1))
            PartNumber  = [string](Get-DriFTFirstNonEmpty `
                $json.PartNumber `
                $json.PartNumberString `
                $json.SparePartNumber `
                $json.Oem.Dell.PartNumber `
                $json.Oem.Dell.SparePartNumber `
                $json.Oem.Dell.MMPartNumber `
                $json.Oem.Dell.ManufacturerPartNumber)
            SerialNumber = [string](Get-DriFTFirstNonEmpty `
                $json.SerialNumber `
                $json.Oem.Dell.SerialNumber)
            FQDD        = [string]$fqdd
            SoftwareId  = [string]$json.SoftwareId
            OdataType   = $odataType
            VendorID    = Convert-DriFTHexId $vendorId
            DeviceID    = Convert-DriFTHexId $deviceId
            SubVendorID = Convert-DriFTHexId $subVendorId
            SubDeviceID = Convert-DriFTHexId $subDeviceId
            KeyText     = (@($keys) -join '|')
            SourceFile  = [string]$jsonFile.FullName
        })
    }

    return @($rows) | Sort-Object OdataId,Id,FQDD,VendorID,DeviceID,SubVendorID,SubDeviceID -Unique
}

function Add-DriFT17GPciIdentityToFirmwareRows {
<#
.SYNOPSIS
    Enriches 17G firmware rows with PCI identity.

.DESCRIPTION
    Correlates firmware inventory RelatedItem/display/FQDD values to PCI identity rows.
#>
    [CmdletBinding()]
    param(
        [AllowNull()][AllowEmptyCollection()][object[]]$FirmwareRows,
        [AllowNull()][AllowEmptyCollection()][object[]]$PciRows,
        [Parameter(Mandatory)]$Context
    )

    $FirmwareRows = @($FirmwareRows)
    $PciRows = @($PciRows)
    $rows = @()

    foreach ($fw in $FirmwareRows) {
        $pci = Find-DriFT17GPciRecordForFirmware -FirmwareRow $fw -PciRows $PciRows

        $vendorId = $fw.VendorID
        $deviceId = $fw.DeviceID
        $subVendorId = $fw.SubVendorID
        $subDeviceId = $fw.SubDeviceID
        $source = $fw.Source

        if ($pci) {
            if ($pci.VendorID) { $vendorId = $pci.VendorID }
            if ($pci.DeviceID) { $deviceId = $pci.DeviceID }
            if ($pci.SubVendorID) { $subVendorId = $pci.SubVendorID }
            if ($pci.SubDeviceID) { $subDeviceId = $pci.SubDeviceID }
            $source = 'RedfishFirmwareInventory+PCI'
        }

        $item = New-DriFTInventoryItem `
            -SourceGeneration '17G' `
            -ComponentType $fw.ComponentType `
            -ComponentID $fw.ComponentID `
            -VendorID $vendorId `
            -DeviceID $deviceId `
            -SubVendorID $subVendorId `
            -SubDeviceID $subDeviceId `
            -Version $fw.Version `
            -Display $fw.Display `
            -ElementName $fw.ElementName `
            -RelatedItem $fw.RelatedItem `
            -Source $source

        if ($pci) {
            foreach ($extra in @(
                @{ Name = 'OriginalElementName'; Value = (Get-DriFTFirstNonEmpty $pci.Description $pci.Name) },
                @{ Name = 'OriginalId';          Value = $pci.Id },
                @{ Name = 'IdentityInfoValue';   Value = $pci.IdentityInfoValue },
                @{ Name = 'Description';         Value = (Get-DriFTFirstNonEmpty $pci.Description $pci.Name) },
                @{ Name = 'PartNumber';          Value = $pci.PartNumber },
                @{ Name = 'SerialNumber';        Value = $pci.SerialNumber },
                @{ Name = 'FQDD';                Value = $pci.FQDD },
                @{ Name = 'PciSourceFile';       Value = $pci.SourceFile }
            )) {
                if (-not [string]::IsNullOrWhiteSpace([string]$extra.Value)) {
                    $item | Add-Member -MemberType NoteProperty -Name $extra.Name -Value ([string]$extra.Value) -Force
                }
            }
        }

        $rows += @($item)
    }

    return @($rows)
}

function Find-DriFT17GPciRecordForFirmware {
<#
.SYNOPSIS
    Finds the best PCI identity row for one firmware row.

.DESCRIPTION
    Prefer DellSoftwareInventory identity records over generic PCIeFunction rows.
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$FirmwareRow,
        [Parameter(Mandatory)][object[]]$PciRows
    )

    if (-not $PciRows) { return $null }

    $needles = New-Object System.Collections.Generic.List[string]
    foreach ($candidate in @($FirmwareRow.RelatedItem, $FirmwareRow.Display, $FirmwareRow.ElementName, $FirmwareRow.ComponentID)) {
        if ([string]::IsNullOrWhiteSpace($candidate)) { continue }

        [void]$needles.Add(([string]$candidate).Trim())
        if ($candidate -match '/') {
            [void]$needles.Add((([string]$candidate).TrimEnd('/') -split '/')[-1])
        }

        foreach ($variant in Get-DriFT17GFqddVariants -Value $candidate) {
            if ($variant) { [void]$needles.Add($variant) }
        }
    }

    $needles = @($needles | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique)

    $prioritized = @($PciRows | Sort-Object @{ Expression = { if ($_.SourceFile -imatch 'DellSoftwareInventory') { 0 } else { 1 } } }, FQDD, Id)

    foreach ($needle in $needles) {
        if ($needle -match '^NIC\.Slot\.\d+$') {
            $escapedBase = [regex]::Escape($needle)
            $hit = $prioritized | Where-Object {
                ($_.SourceFile -imatch 'DellSoftwareInventory') -and
                ($_.VendorID -or $_.DeviceID -or $_.SubVendorID -or $_.SubDeviceID) -and
                (
                    ($_.FQDD -imatch "^$escapedBase(?:-|$)") -or
                    ($_.Id -imatch "DCIM[:_](CURRENT|INSTALLED).*$escapedBase(?:-|$)") -or
                    ($_.KeyText -imatch "(^|\|)$escapedBase(?:-|\||$)")
                )
            } | Sort-Object FQDD | Select-Object -First 1

            if ($hit) { return $hit }
        }
    }

    foreach ($needle in $needles) {
        $escaped = [regex]::Escape($needle)
        $hit = $prioritized | Where-Object {
            ($_.FQDD -ieq $needle) -or
            ($_.Id -ieq $needle) -or
            ($_.OdataId -ieq $needle) -or
            ($_.SoftwareId -ieq $needle) -or
            ($_.KeyText -imatch "(^|\|)$escaped(\||$)")
        } | Select-Object -First 1

        if ($hit) { return $hit }
    }

    return $null
}

#endregion 17G Redfish Parser

#region SAE XML Parser

function Get-DriFTSaeXmlSystemIdentity {
<#
.SYNOPSIS
    Gets system identity from SAE XML collections.

.DESCRIPTION
    Placeholder body keeps SAE XML capability in the new architecture.
    Port existing MaserInfo.xml / Inventory.xml / chasinfo.xml logic here.
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Collection,
        [Parameter(Mandatory)]$Context
    )

    throw 'SAE XML system identity parser not yet ported. Port existing SAEX logic into Get-DriFTSaeXmlSystemIdentity.'
}

function Get-DriFTSaeXmlOperatingSystem {
<#
.SYNOPSIS
    Gets OS information from SAE XML collections.

.DESCRIPTION
    Placeholder body keeps SAE XML capability in the new architecture.
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Collection,
        [Parameter(Mandatory)]$Context
    )

    throw 'SAE XML OS parser not yet ported. Port existing SAEX logic into Get-DriFTSaeXmlOperatingSystem.'
}

function Import-DriFTSaeXmlInventory {
<#
.SYNOPSIS
    Imports installed inventory from SAE XML collections.

.DESCRIPTION
    Placeholder body keeps SAE XML capability in the new architecture.
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Collection,
        [Parameter(Mandatory)]$Context
    )

    throw 'SAE XML inventory parser not yet ported. Port existing SAEX logic into Import-DriFTSaeXmlInventory.'
}

#endregion SAE XML Parser

#region Catalog Download / Import / Filtering

function Initialize-DriFTCatalogSet {
<#
.SYNOPSIS
    Downloads/imports the Dell catalog set.

.DESCRIPTION
    Loads Catalog.cab/Catalog.xml and prepares placeholders for ASHCI and Precision
    catalogs when requested by a system.
#>
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Context)

    $catalogPath = Get-DriFTDellCatalog -Context $Context
    $catalogXmlPath = Expand-DriFTCatalogCab -CabPath $catalogPath -Context $Context

    $catalogXml = [xml](Get-Content -Raw -Path $catalogXmlPath)

    [PSCustomObject]@{
        MainCatalogPath = $catalogXmlPath
        MainCatalog     = $catalogXml
        MainInfo        = "Catalog.xml<br> $($catalogXml.Manifest.version)"
        MainVersion     = [string]$catalogXml.Manifest.version
        AshciCatalog    = $null
        AshciInfo       = $null
        AshciVersion    = $null
        PrecisionCache  = @{}
        CatVerInfo      = "<br> Catalog.xml Info <br>&nbsp&nbspData/Time: $($catalogXml.Manifest.dateTime)<br>&nbsp&nbspReleaseId: $($catalogXml.Manifest.releaseID)<br>&nbsp&nbspVersion: $($catalogXml.Manifest.version)"
    }
}

function Get-DriFTDellCatalog {
<#
.SYNOPSIS
    Downloads Catalog.cab.

.DESCRIPTION
    Uses Invoke-WebRequest with default credentials for proxy-friendly environments.
#>
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Context)

    $url = 'https://downloads.dell.com/catalog/Catalog.cab'
    $path = Join-Path $Context.CatalogRoot 'Catalog.cab'

    Write-DriFTLog -Context $Context -Message 'Downloading Catalog.cab...' -Level Info

    try {
        Invoke-WebRequest -Uri $url -OutFile $path -UseDefaultCredentials -ErrorAction Stop
        return $path
    }
    catch {
        Write-DriFTLog -Context $Context -Message "Catalog download failed: $($_.Exception.Message)" -Level Warn
        throw 'Catalog.cab download failed. Add manual catalog selection fallback here if offline support is required.'
    }
}


function Get-DriFTAshciCatalog {
<#
.SYNOPSIS
    Downloads and extracts the Dell ASHCI catalog.

.DESCRIPTION
    Azure Stack HCI / AX systems should use the ASHCI-Catalog.xml.gz catalog.
    The downloaded .gz is decompressed, then sanitized because some catalog payloads
    can include extra content or repeated XML declarations after the first Manifest.
#>
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Context)

    $url = 'https://downloads.dell.com/catalog/ASHCI-Catalog.xml.gz'
    $gzPath = Join-Path $Context.CatalogRoot 'ASHCI-Catalog.xml.gz'
    $xmlPath = Join-Path $Context.CatalogRoot 'ASHCI-Catalog.xml'

    Write-DriFTLog -Context $Context -Message 'Downloading ASHCI-Catalog.xml.gz...' -Level Info

    try {
        foreach ($stale in @($gzPath, $xmlPath)) {
            if (Test-Path -LiteralPath $stale -PathType Leaf) {
                Remove-Item -LiteralPath $stale -Force -ErrorAction SilentlyContinue
            }
        }

        Invoke-WebRequest -Uri $url -OutFile $gzPath -UseDefaultCredentials -ErrorAction Stop
        Expand-DriFTGzipFile -Path $gzPath -DestinationPath $xmlPath | Out-Null
        Repair-DriFTCatalogXmlFile -Path $xmlPath -Context $Context | Out-Null

        return $xmlPath
    }
    catch {
        Write-DriFTLog -Context $Context -Message "ASHCI catalog download/extract failed: $($_.Exception.Message)" -Level Warn -Indent 1
        return $null
    }
}

function Expand-DriFTGzipFile {
<#
.SYNOPSIS
    Decompresses a .gz file to a target path.
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$DestinationPath
    )

    $sourceStream = [System.IO.File]::OpenRead($Path)
    $gzipStream = $null
    $targetStream = $null

    try {
        $gzipStream = New-Object System.IO.Compression.GzipStream($sourceStream, [System.IO.Compression.CompressionMode]::Decompress)
        $targetStream = [System.IO.File]::Create($DestinationPath)
        $buffer = New-Object byte[] 8192

        while (($read = $gzipStream.Read($buffer, 0, $buffer.Length)) -gt 0) {
            $targetStream.Write($buffer, 0, $read)
        }
    }
    finally {
        if ($targetStream) { $targetStream.Dispose() }
        if ($gzipStream) { $gzipStream.Dispose() }
        if ($sourceStream) { $sourceStream.Dispose() }
    }

    return $DestinationPath
}

function Repair-DriFTCatalogXmlFile {
<#
.SYNOPSIS
    Sanitizes a Dell catalog XML file before XML parsing.

.DESCRIPTION
    Some catalog downloads can contain extra data or repeated XML declarations after
    the first Manifest document. This function keeps only the first complete
    <?xml ...?><Manifest>...</Manifest> document so [xml] parsing is reliable.
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)]$Context
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return $null }

    $raw = [System.IO.File]::ReadAllText($Path, [System.Text.Encoding]::UTF8)

    $match = [regex]::Match(
        $raw,
        '(?s)<\?xml[^>]*>\s*<Manifest\b.*?</Manifest>',
        [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
    )

    if ($match.Success) {
        $clean = $match.Value.Trim()
        if ($clean.Length -lt $raw.Length) {
            Write-DriFTLog -Context $Context -Message "Sanitized catalog XML from $($raw.Length) to $($clean.Length) characters." -Level Warn -Indent 1
        }

        [System.IO.File]::WriteAllText($Path, $clean, [System.Text.Encoding]::UTF8)
        return $Path
    }

    # Fallback: trim anything before the first XML declaration and after the first closing Manifest.
    $start = $raw.IndexOf('<?xml')
    $end = $raw.IndexOf('</Manifest>')
    if ($start -ge 0 -and $end -gt $start) {
        $clean = $raw.Substring($start, ($end + '</Manifest>'.Length) - $start).Trim()
        [System.IO.File]::WriteAllText($Path, $clean, [System.Text.Encoding]::UTF8)
        return $Path
    }

    return $Path
}


function Ensure-DriFTAshciCatalogLoaded {
<#
.SYNOPSIS
    Lazily loads the ASHCI catalog into the catalog set.
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$CatalogSet,
        [Parameter(Mandatory)]$Context
    )

    if ($CatalogSet.AshciCatalog) { return $true }

    $ashciPath = Get-DriFTAshciCatalog -Context $Context
    if (-not $ashciPath -or -not (Test-Path -LiteralPath $ashciPath -PathType Leaf)) {
        return $false
    }

    try {
        Repair-DriFTCatalogXmlFile -Path $ashciPath -Context $Context | Out-Null
        $ashciXmlText = [System.IO.File]::ReadAllText($ashciPath, [System.Text.Encoding]::UTF8)
        $ashciXml = New-Object System.Xml.XmlDocument
        $ashciXml.PreserveWhitespace = $false
        $ashciXml.LoadXml($ashciXmlText)
        $CatalogSet.AshciCatalog = $ashciXml
        $CatalogSet.AshciInfo = "ASHCI-Catalog.xml<br> $($ashciXml.Manifest.version)"
        if ($CatalogSet.PSObject.Properties.Name -notcontains 'AshciVersion') {
            $CatalogSet | Add-Member -MemberType NoteProperty -Name AshciVersion -Value ([string]$ashciXml.Manifest.version) -Force
        }
        else {
            $CatalogSet.AshciVersion = [string]$ashciXml.Manifest.version
        }

        $CatalogSet.CatVerInfo += "<br><br> ASHCI-Catalog.xml Info <br>&nbsp&nbspData/Time: $($ashciXml.Manifest.dateTime)<br>&nbsp&nbspReleaseId: $($ashciXml.Manifest.releaseID)<br>&nbsp&nbspVersion: $($ashciXml.Manifest.version)"

        Write-DriFTLog -Context $Context -Message "ASHCI catalog loaded: version $($ashciXml.Manifest.version)" -Level Success -Indent 1
        return $true
    }
    catch {
        Write-DriFTLog -Context $Context -Message "ASHCI catalog parse failed: $($_.Exception.Message)" -Level Warn -Indent 1
        return $false
    }
}

function Get-DriFTCatalogRowsForSystem {
<#
.SYNOPSIS
    Filters one catalog XML document for the current system and package type.
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$CatalogXml,
        [Parameter(Mandatory)]$System,
        [Parameter(Mandatory)]$OperatingSystem
    )

    return @($CatalogXml.Manifest.SoftwareComponent |
        Where-Object {
            ($_.SupportedSystems.Brand.Model.Display.'#cdata-section' -eq $System.PowerEdge) -or
            ($_.SupportedSystems.Brand.Model.systemID -eq $System.SystemID) -or
            ($_.SupportedSystems.Brand.Model.systemID -match 'VRTX')
        } |
        Where-Object {
            ($_.packageType -eq $OperatingSystem.CatalogPackageType) -or
            ($_.packageType -eq 'LW64') -or
            ($_.packageType -eq 'LWXP')
        })
}

function Expand-DriFTCatalogCab {
<#
.SYNOPSIS
    Extracts Catalog.xml from Catalog.cab.

.DESCRIPTION
    Uses Shell.Application because CAB extraction is supported there on Windows.
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$CabPath,
        [Parameter(Mandatory)]$Context
    )

    $dest = Join-Path $Context.CatalogRoot 'Main'
    New-Item -ItemType Directory -Force -Path $dest | Out-Null
    $target = Join-Path $dest 'Catalog.xml'

    if (Test-Path $target) { Remove-Item $target -Force }

    $shell = New-Object -ComObject Shell.Application
    $cab = $shell.NameSpace($CabPath)
    $folder = $shell.NameSpace($dest)

    if (-not $cab -or -not $folder) {
        throw "Unable to open catalog CAB: $CabPath"
    }

    $item = $cab.ParseName('Catalog.xml')
    if (-not $item) { throw 'Catalog.xml was not found inside Catalog.cab.' }

    $folder.CopyHere($item)

    $timer = [Diagnostics.Stopwatch]::StartNew()
    while (-not (Test-Path $target -PathType Leaf) -and $timer.Elapsed.TotalSeconds -lt 30) {
        Start-Sleep -Milliseconds 250
    }

    if (-not (Test-Path $target -PathType Leaf)) {
        throw 'Timed out waiting for Catalog.xml extraction.'
    }

    return $target
}


function Add-DriFTCatalogSourceInfo {
<#
.SYNOPSIS
    Adds source catalog metadata to filtered SoftwareComponent rows.

.DESCRIPTION
    Catalog XML nodes are live XML element objects. This helper tags each row
    with SourceCatalogName and SourceCatalogInfo note properties so later matching
    and reporting can identify whether the update came from ASHCI-Catalog.xml or
    the general Dell Catalog.xml.
#>
    [CmdletBinding()]
    param(
        [AllowNull()][object[]]$Rows,
        [Parameter(Mandatory)][string]$SourceCatalogName,
        [Parameter(Mandatory)][string]$SourceCatalogInfo
    )

    foreach ($row in @($Rows | Where-Object { $null -ne $_ })) {
        if ($row.PSObject.Properties.Name -notcontains 'SourceCatalogName') {
            $row | Add-Member -MemberType NoteProperty -Name SourceCatalogName -Value $SourceCatalogName -Force
        }
        else {
            $row.SourceCatalogName = $SourceCatalogName
        }

        if ($row.PSObject.Properties.Name -notcontains 'SourceCatalogInfo') {
            $row | Add-Member -MemberType NoteProperty -Name SourceCatalogInfo -Value $SourceCatalogInfo -Force
        }
        else {
            $row.SourceCatalogInfo = $SourceCatalogInfo
        }
    }

    return @($Rows)
}


function Format-DriFTCatalogInfo {
<#
.SYNOPSIS
    Formats catalog info consistently for report output.

.DESCRIPTION
    Always returns:
      Catalog.xml 26.05.15
      ASHCI-Catalog.xml 26.03.03

    instead of mixing multiline and plain catalog names.
#>
    [CmdletBinding()]
    param(
        [AllowNull()][string]$CatalogName,
        [AllowNull()][string]$CatalogVersion
    )

    $name = Get-DriFTFirstNonEmpty $CatalogName 'Catalog.xml'
    $version = Get-DriFTFirstNonEmpty $CatalogVersion

    if ([string]::IsNullOrWhiteSpace($version)) {
        return $name
    }

    return "$name $version"
}


function Get-DriFTCatalogSourceInfo {
<#
.SYNOPSIS
    Returns the report CatalogInfo text for a catalog row.
#>
    [CmdletBinding()]
    param([AllowNull()]$CatalogRow)

    $info = Get-DriFTFirstNonEmpty $CatalogRow.SourceCatalogInfo $CatalogRow.SourceCatalogName
    if ([string]::IsNullOrWhiteSpace($info)) { return 'Catalog.xml' }
    return $info
}

function ConvertTo-DriFTCriticality {
<#
.SYNOPSIS
    Normalizes Dell catalog criticality text.

.DESCRIPTION
    The report should only show Urgent, Recommended, Optional, or Not Available.
    Some catalogs include extra wording; this strips it to the useful severity.
#>
    [CmdletBinding()]
    param([AllowNull()]$Value)

    $text = Get-DriFTFirstNonEmpty $Value
    if ([string]::IsNullOrWhiteSpace($text)) { return 'Not Available' }

    if ($text -imatch 'Urgent') { return 'Urgent' }
    if ($text -imatch 'Recommended') { return 'Recommended' }
    if ($text -imatch 'Optional') { return 'Optional' }

    return $text.Trim()
}

function Get-DriFTCatalogRowKey {
<#
.SYNOPSIS
    Builds a stable deduplication key for a catalog row.
#>
    [CmdletBinding()]
    param([AllowNull()]$CatalogRow)

    if ($null -eq $CatalogRow) { return '' }

    $name = Get-DriFTFirstNonEmpty $CatalogRow.Name.Display.'#cdata-section' $CatalogRow.name
    $version = Get-DriFTFirstNonEmpty $CatalogRow.vendorVersion
    $path = Get-DriFTFirstNonEmpty $CatalogRow.path
    $category = Get-DriFTFirstNonEmpty $CatalogRow.LUCategory.value

    return (@($category, $name, $version, $path) -join '|').ToLowerInvariant()
}

function Join-DriFTCatalogRowsAshciFirst {
<#
.SYNOPSIS
    Merges ASHCI and Dell catalog rows with ASHCI priority.

.DESCRIPTION
    ASHCI rows come first. Dell Catalog.xml rows are appended only when an equivalent
    row is not already present in ASHCI. This gives AX/HCI systems ASHCI-first
    matching while preserving Dell Catalog.xml fallback for components not found
    in ASHCI.
#>
    [CmdletBinding()]
    param(
        [AllowNull()][object[]]$AshciRows,
        [AllowNull()][object[]]$DellRows
    )

    $merged = @()
    $seen = @{}

    foreach ($row in @($AshciRows | Where-Object { $null -ne $_ })) {
        $key = Get-DriFTCatalogRowKey -CatalogRow $row
        if (-not $seen.ContainsKey($key)) {
            $seen[$key] = $true
            $merged += @($row)
        }
    }

    foreach ($row in @($DellRows | Where-Object { $null -ne $_ })) {
        $key = Get-DriFTCatalogRowKey -CatalogRow $row
        if (-not $seen.ContainsKey($key)) {
            $seen[$key] = $true
            $merged += @($row)
        }
    }

    return @($merged)
}



function Get-DriFTPlatformInfo {
<#
.SYNOPSIS
    Detects the hardware platform family for the active collection.

.DESCRIPTION
    Provides a single routing object for catalog selection and matching behavior.
    This keeps PowerEdge, AX/HCI, Precision, and future client platforms from being
    hardcoded throughout the report engine.
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$System,
        [Parameter(Mandatory)]$Context
    )

    $text = @(
        $System.PowerEdge,
        $System.Model,
        $System.SystemModel,
        $System.ChassisName,
        $System.SystemID,
        $System.ServiceTag,
        $System.SourceType
    ) -join ' '

    if ($text -imatch 'Precision') {
        return [PSCustomObject]@{
            Type        = 'Precision'
            Family      = 'Workstation'
            CatalogType = 'Precision'
            IsPrecision = $true
            IsPowerEdge = $false
            IsHCI       = $false
        }
    }

    if ($System.SpecialCatalogNeeded -eq 'HCI' -or
        $System.S2DCatalogNeeded -eq 'YES' -or
        $System.PowerEdge -match '^AX|Azure|Storage Spaces Direct') {
        return [PSCustomObject]@{
            Type        = 'AX/HCI'
            Family      = 'PowerEdge'
            CatalogType = 'ASHCI'
            IsPrecision = $false
            IsPowerEdge = $true
            IsHCI       = $true
        }
    }

    return [PSCustomObject]@{
        Type        = 'PowerEdge'
        Family      = 'Server'
        CatalogType = 'Catalog'
        IsPrecision = $false
        IsPowerEdge = $true
        IsHCI       = $false
    }
}

function Test-DriFTPrecisionSystem {
<#
.SYNOPSIS
    Determines whether a system identity represents a Precision workstation.
#>
    [CmdletBinding()]
    param([Parameter(Mandatory)]$System)

    $text = @(
        $System.PowerEdge,
        $System.Model,
        $System.SystemModel,
        $System.ChassisName
    ) -join ' '

    return ($text -imatch 'Precision')
}

function Get-DriFTPrecisionModelTokens {
<#
.SYNOPSIS
    Builds normalized Precision model tokens for catalog matching.

.DESCRIPTION
    Precision package metadata often uses multiple naming forms, for example:
      Precision 5820 Tower
      Precision Tower 5820
      Precision 5820
      5820
#>
    [CmdletBinding()]
    param([AllowNull()][string]$Model)

    if ([string]::IsNullOrWhiteSpace($Model)) { return @() }

    $m = ([string]$Model).Trim()
    $tokens = @($m)

    if ($m -match '(?i)Precision\s+(?<num>\d{4})') {
        $num = $Matches.num
        $tokens += @(
            "Precision $num",
            "Precision Tower $num",
            "Precision $num Tower",
            $num
        )
    }

    if ($m -match '(?i)Precision\s+(?<word>Mobile|Tower|Rack)\s+(?<num>\d{4})') {
        $num = $Matches.num
        $word = $Matches.word
        $tokens += @(
            "Precision $word $num",
            "Precision $num $word",
            "Precision $num",
            $num
        )
    }

    return @($tokens | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique)
}

function Test-DriFTCatalogRowAppliesToPrecision {
<#
.SYNOPSIS
    Tests whether a catalog row appears applicable to the Precision workstation.

.DESCRIPTION
    This is intentionally permissive during the first Precision pass. Dell client
    catalogs frequently surface supported system metadata differently from server
    catalogs. We first prefer explicit model token matches, then allow rows from a
    catalog that was already filtered to the system.
#>
    [CmdletBinding()]
    param(
        [AllowNull()]$CatalogRow,
        [Parameter(Mandatory)]$System
    )

    if ($null -eq $CatalogRow) { return $false }

    $modelText = Get-DriFTFirstNonEmpty $System.PowerEdge $System.Model $System.SystemModel $System.ChassisName
    $tokens = @(Get-DriFTPrecisionModelTokens -Model $modelText)

    $rowText = @(
        $CatalogRow.SupportedSystems.Brand.Model.Display.'#cdata-section',
        $CatalogRow.SupportedSystems.Brand.Model.systemID,
        $CatalogRow.SupportedSystems.Brand.Model.name,
        $CatalogRow.SupportedSystems.Brand.Display.'#cdata-section',
        $CatalogRow.Name.Display.'#cdata-section',
        $CatalogRow.path,
        $CatalogRow.Description.Display.'#cdata-section'
    ) -join ' '

    foreach ($token in $tokens) {
        if ($rowText -imatch [regex]::Escape($token)) { return $true }
    }

    $systemId = Get-DriFTFirstNonEmpty $System.SystemID
    if (-not [string]::IsNullOrWhiteSpace($systemId) -and $rowText -imatch [regex]::Escape($systemId)) {
        return $true
    }

    return $false
}

function Get-DriFTPrecisionCatalog {
<#
.SYNOPSIS
    Loads a Precision/workstation applicable catalog.

.DESCRIPTION
    First implementation uses the main Dell Catalog.xml already downloaded by DriFT.
    Precision package rows are tagged as Precision catalog rows when the platform is
    Precision. This keeps the implementation safe while preserving future ability to
    replace this with a dedicated Dell Command Update/client catalog source.
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$CatalogSet,
        [Parameter(Mandatory)]$Context
    )

    # Placeholder hook for a future dedicated workstation/client catalog.
    # For now, return the main catalog so Precision uses the same Dell catalog
    # download path but with Precision-aware filtering/matching.
    return $CatalogSet.MainCatalog
}


function Get-DriFTApplicableCatalogRows {
<#
.SYNOPSIS
    Filters catalog rows for the current system and OS.

.DESCRIPTION
    Catalog routing priority:
      1. Precision workstation path
      2. AX / Azure Stack HCI ASHCI-first path
      3. Standard Dell Catalog.xml path
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$CatalogSet,
        [Parameter(Mandatory)]$System,
        [Parameter(Mandatory)]$OperatingSystem,
        [Parameter(Mandatory)]$Context
    )

    $platform = Get-DriFTPlatformInfo -System $System -Context $Context

    if ($Context.PSObject.Properties.Name -notcontains 'Platform') {
        $Context | Add-Member -MemberType NoteProperty -Name Platform -Value $platform -Force
    }
    else {
        $Context.Platform = $platform
    }

    $mainRows = @(Get-DriFTCatalogRowsForSystem `
        -CatalogXml $CatalogSet.MainCatalog `
        -System $System `
        -OperatingSystem $OperatingSystem)

    $mainRows = @(Add-DriFTCatalogSourceInfo `
        -Rows $mainRows `
        -SourceCatalogName 'Catalog.xml' `
        -SourceCatalogInfo (Format-DriFTCatalogInfo `
            -CatalogName 'Catalog.xml' `
            -CatalogVersion $CatalogSet.MainVersion))

    # Precision workstation support
    if ($platform.IsPrecision) {
        Write-DriFTLog -Context $Context -Message "Precision workstation detected. Using Precision-aware Dell catalog matching..." -Level Info -Indent 1

        $precisionCatalog = Get-DriFTPrecisionCatalog -CatalogSet $CatalogSet -Context $Context

        $precisionRows = @(Get-DriFTCatalogRowsForSystem `
            -CatalogXml $precisionCatalog `
            -System $System `
            -OperatingSystem $OperatingSystem)

        if (@($precisionRows).Count -eq 0) {
            # Some workstation catalog metadata does not filter cleanly with the
            # server-style SupportedSystems path. Fall back to model-token scan.
            $allSoftwareComponents = @($precisionCatalog.Manifest.SoftwareComponent | Where-Object { $null -ne $_ })

            $precisionRows = @($allSoftwareComponents | Where-Object {
                Test-DriFTCatalogRowAppliesToPrecision -CatalogRow $_ -System $System
            })
        }

        $precisionRows = @(Add-DriFTCatalogSourceInfo `
            -Rows $precisionRows `
            -SourceCatalogName 'Catalog.xml' `
            -SourceCatalogInfo (Format-DriFTCatalogInfo `
                -CatalogName 'Catalog.xml' `
                -CatalogVersion $CatalogSet.MainVersion))

        $allRows = @(Join-DriFTCatalogRowsAshciFirst -AshciRows $precisionRows -DellRows $mainRows)

        Write-DriFTLog -Context $Context -Message "Precision applicable rows: $(@($precisionRows).Count); Dell fallback rows: $(@($mainRows).Count); merged rows: $(@($allRows).Count)" -Level Info -Indent 1

        return [PSCustomObject]@{
            MainRows      = $mainRows
            AshciRows     = @()
            PrecisionRows = $precisionRows
            AllRows       = @($allRows)
            MainInfo      = (Format-DriFTCatalogInfo -CatalogName 'Catalog.xml' -CatalogVersion $CatalogSet.MainVersion)
            AshciInfo     = $CatalogSet.AshciInfo
            PrecisionInfo = (Format-DriFTCatalogInfo -CatalogName 'Catalog.xml' -CatalogVersion $CatalogSet.MainVersion)
            ActiveInfo    = (Format-DriFTCatalogInfo -CatalogName 'Catalog.xml' -CatalogVersion $CatalogSet.MainVersion)
        }
    }

    $ashciRows = @()
    $allRows = $mainRows
    $activeInfo = $CatalogSet.MainInfo

    if ($platform.IsHCI) {
        Write-DriFTLog -Context $Context -Message 'HCI/AX system detected. Loading ASHCI catalog...' -Level Info -Indent 1

        if (Ensure-DriFTAshciCatalogLoaded -CatalogSet $CatalogSet -Context $Context) {
            $ashciRows = @(Get-DriFTCatalogRowsForSystem `
                -CatalogXml $CatalogSet.AshciCatalog `
                -System $System `
                -OperatingSystem $OperatingSystem)

            $ashciRows = @(Add-DriFTCatalogSourceInfo `
                -Rows $ashciRows `
                -SourceCatalogName 'ASHCI-Catalog.xml' `
                -SourceCatalogInfo (Format-DriFTCatalogInfo `
                    -CatalogName 'ASHCI-Catalog.xml' `
                    -CatalogVersion $CatalogSet.AshciVersion))

            if (@($ashciRows).Count -gt 0) {
                $allRows = @(Join-DriFTCatalogRowsAshciFirst -AshciRows $ashciRows -DellRows $mainRows)
                $activeInfo = "$($CatalogSet.AshciInfo)<br>Fallback: $($CatalogSet.MainInfo)"
                Write-DriFTLog -Context $Context -Message "Using ASHCI catalog first. ASHCI rows: $(@($ashciRows).Count); Dell fallback rows: $(@($mainRows).Count); merged rows: $(@($allRows).Count)" -Level Success -Indent 1
            }
            else {
                Write-DriFTLog -Context $Context -Message "ASHCI catalog loaded but returned no applicable rows. Using main Catalog.xml rows: $(@($mainRows).Count)" -Level Warn -Indent 1
                $allRows = $mainRows
                $activeInfo = $CatalogSet.MainInfo
            }
        }
        else {
            Write-DriFTLog -Context $Context -Message "ASHCI catalog unavailable. Using main Catalog.xml rows: $(@($mainRows).Count)" -Level Warn -Indent 1
            $allRows = $mainRows
            $activeInfo = $CatalogSet.MainInfo
        }
    }
    else {
        Write-DriFTLog -Context $Context -Message "Using main Catalog.xml rows for matching/reporting. Rows: $(@($mainRows).Count)" -Level Info -Indent 1
    }

    [PSCustomObject]@{
        MainRows      = $mainRows
        AshciRows     = $ashciRows
        PrecisionRows = @()
        AllRows       = @($allRows)
        MainInfo      = (Format-DriFTCatalogInfo -CatalogName 'Catalog.xml' -CatalogVersion $CatalogSet.MainVersion)
        AshciInfo     = (Format-DriFTCatalogInfo -CatalogName 'ASHCI-Catalog.xml' -CatalogVersion $CatalogSet.AshciVersion)
        PrecisionInfo = $null
        ActiveInfo    = $activeInfo
    }
}

function New-DriFTCatalogIndex {
<#
.SYNOPSIS
    Builds lookup indexes from filtered catalog rows.

.DESCRIPTION
    Improves speed and accuracy by replacing repeated Where-Object scans with
    component and PCI identity indexes.
#>
    [CmdletBinding()]
    param(
        [AllowEmptyCollection()][object[]]$CatalogRows,
        [Parameter(Mandatory)]$Context
    )

    $byComponent = @{}
    $byTypeComponent = @{}
    $byPci = @{}
    $byTypePci = @{}

    foreach ($row in @($CatalogRows)) {
        $componentType = Get-DriFTCatalogComponentTypeValue -CatalogObject $row

        foreach ($componentId in Get-DriFTCatalogComponentIdValues -CatalogObject $row) {
            if ([string]::IsNullOrWhiteSpace($componentId)) { continue }
            Add-DriFTIndexValue -Index $byComponent -Key $componentId -Value $row
            Add-DriFTIndexValue -Index $byTypeComponent -Key "$componentType|$componentId" -Value $row
        }

        foreach ($pci in Get-DriFTCatalogPciInfoObjects -CatalogObject $row) {
            $pciKey = New-DriFTPciKey `
                -VendorID $pci.vendorID `
                -DeviceID $pci.deviceID `
                -SubVendorID $pci.subVendorID `
                -SubDeviceID $pci.subDeviceID

            if ($pciKey) {
                Add-DriFTIndexValue -Index $byPci -Key $pciKey -Value $row
                Add-DriFTIndexValue -Index $byTypePci -Key "$componentType|$pciKey" -Value $row
            }
        }
    }

    [PSCustomObject]@{
        ByComponent     = $byComponent
        ByTypeComponent = $byTypeComponent
        ByPci           = $byPci
        ByTypePci       = $byTypePci
    }
}

function Add-DriFTIndexValue {
<#
.SYNOPSIS
    Adds one object to a hashtable index.
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][hashtable]$Index,
        [Parameter(Mandatory)][string]$Key,
        [Parameter(Mandatory)]$Value
    )

    if (-not $Index.ContainsKey($Key)) {
        $Index[$Key] = @()
    }

    $Index[$Key] += @($Value)
}

#endregion Catalog Download / Import / Filtering

#region Matching Engine

function Compare-DriFTInventoryToCatalog {
<#
.SYNOPSIS
    Compares normalized inventory to filtered catalog rows.

.DESCRIPTION
    Uses indexed exact component/PCI matching. Accuracy rules:
      - Catalog rows must already be filtered by system/OS.
      - ComponentID 0 is ignored.
      - BIOS gets a component type exception because 17G reports BIOS as FRMW while Catalog.xml uses BIOS.
      - Non-BIOS matches require compatible component type and exact ComponentID or full PCI identity.
#>
    [CmdletBinding()]
    param(
        [AllowNull()][AllowEmptyCollection()][object[]]$Inventory,
        [Parameter(Mandatory)]$CatalogIndex,
        [Parameter(Mandatory)]$CatalogRows,
        [Parameter(Mandatory)]$System,
        [Parameter(Mandatory)]$OperatingSystem,
        [Parameter(Mandatory)]$Context
    )

    $Inventory = @($Inventory | Where-Object { $null -ne $_ })
    $matches = @()

    foreach ($device in @($Inventory)) {
        if ($null -eq $device) { continue }

        $candidateRows = Find-DriFTCatalogCandidates -Device $device -CatalogIndex $CatalogIndex
        $best = $candidateRows |
            Where-Object { Test-DriFTCatalogDeviceMatch -CatalogDevice $_ -Device $device } |
            Sort-Object { ConvertTo-DriFTSafeDateTime $_.releaseDate } |
            Select-Object -Last 1

        $nameFallbackUsed = $false
        if (-not $best) {
            $nameFallback = Find-DriFTCatalogNameFallbackMatch -Device $device -CatalogRows $CatalogRows.AllRows
            if ($nameFallback) {
                $best = $nameFallback
                $nameFallbackUsed = $true
            }
        }

        $method = ''
        $reason = ''

        if ($best) {
            $componentMatch = Test-DriFTCatalogComponentIdMatch -CatalogObject $best -ComponentId $device.ComponentID
            $pciMatch = Test-DriFTCatalogPciIdentityMatch -CatalogObject $best -Device $device

            if ($nameFallbackUsed) { $method = 'NameFallback' }
            elseif ($componentMatch -and $pciMatch) { $method = 'ComponentID+PCI' }
            elseif ($componentMatch) { $method = 'ComponentID' }
            elseif ($pciMatch) { $method = 'PCI' }
            else { $method = 'MatchedByCompatibilityFunction' }
        }
        else {
            $hasComponent = Test-DriFTInstalledComponentIdIsValid -ComponentId $device.ComponentID
            $hasPci = Test-DriFTDeviceHasPciIdentity -Device $device

            if (-not $hasComponent -and -not $hasPci) { $reason = 'No valid componentID and no complete PCI identity' }
            elseif (-not $hasComponent) { $reason = 'No valid componentID; PCI identity did not match filtered catalog' }
            elseif (-not $hasPci) { $reason = 'componentID did not match filtered catalog; no complete PCI identity' }
            else { $reason = 'componentID and PCI identity did not match filtered catalog' }
        }

        $matches += @([PSCustomObject]@{
            Device          = $device
            CatalogRow      = $best
            Matched         = [bool]$best
            MatchMethod     = $method
            UnmatchedReason = $reason
        })
    }

    return @($matches)
}


function Find-DriFTCatalogNameFallbackMatch {
<#
.SYNOPSIS
    Finds a conservative catalog match by normalized device/catalog name.

.DESCRIPTION
    Used only after ComponentID and full PCI identity matching fail. This preserves
    accuracy by requiring the catalog rows to already be filtered to the current
    system/OS and requiring compatible ComponentType. This is needed for some 17G
    Redfish firmware rows, especially drives, that expose a useful product name but
    no catalog-ready ComponentID or complete PCI identity.
#>
    [CmdletBinding()]
    param(
        [AllowNull()]$Device,
        [AllowEmptyCollection()][object[]]$CatalogRows
    )

    if ($null -eq $Device) { return $null }

    $deviceType = Get-DriFTFirstNonEmpty $Device.ComponentType $Device.componentType
    $deviceNames = Get-DriFTComparableNameValues @(
        $Device.ElementName,
        $Device.Display,
        $Device.Name,
        $Device.RelatedItem
    )

    if (-not $deviceNames -or @($deviceNames).Count -eq 0) { return $null }

    $hits = @()

    foreach ($row in @($CatalogRows | Where-Object { $null -ne $_ })) {
        if (-not (Test-DriFTCatalogComponentTypeCompatible -CatalogObject $row -Device $Device)) { continue }

        $catalogName = Get-DriFTFirstNonEmpty $row.Name.Display.'#cdata-section' $row.Name.Display $row.name
        $catalogCategory = Get-DriFTFirstNonEmpty $row.LUCategory.value $row.LUCategory
        $catalogNames = Get-DriFTComparableNameValues @($catalogName, $catalogCategory)

        foreach ($dName in @($deviceNames)) {
            foreach ($cName in @($catalogNames)) {
                if ([string]::IsNullOrWhiteSpace($dName) -or [string]::IsNullOrWhiteSpace($cName)) { continue }

                # Exact normalized name or strong contains match only. Avoid broad short tokens.
                if (($dName.Length -ge 10 -and $cName.Length -ge 10) -and
                    (($dName -eq $cName) -or ($cName.Contains($dName)) -or ($dName.Contains($cName)))) {
                    $hits += @($row)
                    break
                }
            }
            if ($hits -contains $row) { break }
        }
    }

    return @($hits | Sort-Object { ConvertTo-DriFTSafeDateTime $_.releaseDate } | Select-Object -Last 1)
}

function Get-DriFTComparableNameValues {
<#
.SYNOPSIS
    Normalizes names for conservative device/catalog text fallback matching.
#>
    [CmdletBinding()]
    param([Parameter(ValueFromRemainingArguments = $true)]$Values)

    $out = @()

    foreach ($value in @($Values)) {
        foreach ($item in @($value)) {
            if ($null -eq $item) { continue }

            $text = ([string]$item)
            if ([string]::IsNullOrWhiteSpace($text)) { continue }

            if ($text -match '/') {
                $text = (($text.TrimEnd('/') -split '/')[-1])
            }

            $clean = $text `
                -replace '\s+Firmware Inventory$', '' `
                -replace '\s+Firmware$', '' `
                -replace '\s+Controller$', '' `
                -replace '\s+Adapter$', '' `
                -replace '\s+Device$', '' `
                -replace '\s+', ' '

            $clean = $clean.Trim().ToUpperInvariant()

            if ($clean.Length -ge 5) { $out += @($clean) }
        }
    }

    return @($out | Sort-Object -Unique)
}

function Find-DriFTCatalogCandidates {
<#
.SYNOPSIS
    Finds candidate catalog rows for a device from indexes.

.DESCRIPTION
    Returns a small candidate set before applying final compatibility checks.
    Uses a seen-key map to avoid duplicate candidate expansion when component
    and PCI indexes point to the same SoftwareComponent.
#>
    [CmdletBinding()]
    param(
        [AllowNull()]$Device,
        [Parameter(Mandatory)]$CatalogIndex
    )

    if ($null -eq $Device) { return @() }

    $rawCandidates = @()
    $type = Get-DriFTFirstNonEmpty $Device.ComponentType
    $componentId = Get-DriFTFirstNonEmpty $Device.ComponentID

    if (Test-DriFTInstalledComponentIdIsValid -ComponentId $componentId) {
        foreach ($key in @("$type|$componentId", $componentId)) {
            if ($CatalogIndex.ByTypeComponent.ContainsKey($key)) {
                foreach ($row in @($CatalogIndex.ByTypeComponent[$key])) { $rawCandidates += @($row) }
            }

            if ($CatalogIndex.ByComponent.ContainsKey($key)) {
                foreach ($row in @($CatalogIndex.ByComponent[$key])) { $rawCandidates += @($row) }
            }
        }
    }

    $pciKey = New-DriFTPciKey -VendorID $Device.VendorID -DeviceID $Device.DeviceID -SubVendorID $Device.SubVendorID -SubDeviceID $Device.SubDeviceID
    if ($pciKey) {
        foreach ($key in @("$type|$pciKey", $pciKey)) {
            if ($CatalogIndex.ByTypePci.ContainsKey($key)) {
                foreach ($row in @($CatalogIndex.ByTypePci[$key])) { $rawCandidates += @($row) }
            }

            if ($CatalogIndex.ByPci.ContainsKey($key)) {
                foreach ($row in @($CatalogIndex.ByPci[$key])) { $rawCandidates += @($row) }
            }
        }
    }

    $seen = @{}
    $candidates = @()

    foreach ($row in @($rawCandidates | Where-Object { $null -ne $_ })) {
        $rowKey = Get-DriFTCatalogRowKey -CatalogRow $row
        if ([string]::IsNullOrWhiteSpace($rowKey)) {
            $rowKey = [string]([Runtime.CompilerServices.RuntimeHelpers]::GetHashCode($row))
        }

        if (-not $seen.ContainsKey($rowKey)) {
            $seen[$rowKey] = $true
            $candidates += @($row)
        }
    }

    return @($candidates | Sort-Object path, vendorVersion, releaseDate -Unique)
}

function Test-DriFTCatalogDeviceMatch {
<#
.SYNOPSIS
    Tests whether a catalog device row matches an installed device.

.DESCRIPTION
    Final authoritative match check after indexed candidate lookup.
#>
    [CmdletBinding()]
    param(
        [AllowNull()]$CatalogDevice,
        [AllowNull()]$Device
    )

    if (-not $CatalogDevice -or $null -eq $Device) { return $false }

    if (-not (Test-DriFTCatalogComponentTypeCompatible -CatalogObject $CatalogDevice -Device $Device)) {
        return $false
    }

    $componentMatch = $false
    if (Test-DriFTInstalledComponentIdIsValid -ComponentId $Device.ComponentID) {
        $componentMatch = Test-DriFTCatalogComponentIdMatch -CatalogObject $CatalogDevice -ComponentId $Device.ComponentID
    }

    $pciMatch = $false
    if (Test-DriFTDeviceHasPciIdentity -Device $Device) {
        $pciMatch = Test-DriFTCatalogPciIdentityMatch -CatalogObject $CatalogDevice -Device $Device
    }

    return ($componentMatch -or $pciMatch)
}

function Test-DriFTCatalogComponentTypeCompatible {
<#
.SYNOPSIS
    Checks catalog/device component type compatibility.
#>
    [CmdletBinding()]
    param(
        [AllowNull()]$CatalogObject,
        [Parameter(Mandatory)]$Device
    )

    if (Test-DriFTDeviceIsBiosFirmware -Device $Device) { return $true }

    $catalogType = Get-DriFTCatalogComponentTypeValue -CatalogObject $CatalogObject
    $deviceType = Get-DriFTFirstNonEmpty $Device.ComponentType $Device.componentType

    if ([string]::IsNullOrWhiteSpace($catalogType) -or [string]::IsNullOrWhiteSpace($deviceType)) { return $true }
    return ($catalogType -ieq $deviceType)
}

function Test-DriFTDeviceIsBiosFirmware {
<#
.SYNOPSIS
    Detects BIOS inventory rows, including 17G BIOS reported as FRMW componentID 159.
#>
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Device)

    $componentId = Get-DriFTFirstNonEmpty $Device.ComponentID $Device.componentID
    $display = Get-DriFTFirstNonEmpty $Device.Display $Device.display
    $elementName = Get-DriFTFirstNonEmpty $Device.ElementName $Device.Name
    $relatedItem = Get-DriFTFirstNonEmpty $Device.RelatedItem

    if ($componentId -eq '159') { return $true }
    if ($display -imatch '^(Bios|BIOS|System BIOS|BIOS\.Setup)') { return $true }
    if ($elementName -imatch '^(BIOS|System BIOS)$') { return $true }
    if ($relatedItem -imatch '/Bios/?$') { return $true }

    return $false
}

function Test-DriFTInstalledComponentIdIsValid {
<#
.SYNOPSIS
    Determines if installed ComponentID can safely be used for matching.
#>
    [CmdletBinding()]
    param([AllowNull()]$ComponentId)

    $text = Get-DriFTFirstNonEmpty $ComponentId
    if ([string]::IsNullOrWhiteSpace($text)) { return $false }
    if ($text.Trim() -eq '0') { return $false }
    return $true
}

function Test-DriFTDeviceHasPciIdentity {
<#
.SYNOPSIS
    Checks whether a device has a complete PCI identity.
#>
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Device)

    return (
        -not [string]::IsNullOrWhiteSpace((Convert-DriFTHexId $Device.VendorID)) -and
        -not [string]::IsNullOrWhiteSpace((Convert-DriFTHexId $Device.DeviceID)) -and
        -not [string]::IsNullOrWhiteSpace((Convert-DriFTHexId $Device.SubVendorID)) -and
        -not [string]::IsNullOrWhiteSpace((Convert-DriFTHexId $Device.SubDeviceID))
    )
}

#endregion Matching Engine

#region Catalog Helpers

function Get-DriFTCatalogComponentTypeValue {
<#
.SYNOPSIS
    Gets ComponentType from a catalog object.
#>
    [CmdletBinding()]
    param([AllowNull()]$CatalogObject)

    return Get-DriFTFirstNonEmpty `
        $CatalogObject.ComponentType.value `
        $CatalogObject.ComponentType `
        $CatalogObject.componentType.value `
        $CatalogObject.componentType
}

function Get-DriFTCatalogComponentIdValues {
<#
.SYNOPSIS
    Gets all component ID values surfaced in a catalog object.
#>
    [CmdletBinding()]
    param([AllowNull()]$CatalogObject)

    $values = @()

    foreach ($obj in @($CatalogObject)) {
        if ($null -eq $obj) { continue }

        foreach ($candidate in @(
            $obj.ComponentID.value,
            $obj.componentID.value,
            $obj.ComponentID,
            $obj.componentID,
            $obj.SupportedDevices.Device.ComponentID.value,
            $obj.SupportedDevices.Device.componentID.value,
            $obj.SupportedDevices.Device.ComponentID,
            $obj.SupportedDevices.Device.componentID
        )) {
            foreach ($value in @($candidate)) {
                if (-not [string]::IsNullOrWhiteSpace([string]$value)) {
                    $values += @(([string]$value).Trim())
                }
            }
        }
    }

    return @($values | Sort-Object -Unique)
}

function Test-DriFTCatalogComponentIdMatch {
<#
.SYNOPSIS
    Tests exact component ID match.
#>
    [CmdletBinding()]
    param(
        [AllowNull()]$CatalogObject,
        [AllowNull()]$ComponentId
    )

    $componentIdText = Get-DriFTFirstNonEmpty $ComponentId
    if (-not (Test-DriFTInstalledComponentIdIsValid -ComponentId $componentIdText)) { return $false }

    foreach ($candidate in Get-DriFTCatalogComponentIdValues -CatalogObject $CatalogObject) {
        if ($candidate -ieq $componentIdText.Trim()) { return $true }
    }

    return $false
}

function Get-DriFTCatalogPciInfoObjects {
<#
.SYNOPSIS
    Gets PCIInfo rows from a catalog object.

.DESCRIPTION
    Uses plain PowerShell arrays instead of generic lists to avoid PS 5.1 XML node
    type-conversion issues while indexing Catalog.xml.
#>
    [CmdletBinding()]
    param([AllowNull()]$CatalogObject)

    $rows = @()

    foreach ($obj in @($CatalogObject)) {
        if ($null -eq $obj) { continue }

        foreach ($pci in @($obj.PCIInfo)) {
            if ($null -ne $pci) { $rows += @($pci) }
        }

        foreach ($device in @($obj.SupportedDevices.Device)) {
            foreach ($pci in @($device.PCIInfo)) {
                if ($null -ne $pci) { $rows += @($pci) }
            }
        }

        if ($obj.vendorID -or $obj.deviceID -or $obj.subVendorID -or $obj.subDeviceID) {
            $rows += @([PSCustomObject]@{
                vendorID    = Get-DriFTFirstNonEmpty $obj.vendorID.value $obj.vendorID
                deviceID    = Get-DriFTFirstNonEmpty $obj.deviceID.value $obj.deviceID
                subVendorID = Get-DriFTFirstNonEmpty $obj.subVendorID.value $obj.subVendorID
                subDeviceID = Get-DriFTFirstNonEmpty $obj.subDeviceID.value $obj.subDeviceID
            })
        }
    }

    return @($rows)
}

function Test-DriFTCatalogPciIdentityMatch {
<#
.SYNOPSIS
    Tests PCI identity match against catalog PCIInfo.
#>
    [CmdletBinding()]
    param(
        [AllowNull()]$CatalogObject,
        [Parameter(Mandatory)]$Device
    )

    $deviceKey = New-DriFTPciKey -VendorID $Device.VendorID -DeviceID $Device.DeviceID -SubVendorID $Device.SubVendorID -SubDeviceID $Device.SubDeviceID
    if (-not $deviceKey) { return $false }

    foreach ($pci in Get-DriFTCatalogPciInfoObjects -CatalogObject $CatalogObject) {
        $catalogKey = New-DriFTPciKey -VendorID $pci.vendorID -DeviceID $pci.deviceID -SubVendorID $pci.subVendorID -SubDeviceID $pci.subDeviceID
        if ($catalogKey -and $catalogKey -eq $deviceKey) { return $true }
    }

    return $false
}

function New-DriFTPciKey {
<#
.SYNOPSIS
    Creates a normalized full PCI identity key.
#>
    [CmdletBinding()]
    param(
        [AllowNull()]$VendorID,
        [AllowNull()]$DeviceID,
        [AllowNull()]$SubVendorID,
        [AllowNull()]$SubDeviceID
    )

    $v = Convert-DriFTHexId $VendorID
    $d = Convert-DriFTHexId $DeviceID
    $sv = Convert-DriFTHexId $SubVendorID
    $sd = Convert-DriFTHexId $SubDeviceID

    if ([string]::IsNullOrWhiteSpace($v) -or
        [string]::IsNullOrWhiteSpace($d) -or
        [string]::IsNullOrWhiteSpace($sv) -or
        [string]::IsNullOrWhiteSpace($sd)) {
        return $null
    }

    return "$v|$d|$sv|$sd"
}

#endregion Catalog Helpers

#region Report Generation

function New-DriFTReportRows {
<#
.SYNOPSIS
    Converts match objects to report rows.
#>
    [CmdletBinding()]
    param(
        [AllowNull()][AllowEmptyCollection()][object[]]$Matches,
        [Parameter(Mandatory)]$System,
        [Parameter(Mandatory)]$OperatingSystem,
        [Parameter(Mandatory)]$Context
    )

    $rows = @()

    foreach ($match in @($Matches | Where-Object { $null -ne $_ })) {
        if (-not $match.Matched) { continue }
        if ($null -eq $match.CatalogRow) { continue }
        if ($null -eq $match.Device) { continue }

        $catalog = $match.CatalogRow
        $device = $match.Device

        $installed = Compare-DriFTVersionForReport -InstalledVersion $device.Version -AvailableVersion $catalog.vendorVersion

        $row = New-DriFTReportRow `
            -ServiceTag $System.ServiceTag `
            -PowerEdge $System.PowerEdge `
            -OS "$($OperatingSystem.DisplayName) $($OperatingSystem.Version)".Trim() `
            -Type (Get-DriFTFirstNonEmpty $device.ComponentType) `
            -Category $catalog.LUCategory.value `
            -Name $catalog.Name.Display.'#cdata-section' `
            -InstalledVersion $installed `
            -AvailableVersion $catalog.vendorVersion `
            -CatalogInfo (Get-DriFTCatalogSourceInfo -CatalogRow $catalog) `
            -Criticality (ConvertTo-DriFTCriticality $catalog.Criticality.Display.'#cdata-section') `
            -ReleaseDate (($catalog.dateTime -split 'T')[0]) `
            -URL ($script:DriFTDellDownloadRoot + $catalog.path) `
            -Details $catalog.ImportantInfo.URL `
            -SourceType $System.SourceType

        if ($null -ne $row) {
            $rows += @($row)
        }
    }

    $deduped = @()
    foreach ($group in @($rows | Group-Object ServiceTag,PowerEdge,OS,Type,Category,Name,AvailableVersion,URL)) {
        $best = @($group.Group | Sort-Object @{ Expression = {
            try { [version]($_.InstalledVersion -replace '^__DRIFT_OUTDATED__','') }
            catch { [version]'0.0' }
        } } | Select-Object -Last 1)
        if ($best) { $deduped += @($best) }
    }

    return @($deduped | Sort-Object Type,Category,Name)
}



function New-DriFTUnmatchedPciReportRows {
<#
.SYNOPSIS
    Creates manual-check rows for installed PCI devices not found in the catalog.

.DESCRIPTION
    Some 17G Redfish/viewer.html collections expose valid PCI identity data for a
    device, but the filtered Dell catalog may not contain a matching SoftwareComponent
    or PCIInfo entry. Rather than silently dropping those devices, this adds an INFO
    row to the report so the PCI identity and all useful discovered device details
    can be manually checked against Dell Catalog.xml, ASHCI-Catalog.xml, BCG, or
    other sources.
#>
    [CmdletBinding()]
    param(
        [AllowNull()][AllowEmptyCollection()][object[]]$Matches,
        [Parameter(Mandatory)]$System,
        [Parameter(Mandatory)]$OperatingSystem,
        [Parameter(Mandatory)]$Context
    )

    $rows = @()
    $seen = @{}

    foreach ($match in @($Matches | Where-Object { $null -ne $_ })) {
        if ($match.Matched) { continue }

        $device = $match.Device
        if ($null -eq $device) { continue }

        $pciKey = New-DriFTPciKey `
            -VendorID $device.VendorID `
            -DeviceID $device.DeviceID `
            -SubVendorID $device.SubVendorID `
            -SubDeviceID $device.SubDeviceID

        if ([string]::IsNullOrWhiteSpace($pciKey)) { continue }

        $vendorId    = Convert-DriFTHexId $device.VendorID
        $deviceId    = Convert-DriFTHexId $device.DeviceID
        $subVendorId = Convert-DriFTHexId $device.SubVendorID
        $subDeviceId = Convert-DriFTHexId $device.SubDeviceID

        $deviceName = Get-DriFTFirstNonEmpty `
            $device.ElementName `
            $device.Display `
            $device.RelatedItem `
            $device.Name `
            'Unknown PCI Device'

        $partNumber = Get-DriFTFirstNonEmpty `
            $device.PartNumber `
            $device.PartNumberString `
            $device.DevicePartNumber `
            $device.SparePartNumber `
            $device.MMPartNumber `
            $device.PartID `
            $device.ComponentID `
            'Not Available'

        $description = Get-DriFTFirstNonEmpty `
            $device.Description `
            $device.DeviceDescription `
            $device.LongDescription `
            $device.ProductName `
            $device.Model `
            $device.Caption `
            $device.ElementName `
            $device.Display `
            'Not Available'

        $dedupeKey = (@($System.ServiceTag, $vendorId, $deviceId, $subVendorId, $subDeviceId, $deviceName, $partNumber) -join '|').ToLowerInvariant()
        if ($seen.ContainsKey($dedupeKey)) { continue }
        $seen[$dedupeKey] = $true

        $pciText = "VID=$vendorId; DID=$deviceId; SVID=$subVendorId; SSID=$subDeviceId"

        # Preserve the original DellSoftwareInventory fields when they were carried
        # from the PCI identity source. Do not let normalized placeholders like
        # ElementName=NIC.Slot or ComponentID=0 replace the useful source values.
        $originalElementName = Get-DriFTFirstNonEmpty $device.OriginalElementName $device.Description
        $originalId = Get-DriFTFirstNonEmpty $device.OriginalId
        $identityInfoValue = Get-DriFTFirstNonEmpty $device.IdentityInfoValue ("DCIM:firmware:{0}:{1}:{2}:{3}" -f $vendorId,$deviceId,$subVendorId,$subDeviceId)

        $detailPairs = [ordered]@{
            'Reason' = (Get-DriFTFirstNonEmpty $match.UnmatchedReason 'PCI identity was present but no matching catalog row was found')

            'ElementName' = (Get-DriFTFirstNonEmpty `
                $originalElementName `
                $device.ElementName `
                $device.Display `
                'Not Available')

            'Id' = (Get-DriFTFirstNonEmpty `
                $originalId `
                $device.Id `
                'Not Available')

            'IdentityInfoValue' = $identityInfoValue
        }

        $catalogInfo = @(
            'PCI info could not be found in catalog(s). Manual check needed.'
            ($detailPairs.GetEnumerator() | ForEach-Object { '"{0}":"{1}"' -f $_.Key, $_.Value })
        ) -join ','

        $row = New-DriFTReportRow `
            -ServiceTag $System.ServiceTag `
            -PowerEdge $System.PowerEdge `
            -OS "$($OperatingSystem.DisplayName) $($OperatingSystem.Version)".Trim() `
            -Type 'INFO' `
            -Category 'PCI Manual Check' `
            -Name "$deviceName - PCI device not found in catalog(s)" `
            -InstalledVersion (Get-DriFTFirstNonEmpty $device.Version 'Not Available') `
            -AvailableVersion 'No Catalog Match' `
            -CatalogInfo $catalogInfo `
            -Criticality 'Manual Check' `
            -ReleaseDate '' `
            -URL '' `
            -Details '' `
            -SourceType $System.SourceType

        if ($row) { $rows += @($row) }
    }

    if (@($rows).Count -gt 0) {
        Write-DriFTLog -Context $Context -Message "Unmatched PCI manual-check rows added: $(@($rows).Count)" -Level Warn -Indent 1
    }

    return @($rows | Sort-Object Category,Name)
}


function ConvertTo-DriFTClusterReportRows {
<#
.SYNOPSIS
    Converts normal report rows into cluster comparison rows.

.DESCRIPTION
    Legacy DriFT cluster mode groups updates by Type/Name/AvailableVersion and creates
    one InstalledVersion column per ServiceTag. This makes node-to-node drift obvious.
#>
    [CmdletBinding()]
    param(
        [AllowEmptyCollection()][object[]]$Rows,
        [string[]]$ServiceTags
    )

    $serviceTags = @($ServiceTags | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique)
    if (@($serviceTags).Count -le 1) { return @($Rows) }

    $out = @()

    $groups = @($Rows | Where-Object { $null -ne $_ } | Group-Object Type,Name,AvailableVersion,URL,Details,Criticality,ReleaseDate,CatalogInfo)

    foreach ($group in $groups) {
        $first = @($group.Group)[0]
        if ($null -eq $first) { continue }

        $ordered = [ordered]@{
            Type             = $first.Type
            Name             = if ($first.Details) { New-DriFTHtmlLink -Url $first.Details -Text $first.Name } else { ConvertTo-DriFTHtmlText $first.Name }
        }

        foreach ($tag in $serviceTags) {
            $rowForTag = @($group.Group | Where-Object { (($_.ServiceTag -replace '\*','') -eq ($tag -replace '\*','')) } | Select-Object -First 1)
            if ($rowForTag) {
                $ordered[$tag] = $rowForTag.InstalledVersion
            }
            else {
                $ordered[$tag] = 'NA'
            }
        }

        $ordered['AvailableVersion'] = if ($first.URL) { New-DriFTHtmlLink -Url $first.URL -Text $first.AvailableVersion } else { ConvertTo-DriFTHtmlText $first.AvailableVersion }
        $ordered['Criticality']      = $first.Criticality
        $ordered['ReleaseDate']      = $first.ReleaseDate
        $ordered['CatalogInfo']      = $first.CatalogInfo

        $out += @([PSCustomObject]$ordered)
    }

    return @($out | Sort-Object Type,Name)
}

function Test-DriFTShouldUseClusterReport {
<#
.SYNOPSIS
    Determines whether to use cluster comparison report layout.
#>
    [CmdletBinding()]
    param(
        [AllowEmptyCollection()][object[]]$Rows,
        [Parameter(Mandatory)]$Context
    )

    $tags = @($Rows | Where-Object { $_.ServiceTag } | ForEach-Object { $_.ServiceTag -replace '\*','' } | Sort-Object -Unique)
    if (@($tags).Count -gt 1) { return $true }

    return $false
}


function Write-DriFTHtmlReport {
<#
.SYNOPSIS
    Writes the DriFT HTML report.

.DESCRIPTION
    Keeps report rendering isolated. The table shape remains compatible with the
    current DriFT report columns.
#>
    [CmdletBinding()]
    param(
        [AllowEmptyCollection()][object[]]$Rows,
        [Parameter(Mandatory)]$Context,
        [Parameter(Mandatory)]$CatalogSet
    )

    $outputRoot = if ($Context.OutputRoot) { $Context.OutputRoot } else { $PWD.Path }

    # Prefer writing reports beside the source TSR/SupportAssist ZIP.
    if ($Context.CurrentCollection -and $Context.CurrentCollection.SourcePath) {
        $candidateRoot = Split-Path -Parent $Context.CurrentCollection.SourcePath

        if ($candidateRoot -and (Test-Path -LiteralPath $candidateRoot)) {
            $outputRoot = $candidateRoot
        }
    }
    $tagPart = (@(@($Context.ServiceTags)) | Sort-Object -Unique) -join '_'
    $path = Join-Path $outputRoot "$($Context.Version)_$($Context.DateStamp)_$tagPart.html"
    $path = $path.Replace('*','')

    if ($path.Length -gt 248) {
        $path = Join-Path $outputRoot "$($Context.Version)_$($Context.DateStamp).html"
    }

    if (-not $Rows -or @($Rows).Count -eq 0) {
        $Rows = @([PSCustomObject]@{
            ServiceTag = ''
            PowerEdge = ''
            OS = ''
            Type = 'INFO'
            Category = 'No Report Rows'
            Name = 'No firmware, driver, OS, or config rows were generated.'
            InstalledVersion = 'Not Available'
            AvailableVersion = 'Not Available'
            CatalogInfo = 'DriFT'
            Criticality = 'Informational'
            ReleaseDate = ''
            URL = ''
            Details = ''
        })
    }

    $header = Get-DriFTHtmlHeader
    $generatedAt = Get-Date
    $reportTags = (@(@($Context.ServiceTags)) | Sort-Object -Unique) -join ', '
    if ([string]::IsNullOrWhiteSpace($reportTags)) { $reportTags = 'Not Available' }

    $reportOs = Get-DriFTFirstNonEmpty `
        ($Rows | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.OS) } | Select-Object -First 1 -ExpandProperty OS) `
        $OperatingSystem.DisplayName `
        $OperatingSystem.RawName `
        'NO OS Detected in TSR Data'

    $reportOsVersion = Get-DriFTFirstNonEmpty $OperatingSystem.Version
    if (-not [string]::IsNullOrWhiteSpace($reportOsVersion) -and $reportOs -notmatch [regex]::Escape($reportOsVersion)) {
        $reportOs = "$reportOs $reportOsVersion".Trim()
    }

    $pre = @"
<div class="drift-shell">
  <header class="drift-topbar">
    <div class="drift-brand">
      <div class="drift-logo">DriFT</div>
      <div>
        <div class="drift-title">v$($Context.DisplayVer)</div>
        <div class="drift-subtitle">Driver & Firmware Tool</div>
      </div>
    </div>
    <div class="drift-toplinks">
    </div>
  </header>

  <section class="drift-hero">
    <div>
      <div class="drift-eyebrow">Dell Support Report</div>
      <h1>Driver and firmware drift summary</h1>
      <p>Review applicable firmware, driver, catalog, and OS update recommendations for the selected SupportAssist collection.</p>
    </div>
    <div class="drift-meta-card">
      <div class="drift-meta-row"><span>Generated</span><strong>$generatedAt</strong></div>
      <div class="drift-meta-row"><span>Service Tag(s)</span><strong>$reportTags</strong></div>
      <div class="drift-meta-row"><span>Platform</span><strong>$($Context.Platform.Type)</strong></div>
      <div class="drift-meta-row"><span>Model</span><strong>$($System.PowerEdge)</strong></div>
      <div class="drift-meta-row"><span>Operating System</span><strong>$reportOs</strong></div>
    </div>
  </section>

  <section class="drift-legend">
    <div class="drift-legend-card">
      <span class="drift-status-dot drift-red"></span>
      <div><strong>Red InstalledVersion</strong><br><span>Installed version is lower than available version.</span></div>
    </div>
    <div class="drift-legend-card">
      <span class="drift-status-dot drift-yellow"></span>
      <div><strong>Not Available</strong><br><span>Installed version was not contained in the SupportAssist collection.</span></div>
    </div>
    <div class="drift-legend-card">
      <span class="drift-status-dot drift-blue"></span>
      <div><strong>CatalogInfo</strong><br><span>Shows the catalog source and version/date used for the row.</span></div>
    </div>
  </section>
"@


    $isClusterReport = Test-DriFTShouldUseClusterReport -Rows $Rows -Context $Context

    if ($isClusterReport) {
        $clusterTags = @($Rows |
            Where-Object { $_.ServiceTag } |
            ForEach-Object { $_.ServiceTag -replace '\*','' } |
            Sort-Object -Unique)

        $htmlRows = ConvertTo-DriFTClusterReportRows -Rows $Rows -ServiceTags $clusterTags
    }
    else {
        $htmlRows = $Rows |
            Sort-Object ServiceTag,PowerEdge,Type,Category |
            Select-Object Type,Category,Name,InstalledVersion,AvailableVersion,CatalogInfo,Criticality,ReleaseDate,
                @{Label='Documentation';Expression={ New-DriFTHtmlLink -Url $_.Details -Text 'Link' }},
                @{Label='Download Link';Expression={ New-DriFTHtmlLink -Url $_.URL -Text $_.URL }}
    }

    $post = "<footer class='drift-catalog-footer'>$($CatalogSet.CatVerInfo)</footer></div>"

    $html = $htmlRows | ConvertTo-Html -Head $header -PreContent $pre -PostContent $post

    $html = $html -replace '&gt;','>' -replace '&lt;','<' -replace '&#39;', "'" `
        -replace '<td>NA</td>','<td style="background-color: #ffff00">Not Available</td>' `
        -replace '<td>Not Applicable</td>','<td style="background-color: #ffff00">Not Available</td>' 

    $html = $html -replace '<td>__DRIFT_OUTDATED__([^<]+)</td>', '<td style="color: #ffffff; background-color: #ff0000">$1</td>'

    $html | Out-File -FilePath $path -Encoding UTF8

    return $path
}

function Get-DriFTHtmlHeader {
<#
.SYNOPSIS
    Returns DriFT report CSS/header.
#>
    [CmdletBinding()]
    param()

@"
<style TYPE="text/css">
:root {
    --dell-blue: #0672cb;
    --dell-blue-dark: #01447e;
    --dell-blue-soft: #eaf5ff;
    --dell-cyan: #0d98ba;
    --dell-ink: #0e0e0e;
    --dell-text: #2b2b2b;
    --dell-muted: #636363;
    --dell-border: #d7d7d7;
    --dell-bg: #f5f6f7;
    --dell-card: #ffffff;
    --dell-red: #bb2a33;
    --dell-yellow: #fff4cc;
    --dell-row: #fafafa;
}

* {
    box-sizing: border-box;
}

body {
    font-family: "Segoe UI", Arial, Helvetica, sans-serif;
    color: var(--dell-text);
    background: var(--dell-bg);
    margin: 0;
    padding: 0;
}

.drift-shell {
    max-width: 1720px;
    margin: 0 auto;
    padding: 0 28px 32px 28px;
}

.drift-topbar {
    background: #ffffff;
    border-bottom: 1px solid var(--dell-border);
    min-height: 72px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding: 14px 28px;
    margin: 0 -28px;
}

.drift-brand {
    display: flex;
    align-items: center;
    gap: 14px;
}

.drift-logo {
    width: 48px;
    height: 48px;
    border-radius: 50%;
    border: 2px solid var(--dell-blue);
    color: var(--dell-blue);
    display: flex;
    align-items: center;
    justify-content: center;
    font-weight: 800;
    letter-spacing: -0.04em;
}

.drift-title {
    font-size: 20px;
    font-weight: 650;
    color: var(--dell-ink);
}

.drift-subtitle {
    font-size: 13px;
    color: var(--dell-muted);
}

.drift-toplinks {
    display: flex;
    align-items: center;
    gap: 26px;
    color: #4d4d4d;
    font-size: 14px;
}

.drift-hero {
    background: linear-gradient(90deg, #ffffff 0%, #f4faff 100%);
    border: 1px solid var(--dell-border);
    border-top: 4px solid var(--dell-blue);
    border-radius: 2px;
    margin-top: 24px;
    padding: 30px;
    display: grid;
    grid-template-columns: minmax(320px, 1fr) 420px;
    gap: 28px;
    box-shadow: 0 2px 8px rgba(0,0,0,0.08);
}

.drift-eyebrow {
    color: var(--dell-blue);
    font-size: 13px;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 0.05em;
    margin-bottom: 8px;
}

.drift-hero h1 {
    margin: 0 0 10px 0;
    color: var(--dell-ink);
    font-size: 34px;
    font-weight: 620;
    letter-spacing: -0.03em;
}

.drift-hero p {
    margin: 0;
    color: #4a4a4a;
    max-width: 820px;
    font-size: 16px;
    line-height: 1.5;
}

.drift-meta-card {
    background: #ffffff;
    border: 1px solid var(--dell-border);
    border-radius: 2px;
    box-shadow: 0 1px 4px rgba(0,0,0,0.08);
}

.drift-meta-row {
    display: grid;
    grid-template-columns: 120px 1fr;
    gap: 14px;
    padding: 13px 16px;
    border-bottom: 1px solid #ededed;
    font-size: 13px;
}

.drift-meta-row:last-child {
    border-bottom: 0;
}

.drift-meta-row span {
    color: var(--dell-muted);
}

.drift-meta-row strong {
    color: var(--dell-ink);
    font-weight: 600;
}

.drift-legend {
    display: grid;
    grid-template-columns: repeat(3, minmax(220px, 1fr));
    gap: 16px;
    margin: 18px 0 22px 0;
}

.drift-legend-card {
    background: #ffffff;
    border: 1px solid var(--dell-border);
    border-radius: 2px;
    padding: 16px;
    display: flex;
    gap: 12px;
    align-items: flex-start;
    box-shadow: 0 1px 4px rgba(0,0,0,0.06);
}

.drift-legend-card strong {
    color: var(--dell-ink);
    font-size: 14px;
}

.drift-legend-card span {
    color: var(--dell-muted);
    font-size: 13px;
}

.drift-status-dot {
    width: 12px;
    height: 12px;
    min-width: 12px;
    border-radius: 50%;
    margin-top: 4px;
    display: inline-block;
}

.drift-status-dot.drift-red { background: var(--dell-red); }
.drift-status-dot.drift-yellow { background: #e6ac00; }
.drift-status-dot.drift-blue { background: var(--dell-blue); }

table {
    margin-left: 0 !important;
    width: 100%;
    width: 100%;
    margin: 0 0 30px 0;
    border-collapse: collapse;
    background: var(--dell-card);
    border: 1px solid var(--dell-border);
    box-shadow: 0 2px 8px rgba(0,0,0,0.08);
}

th {
    cursor: pointer;
    user-select: none;
    position: sticky;
    top: 0;
    z-index: 2;
    background: #f2f2f2;
    color: var(--dell-ink);
    border-bottom: 2px solid var(--dell-border);
    padding: 12px 10px;
    text-align: left;
    font-size: 12px;
    font-weight: 700;
    text-transform: none;
    white-space: nowrap;
}

th::after {
    content: " ⇅";
    color: #8a8a8a;
    font-size: 10px;
    font-weight: 400;
}

th.sort-asc::after {
    content: " ▲";
    color: var(--dell-blue);
}

th.sort-desc::after {
    content: " ▼";
    color: var(--dell-blue);
}

td {
    border-bottom: 1px solid #e6e6e6;
    padding: 10px;
    font-size: 13px;
    line-height: 1.35;
    vertical-align: top;
}

tr:nth-child(even) td {
    background-color: var(--dell-row);
}

tr:hover td {
    background-color: var(--dell-blue-soft);
}

a {
    color: var(--dell-blue);
    text-decoration: none;
    font-weight: 600;
}

a:hover {
    text-decoration: underline;
}

td[style*="background-color: #ff0000"],
td[style*="background-color:#ff0000"] {
    background-color: var(--dell-red) !important;
    color: #ffffff !important;
    font-weight: 700;
}

td[style*="background-color: #ffff00"],
td[style*="background-color:#ffff00"] {
    background-color: var(--dell-yellow) !important;
    color: #5c4600 !important;
    font-weight: 700;
}

.drift-cluster-note {
    color: var(--dell-muted);
    font-size: 13px;
    margin: -6px 28px 18px 28px;
}

td:nth-child(12) a,
td:nth-child(13) a {
    display: inline-block;
    border: 1px solid var(--dell-blue);
    border-radius: 2px;
    padding: 5px 9px;
    background: #ffffff;
}

td:nth-child(12) a:hover,
td:nth-child(13) a:hover {
    background: var(--dell-blue);
    color: #ffffff;
    text-decoration: none;
}

body > br,
body > a {
    margin-left: 28px;
}

body > br:last-of-type {
    display: none;
}

.drift-catalog-footer {
    width: calc(100% - 56px);
    margin: 0 28px 30px 28px;
    background: #ffffff;
    border: 1px solid var(--dell-border);
    border-left: 4px solid var(--dell-blue);
    padding: 16px 18px;
    color: var(--dell-muted);
    font-size: 13px;
    box-shadow: 0 1px 4px rgba(0,0,0,0.06);
}

.tooltip {
    position: relative;
    display: inline-block;
    border-bottom: 1px dotted black;
}

.tooltip .tooltiptext {
    visibility: hidden;
    width: 160px;
    background-color: #111827;
    color: #fff;
    text-align: center;
    border-radius: 4px;
    padding: 8px;
    position: absolute;
    z-index: 10;
}

.tooltip:hover .tooltiptext {
    visibility: visible;
}

@media (max-width: 1000px) {
    .drift-hero {
        grid-template-columns: 1fr;
    }

    .drift-legend {
        grid-template-columns: 1fr;
    }

    .drift-toplinks {
        display: none;
    }

    table {
        display: block;
        overflow-x: auto;
    }
}

.report-table,
.report-table-container,
.table-responsive,
.drift-table-wrapper,
.drift-table-container {
    margin-left: 0 !important;
    padding-left: 0 !important;
    width: 100% !important;
    box-sizing: border-box;
}

.report-table table,
.table-responsive table,
.drift-table-wrapper table,
.drift-table-container table {
    margin-left: 0 !important;
    width: 100% !important;
}

</style>

<script type="text/javascript">
document.addEventListener("DOMContentLoaded", function () {
    const table = document.querySelector("table");
    if (!table) { return; }

    const headers = Array.from(table.querySelectorAll("th"));

    function normalizeText(value) {
        return (value || "").replace(/\s+/g, " ").trim();
    }

    function criticalityRank(value) {
        const v = normalizeText(value).toLowerCase();
        if (v.includes("urgent")) { return 1; }
        if (v.includes("recommended")) { return 2; }
        if (v.includes("optional")) { return 3; }
        if (v.includes("not available") || v.includes("no data")) { return 4; }
        return 5;
    }

    function parseVersion(value) {
        const text = normalizeText(value);
        const matches = text.match(/\d+|[a-zA-Z]+/g);
        if (!matches) { return null; }

        return matches.map(part => {
            const n = Number(part);
            return Number.isNaN(n) ? part.toLowerCase() : n;
        });
    }

    function compareVersionLike(a, b) {
        const av = parseVersion(a);
        const bv = parseVersion(b);
        if (!av || !bv) { return null; }

        const max = Math.max(av.length, bv.length);
        for (let i = 0; i < max; i++) {
            const x = av[i] ?? 0;
            const y = bv[i] ?? 0;

            if (typeof x === "number" && typeof y === "number") {
                if (x !== y) { return x - y; }
            }
            else {
                const sx = String(x);
                const sy = String(y);
                if (sx !== sy) { return sx.localeCompare(sy); }
            }
        }

        return 0;
    }

    function getCellValue(row, index) {
        const cell = row.children[index];
        return cell ? normalizeText(cell.innerText || cell.textContent) : "";
    }

    function sortTable(index, direction) {
        const tbody = table.tBodies[0] || table;
        const rows = Array.from(tbody.querySelectorAll("tr")).filter(row => row.querySelectorAll("td").length > 0);
        const headerText = normalizeText(headers[index].innerText).toLowerCase();

        rows.sort((rowA, rowB) => {
            const a = getCellValue(rowA, index);
            const b = getCellValue(rowB, index);

            let result = 0;

            if (headerText === "criticality") {
                result = criticalityRank(a) - criticalityRank(b);
            }
            else if (headerText.includes("date")) {
                const da = Date.parse(a);
                const db = Date.parse(b);
                if (!Number.isNaN(da) && !Number.isNaN(db)) {
                    result = da - db;
                }
                else {
                    result = a.localeCompare(b, undefined, { numeric: true, sensitivity: "base" });
                }
            }
            else if (headerText.includes("version")) {
                const versionCompare = compareVersionLike(a, b);
                result = versionCompare === null ? a.localeCompare(b, undefined, { numeric: true, sensitivity: "base" }) : versionCompare;
            }
            else {
                result = a.localeCompare(b, undefined, { numeric: true, sensitivity: "base" });
            }

            return direction === "asc" ? result : -result;
        });

        headers.forEach(h => h.classList.remove("sort-asc", "sort-desc"));
        headers[index].classList.add(direction === "asc" ? "sort-asc" : "sort-desc");

        rows.forEach(row => tbody.appendChild(row));
    }

    headers.forEach((header, index) => {
        header.setAttribute("title", "Click to sort");
        header.addEventListener("click", function () {
            const current = header.getAttribute("data-sort-direction") || "none";
            const next = current === "asc" ? "desc" : "asc";

            headers.forEach(h => h.removeAttribute("data-sort-direction"));
            header.setAttribute("data-sort-direction", next);

            sortTable(index, next);
        });
    });

    const criticalityIndex = headers.findIndex(h => normalizeText(h.innerText).toLowerCase() === "criticality");
    if (criticalityIndex >= 0) {
        headers[criticalityIndex].setAttribute("data-sort-direction", "asc");
        sortTable(criticalityIndex, "asc");
    }
});
</script>

<title>DriFT Report</title>
"@
}

#endregion Report Generation

#region Supplemental Capability Placeholders


function Add-DriFTPrecisionWorkstationDriverRows {
<#
.SYNOPSIS
    Adds Precision workstation driver rows not represented by firmware inventory.

.DESCRIPTION
    Precision TSR firmware inventory normally does not include installed OS driver
    versions. For first-pass workstation parity, this adds latest applicable client
    driver rows from the filtered Precision catalog and marks InstalledVersion as
    Not Available.
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Inventory,
        [Parameter(Mandatory)]$CatalogRows,
        [Parameter(Mandatory)]$System,
        [Parameter(Mandatory)]$OperatingSystem,
        [Parameter(Mandatory)]$Context
    )

    if (-not (Test-DriFTPrecisionSystem -System $System)) { return @() }
    if ($OperatingSystem.Family -ne 'Windows') { return @() }

    $rows = @()
    $allCatalogRows = @($CatalogRows.AllRows | Where-Object { $null -ne $_ })

    $driverCategories = @(
        'Chipset',
        'Video',
        'Network',
        'Audio',
        'Serial ATA',
        'Storage',
        'Security',
        'Mouse, Keyboard & Input Devices',
        'Systems Management'
    )

    foreach ($category in $driverCategories) {
        $driver = @($allCatalogRows | Where-Object {
            (Get-DriFTCatalogComponentTypeValue -CatalogObject $_) -ieq 'DRVR' -and
            $_.LUCategory.value -ieq $category
        } | Sort-Object { ConvertTo-DriFTSafeDateTime $_.releaseDate } | Select-Object -Last 1)

        if ($driver) {
            $rows += @(New-DriFTReportRowFromCatalog `
                -CatalogRow $driver `
                -System $System `
                -OperatingSystem $OperatingSystem `
                -Type 'DRVR' `
                -InstalledVersion 'NA' `
                -CatalogInfo '' `
                -SourceType $System.SourceType)
        }
    }

    $rows = @($rows | Sort-Object Type,Category,Name,URL -Unique)

    Write-DriFTLog -Context $Context -Message "Precision supplemental driver rows added: $(@($rows).Count)" -Level Info -Indent 1

    return @($rows)
}


function Add-DriFTWindowsDriverRows {
<#
.SYNOPSIS
    Adds Windows driver rows not always represented by installed firmware inventory.

.DESCRIPTION
    Adds supplemental Windows driver rows for devices discovered in inventory. 17G
    Redfish firmware inventory contains firmware versions but not installed Windows
    driver versions, so these rows intentionally show InstalledVersion as NA/Not
    Available, matching the legacy DriFT behavior.
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Inventory,
        [Parameter(Mandatory)]$CatalogRows,
        [Parameter(Mandatory)]$System,
        [Parameter(Mandatory)]$OperatingSystem,
        [Parameter(Mandatory)]$Context
    )

    if (Test-DriFTPrecisionSystem -System $System) {
        return @(Add-DriFTPrecisionWorkstationDriverRows `
            -Inventory $Inventory `
            -CatalogRows $CatalogRows `
            -System $System `
            -OperatingSystem $OperatingSystem `
            -Context $Context)
    }

    $rows = @()
    $allCatalogRows = @($CatalogRows.AllRows | Where-Object { $null -ne $_ })

    if ($OperatingSystem.Family -ne 'Windows') { return @() }
    if (-not $OperatingSystem.DriverSupport) { return @() }

    # Chipset driver: legacy DriFT adds the newest non-USB chipset driver for Windows.
    $chipsetDriver = @($allCatalogRows | Where-Object {
        (Get-DriFTCatalogComponentTypeValue -CatalogObject $_) -ieq 'DRVR' -and
        $_.LUCategory.value -ieq 'Chipset' -and
        $_.Name.Display.'#cdata-section' -inotmatch 'USB'
    } | Sort-Object { ConvertTo-DriFTSafeDateTime $_.releaseDate } | Select-Object -Last 1)

    if ($chipsetDriver) {
        $rows += @(New-DriFTReportRowFromCatalog `
            -CatalogRow $chipsetDriver `
            -System $System `
            -OperatingSystem $OperatingSystem `
            -Type 'DRVR' `
            -InstalledVersion 'NA' `
            -CatalogInfo '' `
            -SourceType $System.SourceType)
    }

    # Network drivers
    $installedNetworkDevices = @($Inventory | Where-Object {
        $_ -and (
            $_.Display -imatch 'NIC|Ethernet|FastLinQ|QLogic|Intel|Mellanox|ConnectX|Network' -or
            $_.ElementName -imatch 'NIC|Ethernet|FastLinQ|QLogic|Intel|Mellanox|ConnectX|Network' -or
            $_.RelatedItem -imatch 'NIC|Ethernet|Network|Mellanox|ConnectX'
        )
    })

    foreach ($device in $installedNetworkDevices) {
        $driver = Find-DriFTBestDriverCatalogRow `
            -Device $device `
            -CatalogRows $allCatalogRows `
            -Category 'Network' `
            -OperatingSystem $OperatingSystem

        if ($driver) {
            $rows += @(New-DriFTReportRowFromCatalog `
                -CatalogRow $driver `
                -System $System `
                -OperatingSystem $OperatingSystem `
                -Type 'DRVR' `
                -InstalledVersion 'NA' `
                -CatalogInfo '' `
                -SourceType $System.SourceType)
        }
    }

    # SAS RAID driver: add only when a RAID/PERC/HBA/BOSS/AHCI-like controller is present.
    $installedRaidDevices = @($Inventory | Where-Object {
        $_ -and (
            $_.Display -imatch 'RAID|PERC|HBA|BOSS|AHCI|H9[0-9]{2}|HBA[0-9]' -or
            $_.ElementName -imatch 'RAID|PERC|HBA|BOSS|AHCI|H9[0-9]{2}|HBA[0-9]' -or
            $_.RelatedItem -imatch 'RAID|PERC|HBA|BOSS|AHCI|Storage'
        )
    })

    if ($installedRaidDevices) {
        $raidDriver = $null

        foreach ($device in $installedRaidDevices) {
            $raidDriver = Find-DriFTBestDriverCatalogRow `
                -Device $device `
                -CatalogRows $allCatalogRows `
                -Category 'SAS RAID' `
                -OperatingSystem $OperatingSystem

            if ($raidDriver) { break }
        }

        if (-not $raidDriver) {
            # Some RAID driver catalog rows do not expose matching PCI identity.
            # Legacy DriFT falls back to latest SAS RAID driver for the filtered model catalog.
            $raidDriver = @($allCatalogRows | Where-Object {
                (Get-DriFTCatalogComponentTypeValue -CatalogObject $_) -ieq 'DRVR' -and
                $_.LUCategory.value -ieq 'SAS RAID'
            } | Sort-Object { ConvertTo-DriFTSafeDateTime $_.releaseDate } | Select-Object -Last 1)
        }

        if ($raidDriver) {
            $rows += @(New-DriFTReportRowFromCatalog `
                -CatalogRow $raidDriver `
                -System $System `
                -OperatingSystem $OperatingSystem `
                -Type 'DRVR' `
                -InstalledVersion 'NA' `
                -CatalogInfo '' `
                -SourceType $System.SourceType)
        }
    }

    return @($rows | Sort-Object URL -Unique | Sort-Object Type,Category,Name)
}

function Find-DriFTBestDriverCatalogRow {
<#
.SYNOPSIS
    Finds the best Windows driver catalog row for an installed device.
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Device,
        [AllowEmptyCollection()][object[]]$CatalogRows,
        [Parameter(Mandatory)][string]$Category,
        [Parameter(Mandatory)]$OperatingSystem
    )

    $candidates = @($CatalogRows | Where-Object {
        (Get-DriFTCatalogComponentTypeValue -CatalogObject $_) -ieq 'DRVR' -and
        ($_.LUCategory.value -ieq $Category)
    })

    # Prefer exact PCI identity when available.
    $pciMatches = @($candidates | Where-Object {
        Test-DriFTCatalogPciIdentityMatch -CatalogObject $_ -Device $Device
    })

    if ($pciMatches) {
        return @($pciMatches | Sort-Object { ConvertTo-DriFTSafeDateTime $_.releaseDate } | Select-Object -Last 1)
    }

    # Driver catalog entries sometimes only carry vendor/device or model-specific description.
    $deviceVendor = Convert-DriFTHexId $Device.VendorID
    $deviceId = Convert-DriFTHexId $Device.DeviceID

    if ($deviceVendor -and $deviceId) {
        $partialPciMatches = @($candidates | Where-Object {
            $hit = $false
            foreach ($pci in @(Get-DriFTCatalogPciInfoObjects -CatalogObject $_)) {
                if ((Convert-DriFTHexId $pci.vendorID) -eq $deviceVendor -and
                    (Convert-DriFTHexId $pci.deviceID) -eq $deviceId) {
                    $hit = $true
                    break
                }
            }
            $hit
        })

        if ($partialPciMatches) {
            return @($partialPciMatches | Sort-Object { ConvertTo-DriFTSafeDateTime $_.releaseDate } | Select-Object -Last 1)
        }
    }

    return $null
}

function New-DriFTReportRowFromCatalog {
<#
.SYNOPSIS
    Creates a DriFT report row directly from a catalog SoftwareComponent.
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$CatalogRow,
        [Parameter(Mandatory)]$System,
        [Parameter(Mandatory)]$OperatingSystem,
        [Parameter(Mandatory)][string]$Type,
        [Parameter(Mandatory)][string]$InstalledVersion,
        [AllowEmptyString()][string]$CatalogInfo,
        [string]$SourceType = 'Unknown'
    )

    $effectiveCatalogInfo = Get-DriFTFirstNonEmpty $CatalogInfo (Get-DriFTCatalogSourceInfo -CatalogRow $CatalogRow)

    return New-DriFTReportRow `
        -ServiceTag $System.ServiceTag `
        -PowerEdge $System.PowerEdge `
        -OS "$($OperatingSystem.DisplayName) $($OperatingSystem.Version)".Trim() `
        -Type $Type `
        -Category $CatalogRow.LUCategory.value `
        -Name $CatalogRow.Name.Display.'#cdata-section' `
        -InstalledVersion $InstalledVersion `
        -AvailableVersion $CatalogRow.vendorVersion `
        -CatalogInfo $effectiveCatalogInfo `
        -Criticality (ConvertTo-DriFTCriticality $CatalogRow.Criticality.Display.'#cdata-section') `
        -ReleaseDate (($CatalogRow.dateTime -split 'T')[0]) `
        -URL ($script:DriFTDellDownloadRoot + $CatalogRow.path) `
        -Details $CatalogRow.ImportantInfo.URL `
        -SourceType $SourceType
}

function Add-DriFTKbDownloadLinkType {
<#
.SYNOPSIS
    Adds the GetKBDLLink C# helper used to resolve Microsoft Catalog download URLs.
#>
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Context)

    if ('GetKBDLLink' -as [type]) { return }

    try {
        Add-Type @"
using System.Net;
using System.IO;
using System.Text.RegularExpressions;

public static class GetKBDLLink
{
    public static string GetDownloadLink(string KBNumber, string Product)
    {
        string kbGUID = "";
        string kbDLUriSource = "";

        var webRequest = WebRequest.Create("https://www.catalog.update.microsoft.com/Search.aspx?q=" + KBNumber);
        webRequest.Method = "GET";
        var webResponse = webRequest.GetResponse();
        var responseStream = webResponse.GetResponseStream();
        var streamReader = new StreamReader(responseStream);
        string responseContent = streamReader.ReadToEnd();

        var kbMatches = Regex.Matches(responseContent, @"id=(?:""|')(.*?)(?=_link)([^\/]*)");

        foreach (Match ItemMatch in kbMatches)
        {
            if (ItemMatch.Groups[2].Value.ToLower().Contains(Product.ToLower()))
            {
                kbGUID = ItemMatch.Groups[1].Value;
                break;
            }
        }

        if (string.IsNullOrWhiteSpace(kbGUID))
        {
            return "";
        }

        string post1 = "https://www.catalog.update.microsoft.com/DownloadDialog.aspx?updateIDs=[{%22size%22%3A0%2C%22languages%22%3A%22%22%2C%22uidInfo%22%3A%22";
        string post2 = "%22%2C%22updateID%22%3A%22";
        string post3 = "%22}]&updateIDsBlockedForImport=&wsusApiPresent=&contentImport=&sku=&serverName=&ssl=&portNumber=&version=";

        string postText = post1 + kbGUID + post2 + kbGUID + post3;

        webRequest = WebRequest.Create(postText);
        webRequest.Method = "GET";
        webResponse = webRequest.GetResponse();
        responseStream = webResponse.GetResponseStream();
        streamReader = new StreamReader(responseStream);
        responseContent = streamReader.ReadToEnd();

        kbDLUriSource = Regex.Match(responseContent, @"(?<=downloadInformation\[0\].files\[0\].url = '|"")(.*?)(?='|"";)").Groups[1].Value;

        return kbDLUriSource;
    }
}
"@ -ErrorAction Stop
    }
    catch {
        Write-DriFTLog -Context $Context -Message "Unable to load KB download link helper: $($_.Exception.Message)" -Level Warn -Indent 1
    }
}

function Get-DriFTMicrosoftCatalogProductFilter {
<#
.SYNOPSIS
    Returns the Microsoft Update Catalog product string used to choose a KB package.

.DESCRIPTION
    The catalog search can return multiple products for the same KB. The old DriFT
    logic used "server operating system" for Windows Server. Keep that as the
    default while allowing future special cases.
#>
    [CmdletBinding()]
    param([AllowNull()][string]$OsText)

    $text = [string]$OsText

    if ($text -imatch 'Azure Stack HCI|Azure Local|23H2|22H2|21H2|20H2') {
        return 'server operating system'
    }

    if ($text -imatch 'Windows Server|Server|2016|2019|2022|2025|2012|2008') {
        return 'server operating system'
    }

    return 'server operating system'
}

function Get-DriFTKbDownloadLink {
<#
.SYNOPSIS
    Resolves a direct Microsoft Update Catalog download URL for a KB.
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$KBNumber,
        [Parameter(Mandatory)][string]$Product,
        [Parameter(Mandatory)]$Context
    )

    if ([string]::IsNullOrWhiteSpace($KBNumber)) { return '' }

    try {
        Add-DriFTKbDownloadLinkType -Context $Context

        if (-not ('GetKBDLLink' -as [type])) { return '' }

        $download = [GetKBDLLink]::GetDownloadLink($KBNumber, $Product)

        if ([string]::IsNullOrWhiteSpace($download)) { return '' }

        return $download.Replace('&amp;', '&')
    }
    catch {
        Write-DriFTLog -Context $Context -Message "Microsoft Catalog download link lookup failed for ${KBNumber}: $($_.Exception.Message)" -Level Warn -Indent 1
        return ''
    }
}


function Get-DriFTWindowsUpdateHistoryUrl {
<#
.SYNOPSIS
    Maps detected Windows / Azure Local OS text to the Microsoft update history page.
#>
    [CmdletBinding()]
    param([AllowNull()][string]$OsText)

    $text = [string]$OsText

    if ($text -imatch '2025|26100')       { return 'https://support.microsoft.com/en-us/help/5047442' }
    if ($text -imatch '2022|20348')       { return 'https://support.microsoft.com/en-us/help/5005454' }
    if ($text -imatch '2019|17763')       { return 'https://support.microsoft.com/en-us/help/4464619' }
    if ($text -imatch '2016|14393')       { return 'https://support.microsoft.com/en-us/help/4000825' }
    if ($text -imatch '2012\s*R2|9600')   { return 'https://support.microsoft.com/en-us/help/4009470' }
    if ($text -imatch '2008\s*R2|7601')   { return 'https://support.microsoft.com/en-us/help/4009469' }

    if ($text -imatch '23H2|25398')       { return 'https://support.microsoft.com/en-us/help/5031680' }
    if ($text -imatch '22H2')             { return 'https://support.microsoft.com/en-us/help/5018894' }
    if ($text -imatch '21H2')             { return 'https://support.microsoft.com/en-us/help/5004047' }
    if ($text -imatch '20H2')             { return 'https://support.microsoft.com/en-us/help/4595086' }

    return $null
}

function Get-DriFTWindowsUpdateBuildToken {
<#
.SYNOPSIS
    Returns the OS build token used to filter Microsoft update history links.
#>
    [CmdletBinding()]
    param([AllowNull()][string]$OsText)

    $text = [string]$OsText

    if ($text -imatch '26100') { return '26100' }
    if ($text -imatch '25398') { return '25398' }
    if ($text -imatch '20348') { return '20348' }
    if ($text -imatch '17763') { return '17763' }
    if ($text -imatch '14393') { return '14393' }
    if ($text -imatch '9600')  { return '9600' }
    if ($text -imatch '7601')  { return '7601' }

    if ($text -imatch '2025')        { return '26100' }
    if ($text -imatch '23H2')        { return '25398' }
    if ($text -imatch '2022|21H2')   { return '20348' }
    if ($text -imatch '2019')        { return '17763' }
    if ($text -imatch '2016')        { return '14393' }
    if ($text -imatch '2012\s*R2')   { return '9600' }
    if ($text -imatch '2008\s*R2')   { return '7601' }

    return ''
}

function Get-DriFTLatestWindowsUpdateFromMicrosoft {
<#
.SYNOPSIS
    Scrapes the Microsoft update history page for the latest non-preview KB.
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Url,
        [AllowNull()][string]$BuildToken,
        [int]$KBItemsToShow = 1,
        [Parameter(Mandatory)]$Context
    )

    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

        $webClient = New-Object System.Net.WebClient
        $webClient.Proxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
        $htmlpage = $webClient.DownloadString($Url)

        if ([string]::IsNullOrWhiteSpace($htmlpage)) { return @() }

        $safeBuild = if ([string]::IsNullOrWhiteSpace($BuildToken)) { '\d{4,5}' } else { [regex]::Escape($BuildToken) }

        # Legacy DriFT-style pattern from 1.79.
        $legacyPattern = 'supLeftNavLink.*?(href=\".*?\")>(.*?)(KB\d{7})\D+((?:(?!Preview).)' + $safeBuild + '.*?)(?:\)|<)'
        $links = [regex]::Matches(
            $htmlpage,
            $legacyPattern,
            [System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [System.Text.RegularExpressions.RegexOptions]::Singleline
        )

        # Broader fallback for Microsoft page markup changes.
        if (@($links).Count -eq 0) {
            $fallbackPattern = '<a[^>]+href=\"(?<href>[^"]+)\"[^>]*>\s*(?<text>(?:(?!Preview).)*?(?<kb>KB\d{7})(?:(?!Preview).)*?' + $safeBuild + '.*?)</a>'
            $links = [regex]::Matches(
                $htmlpage,
                $fallbackPattern,
                [System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [System.Text.RegularExpressions.RegexOptions]::Singleline
            )
        }

        $rows = @()

        foreach ($link in @($links)) {
            $rawText = ''
            $kb = ''
            $href = ''
            $buildText = ''

            if ($link.Groups['kb'] -and $link.Groups['kb'].Success) {
                $kb = $link.Groups['kb'].Value
                $rawText = $link.Groups['text'].Value
                $href = $link.Groups['href'].Value
                $buildText = $rawText
            }
            else {
                $kb = $link.Groups[3].Value
                $rawText = $link.Groups[2].Value
                $buildText = $link.Groups[4].Value
                $href = (($link.Groups[1].Value -split 'href="')[-1] -split '"')[0]
            }

            if ([string]::IsNullOrWhiteSpace($kb)) { continue }
            if ($rawText -imatch 'Preview' -or $buildText -imatch 'Preview') { continue }

            $dateText = (($rawText -replace '&#x2014;', ' ') -replace '<.*?>',' ' -replace '\s+',' ').Trim()
            $buildClean = (($buildText -replace '<.*?>',' ') -replace '\s+',' ').Trim()
            $desc = (($dateText + ' ' + $kb + ' ' + $buildClean) -replace '\s+',' ').Trim()

            if ($href -notmatch '^https?://') {
                if ($href.StartsWith('/')) {
                    $href = "https://support.microsoft.com$href"
                }
                else {
                    $href = "https://support.microsoft.com/$href"
                }
            }

            $rows += @([PSCustomObject]@{
                KBNumber     = $kb
                Date         = $dateText
                Description  = $desc
                BuildText    = $buildClean
                InfoLink     = $href
                DownloadLink = ''
            })
        }

        return @($rows | Select-Object -First $KBItemsToShow)
    }
    catch {
        Write-DriFTLog -Context $Context -Message "Microsoft update history lookup failed: $($_.Exception.Message)" -Level Warn -Indent 1
        return @()
    }
}

function Add-DriFTWindowsUpdateRows {
<#
.SYNOPSIS
    Adds latest Windows cumulative update row.

.DESCRIPTION
    Dynamically scrapes the correct Microsoft update history page for the detected
    Windows Server / Azure Local / Azure Stack HCI OS and adds the latest non-preview
    KB row to the report.
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$System,
        [Parameter(Mandatory)]$OperatingSystem,
        [Parameter(Mandatory)]$Context
    )

    if ($OperatingSystem.Family -ne 'Windows') { return @() }

    $osText = @(
        $OperatingSystem.RawName,
        $OperatingSystem.DisplayName,
        $OperatingSystem.Version,
        $OperatingSystem.Build,
        $OperatingSystem.MajorVersion,
        $OperatingSystem.MinorVersion
    ) -join ' '

    $osText = [string]$osText

    $url = Get-DriFTWindowsUpdateHistoryUrl -OsText $osText
    if ([string]::IsNullOrWhiteSpace($url)) {
        Write-DriFTLog -Context $Context -Message "No Microsoft update history URL mapping found for OS: $osText" -Level Warn -Indent 1
        return @()
    }

    $buildToken = Get-DriFTWindowsUpdateBuildToken -OsText $osText
    $catalogProduct = Get-DriFTMicrosoftCatalogProductFilter -OsText $osText

    Write-DriFTLog -Context $Context -Message "Checking Microsoft update history: $url" -Level Info -Indent 1

    $kbRows = @(Get-DriFTLatestWindowsUpdateFromMicrosoft `
        -Url $url `
        -BuildToken $buildToken `
        -KBItemsToShow 1 `
        -Context $Context)

    if (@($kbRows).Count -eq 0) {
        return @(
            New-DriFTReportRow `
                -ServiceTag $System.ServiceTag `
                -PowerEdge $System.PowerEdge `
                -OS "$($OperatingSystem.DisplayName) $($OperatingSystem.Version)".Trim() `
                -Type 'OS' `
                -Category 'Microsoft Update' `
                -Name "Microsoft update history lookup returned no non-preview KB for $osText" `
                -InstalledVersion 'NA' `
                -AvailableVersion 'Not Available' `
                -CatalogInfo 'support.microsoft.com' `
                -Criticality 'Not Available' `
                -ReleaseDate '' `
                -URL $url `
                -Details $url `
                -SourceType $System.SourceType
        )
    }

    $rows = @()

    foreach ($kb in $kbRows) {
        $downloadLink = Get-DriFTFirstNonEmpty $kb.DownloadLink

        if ([string]::IsNullOrWhiteSpace($downloadLink)) {
            $downloadLink = Get-DriFTKbDownloadLink `
                -KBNumber $kb.KBNumber `
                -Product $catalogProduct `
                -Context $Context
        }

        $rows += @(
            New-DriFTReportRow `
                -ServiceTag $System.ServiceTag `
                -PowerEdge $System.PowerEdge `
                -OS "$($OperatingSystem.DisplayName) $($OperatingSystem.Version)".Trim() `
                -Type 'OS' `
                -Category 'Microsoft Update' `
                -Name $kb.Description `
                -InstalledVersion 'NA' `
                -AvailableVersion $kb.KBNumber `
                -CatalogInfo 'support.microsoft.com / catalog.update.microsoft.com' `
                -Criticality 'Not Available' `
                -ReleaseDate $kb.Date `
                -URL (Get-DriFTFirstNonEmpty $downloadLink $kb.InfoLink) `
                -Details $kb.InfoLink `
                -SourceType $System.SourceType
        )
    }

    return @($rows)
}


function Invoke-DriFTBroadcomIoCompatibility {
<#
.SYNOPSIS
    Queries the Broadcom Compatibility Guide IO endpoint.

.DESCRIPTION
    Tries several ESXi release label formats because Broadcom's filter values can
    vary between UI/API versions.
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$VendorID,
        [Parameter(Mandatory)][string]$DeviceID,
        [Parameter(Mandatory)][string]$SubVendorID,
        [Parameter(Mandatory)][string]$SubDeviceID,
        [Parameter(Mandatory)][string]$EsxiVersion
    )

    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    $releaseCandidates = @(
        "ESXi $EsxiVersion",
        "VMware ESXi $EsxiVersion",
        $EsxiVersion
    )

    if ($EsxiVersion -match '^(?<base>\d+\.\d+)\s+U(?<u>\d+)$') {
        $releaseCandidates += @(
            "ESXi $($Matches.base) Update $($Matches.u)",
            "VMware ESXi $($Matches.base) Update $($Matches.u)"
        )
    }

    $releaseCandidates = @($releaseCandidates | Sort-Object -Unique)

    foreach ($release in $releaseCandidates) {
        $body = @{
            programId = 'io'
            filters   = @(
                @{ displayKey = 'vid';                   filterValues = @($VendorID) },
                @{ displayKey = 'did';                   filterValues = @($DeviceID) },
                @{ displayKey = 'svid';                  filterValues = @($SubVendorID) },
                @{ displayKey = 'ssid';                  filterValues = @($SubDeviceID) },
                @{ displayKey = 'productReleaseVersion'; filterValues = @($release) }
            )
            keyword = @()
        } | ConvertTo-Json -Depth 8

        try {
            $result = Invoke-RestMethod `
                -Uri 'https://compatibilityguide.broadcom.com/compguide/programs/viewResults?limit=50&page=1&sortBy=&sortType=ASC' `
                -Method Post `
                -ContentType 'application/json' `
                -Body $body `
                -UseBasicParsing `
                -ErrorAction Stop

            $rows = @(Get-DriFTBcgResultRows -BcgResponse $result)
            if (@($rows).Count -gt 0) {
                return $result
            }
        }
        catch {
            $lastError = $_.Exception.Message
        }
    }

    return [PSCustomObject]@{
        success = $false
        data    = [PSCustomObject]@{ count = 0; fieldValues = @() }
        error   = $lastError
    }
}

function Get-DriFTBcgCompatibilityMatch {
<#
.SYNOPSIS
    Gets the best Broadcom Compatibility Guide match for an installed PCI device.

.DESCRIPTION
    Tries exact VendorID/DeviceID/SubVendorID/SubDeviceID first. If no match is
    returned, retries with SubVendorID and SubDeviceID swapped to preserve the legacy
    DriFT workaround for devices whose subsystem IDs are reported in reverse order.
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Device,
        [Parameter(Mandatory)][string]$EsxiVersion
    )

    $vendorId    = Convert-DriFTHexId $Device.VendorID
    $deviceId    = Convert-DriFTHexId $Device.DeviceID
    $subVendorId = Convert-DriFTHexId $Device.SubVendorID
    $subDeviceId = Convert-DriFTHexId $Device.SubDeviceID

    if ([string]::IsNullOrWhiteSpace($vendorId) -or
        [string]::IsNullOrWhiteSpace($deviceId) -or
        [string]::IsNullOrWhiteSpace($subVendorId) -or
        [string]::IsNullOrWhiteSpace($subDeviceId)) {
        return $null
    }

    $found = Invoke-DriFTBroadcomIoCompatibility `
        -VendorID $vendorId `
        -DeviceID $deviceId `
        -SubVendorID $subVendorId `
        -SubDeviceID $subDeviceId `
        -EsxiVersion $EsxiVersion

    if (@(Get-DriFTBcgResultRows -BcgResponse $found).Count -eq 0) {
        $found = Invoke-DriFTBroadcomIoCompatibility `
            -VendorID $vendorId `
            -DeviceID $deviceId `
            -SubVendorID $subDeviceId `
            -SubDeviceID $subVendorId `
            -EsxiVersion $EsxiVersion
    }

    return $found
}


function Get-DriFTBcgResultRows {
<#
.SYNOPSIS
    Normalizes Broadcom Compatibility Guide API response row shapes.
#>
    [CmdletBinding()]
    param([AllowNull()]$BcgResponse)

    if ($null -eq $BcgResponse) { return @() }

    foreach ($candidate in @(
        $BcgResponse.data.fieldValues,
        $BcgResponse.data.results,
        $BcgResponse.data.items,
        $BcgResponse.fieldValues,
        $BcgResponse.results,
        $BcgResponse.items,
        $BcgResponse.rows
    )) {
        $rows = @($candidate | Where-Object { $null -ne $_ })
        if (@($rows).Count -gt 0) { return @($rows) }
    }

    return @()
}



function Get-DriFTBcgRowText {
<#
.SYNOPSIS
    Flattens a Broadcom Compatibility Guide row to searchable text.
#>
    [CmdletBinding()]
    param([AllowNull()]$Row)

    if ($null -eq $Row) { return '' }

    $parts = @()

    try {
        foreach ($prop in @($Row.PSObject.Properties)) {
            if ($null -ne $prop.Value) {
                $parts += @([string]$prop.Name)
                $parts += @([string]$prop.Value)
            }
        }

        if ($Row.model) {
            foreach ($m in @($Row.model)) {
                foreach ($prop in @($m.PSObject.Properties)) {
                    if ($null -ne $prop.Value) { $parts += @([string]$prop.Value) }
                }
            }
        }

        if ($Row.hoverData) {
            foreach ($h in @($Row.hoverData)) {
                foreach ($prop in @($h.PSObject.Properties)) {
                    if ($null -ne $prop.Value) { $parts += @([string]$prop.Value) }
                }
            }
        }
    }
    catch { }

    return (($parts | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) -join ' | ')
}

function Get-DriFTBcgProductIdFromRow {
<#
.SYNOPSIS
    Extracts the published productId from a BCG result row.
#>
    [CmdletBinding()]
    param([AllowNull()]$Row)

    if ($null -eq $Row) { return '' }

    try {
        if ($Row.model -and $Row.model[0]) {
            $modelUrl = Get-DriFTFirstNonEmpty $Row.model[0].url
            if ($modelUrl -match 'productId=([^&]+)') { return $Matches[1] }
        }

        foreach ($prop in @($Row.PSObject.Properties)) {
            if ([string]$prop.Value -match 'productId=([^&]+)') { return $Matches[1] }
        }

        # Only use explicit productId fields, never generic row id.
        return Get-DriFTFirstNonEmpty `
            $Row.productId `
            $Row.productID `
            $Row.product.id `
            $Row.product.productId
    }
    catch {
        return ''
    }
}

function Get-DriFTBcgRowScore {
<#
.SYNOPSIS
    Scores a BCG result row against the installed device.

.DESCRIPTION
    Prevents the rewrite from picking the first returned BCG row when multiple
    products share the same Broadcom DeviceID. Exact subsystem identity and model
    text matches are preferred, which restores legacy DriFT behavior such as
    BCM57414 selecting productId 43268 instead of unrelated/unpublished IDs.
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Row,
        [Parameter(Mandatory)]$Device
    )

    $score = 0
    $text = Get-DriFTBcgRowText -Row $Row

    $vendorId    = Convert-DriFTHexId $Device.VendorID
    $deviceId    = Convert-DriFTHexId $Device.DeviceID
    $subVendorId = Convert-DriFTHexId $Device.SubVendorID
    $subDeviceId = Convert-DriFTHexId $Device.SubDeviceID

    foreach ($id in @($vendorId, $deviceId, $subVendorId, $subDeviceId)) {
        if ([string]::IsNullOrWhiteSpace($id)) { continue }

        if ($text -imatch "(^|[^0-9A-F])$([regex]::Escape($id))([^0-9A-F]|$)") {
            $score += 25
        }
    }

    $display = Get-DriFTFirstNonEmpty $Device.ElementName $Device.Display $Device.RelatedItem
    if (-not [string]::IsNullOrWhiteSpace($display)) {
        $tokens = @(([string]$display -split '[\s,;/\(\)\[\]\-]+' | Where-Object { $_.Length -ge 4 } | Sort-Object -Unique))
        foreach ($token in $tokens) {
            if ($text -imatch [regex]::Escape($token)) { $score += 3 }
        }
    }

    $productId = Get-DriFTBcgProductIdFromRow -Row $Row
    if (-not [string]::IsNullOrWhiteSpace($productId)) { $score += 10 }

    # Prefer rows with a published model URL because legacy DriFT derived the
    # correct detail links from model[0].url.
    try {
        $modelUrl = Get-DriFTFirstNonEmpty $Row.model[0].url
        if ($modelUrl -match 'productId=') { $score += 20 }
    }
    catch { }

    return $score
}

function Select-DriFTBestBcgResultRow {
<#
.SYNOPSIS
    Selects the best BCG row for one installed device.
#>
    [CmdletBinding()]
    param(
        [AllowEmptyCollection()][object[]]$Rows,
        [Parameter(Mandatory)]$Device
    )

    $scored = @()

    foreach ($row in @($Rows | Where-Object { $null -ne $_ })) {
        $score = Get-DriFTBcgRowScore -Row $row -Device $Device
        $productId = Get-DriFTBcgProductIdFromRow -Row $row

        $scored += @([PSCustomObject]@{
            Score     = $score
            ProductId = $productId
            Row       = $row
        })
    }

    # If rows are otherwise equal, choose the lowest numeric productId because
    # Broadcom often returns newer/unpublished duplicate product records later.
    return @($scored |
        Sort-Object `
            @{ Expression = 'Score'; Descending = $true },
            @{ Expression = {
                try { [int]$_.ProductId }
                catch { [int]::MaxValue }
            }; Descending = $false } |
        Select-Object -First 1 -ExpandProperty Row)
}



function Add-DriFTEsxiReleaseToBcgUrl {
<#
.SYNOPSIS
    Adds ESXi release selection parameters to a Broadcom Compatibility Guide detail URL.

.DESCRIPTION
    Broadcom detail pages default to the newest ESXi release. Adding
    productReleaseVersion and redirectFrom causes the page to open with the ESXi
    version detected from the TSR, such as ESXi 8.0 U3.
#>
    [CmdletBinding()]
    param(
        [AllowNull()][string]$Url,
        [AllowNull()][string]$EsxiVersion
    )

    if ([string]::IsNullOrWhiteSpace($Url)) { return $Url }
    if ([string]::IsNullOrWhiteSpace($EsxiVersion)) { return $Url }

    $cleanUrl = ([string]$Url).Replace('&amp;','&')

    if ($cleanUrl -match 'productReleaseVersion=') {
        return $cleanUrl
    }

    $releaseText = "ESXi $EsxiVersion"
    $encodedProductRelease = [System.Uri]::EscapeDataString("[$releaseText]")
    $encodedRedirect = [System.Uri]::EscapeDataString($releaseText)

    $separator = if ($cleanUrl.Contains('?')) { '&' } else { '?' }

    return "$cleanUrl${separator}productReleaseVersion=$encodedProductRelease&redirectFrom=$encodedRedirect"
}


function Get-DriFTBcgProductInfo {
<#
.SYNOPSIS
    Extracts product display values from a Broadcom Compatibility Guide row.
#>
    [CmdletBinding()]
    param([Parameter(Mandatory)]$BcgResult)

    $productId = Get-DriFTBcgProductIdFromRow -Row $BcgResult
    $model = 'No Data Found'
    $deviceType = 'No Data Found'
    $detailsLink = ''

    try {
        if ($BcgResult.model -and $BcgResult.model[0]) {
            $model = Get-DriFTFirstNonEmpty $BcgResult.model[0].name $model

            $modelUrl = Get-DriFTFirstNonEmpty $BcgResult.model[0].url
            if (-not [string]::IsNullOrWhiteSpace($modelUrl)) {
                if ($modelUrl -match '^https?://') {
                    $detailsLink = $modelUrl
                }
                elseif ($modelUrl.StartsWith('/')) {
                    $detailsLink = "https://compatibilityguide.broadcom.com$modelUrl"
                }
            }
        }

        if ($BcgResult.deviceType -and $BcgResult.deviceType[0].name) {
            $deviceType = $BcgResult.deviceType[0].name
        }
        elseif ($BcgResult.deviceTypes -and $BcgResult.deviceTypes[0].name) {
            $deviceType = $BcgResult.deviceTypes[0].name
        }
        elseif ($BcgResult.hoverData) {
            $hoverDeviceType = $BcgResult.hoverData |
                Where-Object { $_.displayName -eq 'Device Type' } |
                Select-Object -First 1

            if ($hoverDeviceType -and $hoverDeviceType.value) {
                $deviceType = $hoverDeviceType.value
            }
        }

        if ($model -eq 'No Data Found') {
            $model = Get-DriFTFirstNonEmpty `
                $BcgResult.modelName `
                $BcgResult.productName `
                $BcgResult.name `
                $BcgResult.displayName `
                $model
        }

        if ($deviceType -eq 'No Data Found') {
            $deviceType = Get-DriFTFirstNonEmpty `
                $BcgResult.deviceTypeName `
                $BcgResult.category `
                $deviceType
        }
    }
    catch { }

    if ([string]::IsNullOrWhiteSpace($detailsLink)) {
        if (-not [string]::IsNullOrWhiteSpace($productId)) {
            $detailsLink = "https://compatibilityguide.broadcom.com/detail?persona=live&productId=$productId&program=io"
        }
        else {
            $detailsLink = 'https://compatibilityguide.broadcom.com/search?program=io'
        }
    }

    if ($detailsLink -match 'productId=([^&]+)') {
        $productId = $Matches[1]
        $detailsLink = "https://compatibilityguide.broadcom.com/detail?persona=live&productId=$productId&program=io"
    }

    [PSCustomObject]@{
        ProductId   = $productId
        Model       = $model
        DeviceType  = $deviceType
        DetailsLink = $detailsLink
    }
}

function Get-DriFTVmwareCompatibilityDeviceRows {
<#
.SYNOPSIS
    Filters installed inventory to devices useful for VMware/Broadcom IO lookup.
#>
    [CmdletBinding()]
    param([AllowEmptyCollection()][object[]]$Inventory)

    return @($Inventory |
        Where-Object {
            $_ -and
            -not [string]::IsNullOrWhiteSpace((Convert-DriFTHexId $_.VendorID)) -and
            -not [string]::IsNullOrWhiteSpace((Convert-DriFTHexId $_.DeviceID)) -and
            -not [string]::IsNullOrWhiteSpace((Convert-DriFTHexId $_.SubVendorID)) -and
            -not [string]::IsNullOrWhiteSpace((Convert-DriFTHexId $_.SubDeviceID))
        } |
        Sort-Object VendorID,DeviceID,SubVendorID,SubDeviceID -Unique)
}


function Add-DriFTVmwareCompatibilityRows {
<#
.SYNOPSIS
    Adds VMware/vSAN Broadcom Compatibility Guide driver rows.

.DESCRIPTION
    Restores legacy DriFT behavior for ESXi/vSAN systems. The installed hardware
    PCI identity is sent to Broadcom Compatibility Guide and matching IO device rows
    are added to the report as DRVR rows with VMware/Broadcom support links.
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Inventory,
        [Parameter(Mandatory)]$System,
        [Parameter(Mandatory)]$OperatingSystem,
        [Parameter(Mandatory)]$Context
    )

    $osText = "$($OperatingSystem.RawName) $($OperatingSystem.DisplayName) $($OperatingSystem.Version)"
    if ($OperatingSystem.Family -ne 'VMware' -and $osText -notmatch 'VMware|ESXi|vSAN') {
        return @()
    }

    $esxiVersion = Get-DriFTEsxiVersionFromText -Text $OperatingSystem.Version
    if ([string]::IsNullOrWhiteSpace($esxiVersion)) {
        $esxiVersion = ConvertTo-DriFTVmwareVersion -OSName $OperatingSystem.RawName -OSVersion $OperatingSystem.Version
    }
    if ([string]::IsNullOrWhiteSpace($esxiVersion)) {
        $esxiVersion = ConvertTo-DriFTVmwareVersion -OSName $OperatingSystem.DisplayName -OSVersion ''
    }

    if ([string]::IsNullOrWhiteSpace($esxiVersion)) {
        Write-DriFTLog -Context $Context -Message 'VMware/vSAN compatibility lookup skipped: ESXi version could not be determined.' -Level Warn -Indent 1
        return @()
    }

    Write-DriFTLog -Context $Context -Message "Gathering VMware/vSAN supported driver versions from Broadcom Compatibility Guide for ESXi $esxiVersion..." -Level Info -Indent 1

    $rows = @()
    $devices = @(Get-DriFTVmwareCompatibilityDeviceRows -Inventory $Inventory)
    Export-DriFTDebugData -Context $Context -Name "$($System.ServiceTag)_VMware_BCG_InputDevices.csv" -InputObject $devices

    foreach ($device in $devices) {
        $found = Get-DriFTBcgCompatibilityMatch -Device $device -EsxiVersion $esxiVersion

        $bcgRows = @(Get-DriFTBcgResultRows -BcgResponse $found)

        if (@($bcgRows).Count -gt 0) {
            $bcgResult = Select-DriFTBestBcgResultRow -Rows $bcgRows -Device $device

            if (-not $bcgResult) { continue }

            $bcgInfo = Get-DriFTBcgProductInfo -BcgResult $bcgResult
            $bcgInfo.DetailsLink = Add-DriFTEsxiReleaseToBcgUrl -Url $bcgInfo.DetailsLink -EsxiVersion $esxiVersion

            Write-DriFTLog -Context $Context -Message ("BCG selected for {0}: productId={1}; model={2}; score={3}; ESXi={4}" -f (Get-DriFTFirstNonEmpty $device.ElementName $device.Display), $bcgInfo.ProductId, $bcgInfo.Model, (Get-DriFTBcgRowScore -Row $bcgResult -Device $device), $esxiVersion) -Level Info -Indent 2

            $rows += @(New-DriFTReportRow `
                -ServiceTag $System.ServiceTag `
                -PowerEdge $System.PowerEdge `
                -OS "$($OperatingSystem.DisplayName) $($OperatingSystem.Version)".Trim() `
                -Type 'DRVR' `
                -Category (Get-DriFTFirstNonEmpty $bcgInfo.DeviceType 'No Data Found') `
                -Name (Get-DriFTFirstNonEmpty $bcgInfo.Model $device.ElementName $device.Display 'No Data Found') `
                -InstalledVersion 'NA' `
                -AvailableVersion 'See Broadcom Compatibility Guide' `
                -CatalogInfo 'Broadcom Compatibility Guide' `
                -Criticality 'No Data Found' `
                -ReleaseDate 'No Data Found' `
                -URL $bcgInfo.DetailsLink `
                -Details $bcgInfo.DetailsLink `
                -SourceType $System.SourceType)
        }
    }

    $rows = @($rows | Sort-Object ServiceTag,Type,Category,Name,URL -Unique)

    Write-DriFTLog -Context $Context -Message "VMware/vSAN Broadcom Compatibility Guide rows added: $(@($rows).Count)" -Level Info -Indent 1
    if (@($rows).Count -eq 0) {
        Write-DriFTLog -Context $Context -Message "No BCG rows were returned. Check ESXi version parsed as '$esxiVersion' and PCI identities in normalized inventory." -Level Warn -Indent 1
    }

    return @($rows)
}

function Get-DriFTSelHealthRows {
<#
.SYNOPSIS
    Parses SEL warnings/errors from CurrentMBSel.txt.

.DESCRIPTION
    Port existing SEL parsing here. Keep it scoped to known SEL directories to avoid
    recursive PathTooLong issues in extracted 17G Redfish trees.
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Collection,
        [Parameter(Mandatory)]$System,
        [Parameter(Mandatory)]$Context
    )

    return @()
}

function Get-DriFTBiosAndIdracConfigRows {
<#
.SYNOPSIS
    Adds BIOS/iDRAC configuration compliance rows.

.DESCRIPTION
    Port existing Azure Stack HCI, Azure Stack Hub, and QLogic config checks here.
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Collection,
        [Parameter(Mandatory)]$System,
        [Parameter(Mandatory)]$OperatingSystem,
        [Parameter(Mandatory)]$Context
    )

    return @()
}

function Get-DriFTSwitchPortMapRows {
<#
.SYNOPSIS
    Builds CluChk switch-port-to-host map.

.DESCRIPTION
    Port existing SwitchPortConnectionID mapping here.
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Collection,
        [Parameter(Mandatory)]$System,
        [Parameter(Mandatory)]$Context
    )

    return @()
}

function Write-DriFTCluChkOutputs {
<#
.SYNOPSIS
    Writes CluChk supplemental XML outputs.

.DESCRIPTION
    Combines BIOS/iDRAC config rows, switch map rows, and SEL rows into the expected
    CluChk output file.
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Context,
        [AllowEmptyCollection()][object[]]$BiosConfigRows,
        [AllowEmptyCollection()][object[]]$SwitchMapRows,
        [AllowEmptyCollection()][object[]]$SelRows
    )

    if (-not $Context.FileNameGuid) { return }

    $outputRoot = if ($Context.OutputRoot) { $Context.OutputRoot } else { $PWD.Path }
    $path = Join-Path $outputRoot "$($Context.FileNameGuid)_BIOSandNICCFG.xml"

    @($BiosConfigRows + $SwitchMapRows + $SelRows) | Export-Clixml -Path $path
    Write-DriFTLog -Context $Context -Message "CluChk output written to: $path" -Level Success
}

#endregion Supplemental Capability Placeholders

#region Generic Helpers


function ConvertTo-DriFTSafeDateTime {
<#
.SYNOPSIS
    Safely converts catalog release dates for sorting.

.DESCRIPTION
    Catalog data can occasionally have missing or malformed date values. This helper
    prevents Sort-Object scriptblocks from throwing during matching or fallback
    selection.
#>
    [CmdletBinding()]
    param([AllowNull()]$Value)

    try {
        if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) {
            return [datetime]::MinValue
        }

        return [datetime]$Value
    }
    catch {
        return [datetime]::MinValue
    }
}

function ConvertTo-DriFTHtmlText {
<#
.SYNOPSIS
    HTML-encodes report text.

.DESCRIPTION
    Keeps report rendering safe if a catalog field or URL ever contains quotes,
    ampersands, or other HTML-sensitive characters.
#>
    [CmdletBinding()]
    param([AllowNull()]$Value)

    if ($null -eq $Value) { return '' }

    Add-Type -AssemblyName System.Web -ErrorAction SilentlyContinue
    return [System.Web.HttpUtility]::HtmlEncode([string]$Value)
}

function New-DriFTHtmlLink {
<#
.SYNOPSIS
    Creates a HTML hyperlink for the DriFT report.

.DESCRIPTION
    Prevents Broadcom Compatibility Guide URLs from displaying '&amp;' in the
    visible text while still remaining valid HTML links.
#>
    [CmdletBinding()]
    param(
        [AllowNull()][string]$Url,
        [AllowNull()][string]$Text
    )

    if ([string]::IsNullOrWhiteSpace($Url)) {
        return ''
    }

    $safeUrl = ([string]$Url).Replace('&amp;','&')

    $displayText = Get-DriFTFirstNonEmpty $Text $Url
    if ($null -eq $displayText) {
        $displayText = $safeUrl
    }

    $displayText = ([string]$displayText).Replace('&amp;','&')

    # URL display text should remain human-readable and not show HTML entities.
    if ($displayText -match '^https?://') {
        $safeText = $displayText
    }
    else {
        $safeText = ConvertTo-DriFTHtmlText $displayText
    }

    return "<a href='$safeUrl' target='_blank'>$safeText</a>"
}


function Get-DriFTFirstNonEmpty {
<#
.SYNOPSIS
    Returns the first non-empty value from a candidate list.
#>
    [CmdletBinding()]
    param([Parameter(ValueFromRemainingArguments = $true)]$Values)

    foreach ($value in $Values) {
        if ($null -eq $value) { continue }

        foreach ($item in @($value)) {
            if ($null -ne $item -and -not [string]::IsNullOrWhiteSpace([string]$item)) {
                return [string]$item
            }
        }
    }

    return $null
}

function Get-DriFTObjectProperty {
<#
.SYNOPSIS
    Safely reads a property from JSON/PSCustomObject/XML objects.

.DESCRIPTION
    Avoids PropertyNotFoundStrict errors when optional 17G metadata or Redfish
    fields are missing.
#>
    [CmdletBinding()]
    param(
        [AllowNull()]$InputObject,
        [Parameter(Mandatory)][string]$PropertyName
    )

    if ($null -eq $InputObject) { return $null }

    try {
        $prop = $InputObject.PSObject.Properties[$PropertyName]
        if ($prop) { return $prop.Value }
    }
    catch { }

    return $null
}


function Convert-DriFTHexId {
<#
.SYNOPSIS
    Normalizes PCI IDs to uppercase hex without 0x or separators.
#>
    [CmdletBinding()]
    param([AllowNull()]$Value)

    if ($null -eq $Value) { return '' }

    $text = ([string]$Value).Trim()
    if ([string]::IsNullOrWhiteSpace($text)) { return '' }

    if ($text -match '0x([0-9a-fA-F]+)') { $text = $Matches[1] }

    $text = $text -replace '[^0-9a-fA-F]', ''
    if ($text.Length -eq 0) { return '' }

    return $text.ToUpper()
}

function Get-DriFTCimPropertyValue {
<#
.SYNOPSIS
    Gets a property value from a CIM XML INSTANCE node.
#>
    [CmdletBinding()]
    param(
        [AllowNull()]$Instance,
        [Parameter(Mandatory)][string]$Name
    )

    if (-not $Instance) { return $null }

    return ($Instance.Property | Where-Object { $_.Name -eq $Name } | Select-Object -First 1).Value
}

function Get-DriFTCimArrayCurrentValue {
<#
.SYNOPSIS
    Gets CurrentValue from a CIM XML INSTANCE node.
#>
    [CmdletBinding()]
    param([AllowNull()]$Instance)

    if (-not $Instance) { return $null }

    $current = $Instance.'PROPERTY.ARRAY' |
        Where-Object { $_.Name -eq 'CurrentValue' } |
        Select-Object -First 1

    if ($current.'VALUE.ARRAY'.VALUE) {
        return @($current.'VALUE.ARRAY'.VALUE)[0]
    }

    return $null
}

function Get-DriFTMetadataJson {
<#
.SYNOPSIS
    Loads TSR metadata.json if present.
#>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Root)

    $path = Get-ChildItem -Path $Root -Filter 'metadata.json' -File -Recurse -Force -ErrorAction SilentlyContinue |
        Select-Object -First 1 -ExpandProperty FullName

    if (-not $path) { return $null }

    try { return Get-Content -Raw -Path $path | ConvertFrom-Json }
    catch { return $null }
}

function ConvertTo-DriFTServerModel {
<#
.SYNOPSIS
    Normalizes Dell model strings to catalog model names.

.DESCRIPTION
    Handles PowerEdge, AX, S2D Ready Node, Precision, XR2, XC, and R320/NX400 cases.
#>
    [CmdletBinding()]
    param([AllowNull()][string]$Model)

    if ([string]::IsNullOrWhiteSpace($Model)) { return '' }

    $serverType = $Model.Trim()

    if ($serverType -match 'XR2') { return 'R440' }

    if ($serverType -match 'AX-') {
        return ($serverType -replace 'AX-', 'R' -split '\s+')[0]
    }

    if ($serverType -match 'Storage Spaces Direct') {
        $serverType = $serverType -replace ' Storage Spaces Direct RN','' -replace ' Storage Spaces Direct R',''
    }

    if ($serverType -match 'Precision') {
        return $serverType
    }

    if ($serverType.Length -gt 4) {
        $parts = $serverType -split '\W'
        if (@($parts).Count -gt 1) { $serverType = $parts[1] }
    }

    if (($serverType -like 'XC*') -and ((([regex]::Match($serverType, '\d+').Groups[0].Value).Trim()).Length -eq 4)) {
        $serverType = $serverType -replace 'XC', 'C'
    }
    else {
        $serverType = $serverType -replace 'XC', 'R'
    }

    if ($serverType -eq 'R320') { $serverType = 'R320/NX400' }

    return $serverType
}

function Get-DriFTCatalogNeed {
<#
.SYNOPSIS
    Determines whether special catalogs are needed.

.DESCRIPTION
    Keeps HCI/Precision detection in one place.
#>
    [CmdletBinding()]
    param(
        [AllowNull()][string]$Model,
        [AllowNull()][string]$NormalizedModel
    )

    $special = 'NO'
    $s2d = 'NO'

    if ($Model -match 'Precision') {
        $special = 'Precision'
    }
    elseif ($Model -imatch 'AX|Azure Stack HCI|Storage Spaces Direct') {
        $special = 'HCI'
        $s2d = 'YES'
    }

    [PSCustomObject]@{
        SpecialCatalogNeeded = $special
        S2DCatalogNeeded     = $s2d
    }
}

function ConvertTo-DriFTOperatingSystemInfo {
<#
.SYNOPSIS
    Normalizes OS name/version to DriFT OS object.
#>
    [CmdletBinding()]
    param(
        [AllowNull()][string]$OSName,
        [AllowNull()][string]$OSVersion
    )

    if ([string]::IsNullOrWhiteSpace($OSName)) {
        return New-DriFTOperatingSystemInfo `
            -RawName '' `
            -Family 'Windows' `
            -DisplayName 'NO OS Detected in TSR Data: Assuming Windows 64bit' `
            -Version '' `
            -MajorVersion 6 `
            -MinorVersion 3 `
            -Build ''
    }

    if ($OSName -imatch 'VMware|ESXi|vSAN' -or $OSVersion -imatch 'VMware|ESXi|vSAN|\d+\.\d+\s*U\d+') {
        # vSAN Ready Node / vSAN OS strings should use the same VMware/Broadcom
        # driver and firmware compatibility path as ESXi.
        $vmw = ConvertTo-DriFTVmwareVersion -OSName $OSName -OSVersion $OSVersion

        return New-DriFTOperatingSystemInfo `
            -RawName $OSName `
            -Family 'VMware' `
            -DisplayName $OSName `
            -Version $vmw `
            -CatalogPackageType 'LW64' `
            -DriverSupport $true
    }

    $year = ''
    $major = ''
    $minor = ''
    $build = ''

    switch -Regex ($OSName) {
        '2008.*R2' { $year = '2008 R2'; $major = 6; $minor = 1; break }
        '2008'     { $year = '2008';    $major = 6; $minor = 0; break }
        '2012.*R2' { $year = '2012 R2'; $major = 6; $minor = 3; break }
        '2012'     { $year = '2012';    $major = 6; $minor = 2; break }
        '2016'     { $year = '2016';    $major = 10; $minor = 0; break }
        '2019'     { $year = '2019';    $major = 10; $minor = 17763; $build = '17763'; break }
        '2022'     { $year = '2022';    $major = 10; $minor = 0; $build = '20348'; break }
        '20H2'     { $year = '20H2';    $major = 10; $minor = 0; $build = '17784'; break }
        '21H2'     { $year = '21H2';    $major = 10; $minor = 0; $build = '20348'; break }
        '22H2'     { $year = '22H2';    $major = 10; $minor = 0; $build = '20349'; break }
        '23H2'     { $year = '23H2';    $major = 10; $minor = 0; $build = '25398'; break }
        default    { $year = $OSName;   $major = ''; $minor = ''; break }
    }

    return New-DriFTOperatingSystemInfo `
        -RawName $OSName `
        -Family 'Windows' `
        -DisplayName $OSName `
        -Version $build `
        -CatalogPackageType 'LW64' `
        -MajorVersion $major `
        -MinorVersion $minor `
        -Build $build `
        -DriverSupport $true
}


function Get-DriFTEsxiVersionFromText {
<#
.SYNOPSIS
    Extracts an ESXi/vSAN version string suitable for Broadcom Compatibility Guide.

.DESCRIPTION
    Handles strings such as:
      Dell-VMware ESXi 8.0 U3
      VMware ESXi 8.0 Update 3
      8.0.3
      7.0 U3
    and ignores image/profile strings such as Dell-ESXi when no version is present.
#>
    [CmdletBinding()]
    param([AllowNull()][string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) { return '' }

    $value = ([string]$Text).Trim()

    if ($value -match '(?<major>[678])\.(?<minor>[057])\s*(?:Update\s*|U\s*)?(?<update>[123])') {
        return "$($Matches.major).$($Matches.minor) U$($Matches.update)"
    }

    if ($value -match '(?<major>[678])\.(?<minor>[057])\.(?<patch>[123])') {
        return "$($Matches.major).$($Matches.minor) U$($Matches.patch)"
    }

    if ($value -match '(?<major>[678])\.(?<minor>[057])') {
        return "$($Matches.major).$($Matches.minor)"
    }

    return ''
}

function Get-DriFTBestEsxiVersion {
<#
.SYNOPSIS
    Finds the best ESXi/vSAN version from OS name/version fields.
#>
    [CmdletBinding()]
    param(
        [AllowNull()][string]$OSName,
        [AllowNull()][string]$OSVersion,
        [AllowNull()][string]$DisplayName
    )

    foreach ($candidate in @($OSName, $OSVersion, $DisplayName)) {
        $version = Get-DriFTEsxiVersionFromText -Text $candidate
        if (-not [string]::IsNullOrWhiteSpace($version)) { return $version }
    }

    return ''
}


function ConvertTo-DriFTVmwareVersion {
<#
.SYNOPSIS
    Normalizes VMware ESXi/vSAN versions for compatibility lookup.
#>
    [CmdletBinding()]
    param(
        [AllowNull()][string]$OSName,
        [AllowNull()][string]$OSVersion
    )

    $version = Get-DriFTBestEsxiVersion -OSName $OSName -OSVersion $OSVersion -DisplayName ''
    if (-not [string]::IsNullOrWhiteSpace($version)) { return $version }

    $value = Get-DriFTFirstNonEmpty $OSVersion $OSName
    if ([string]::IsNullOrWhiteSpace($value)) { return '' }

    $value = $value -replace 'VMware ', '' -replace 'ESXi ', '' -replace 'vSAN ', '' -replace 'VMware vSAN ', '' -replace ' Update ', ' U'
    if ($value -match 'Build') { $value = ($value -split 'Build')[0] }
    if ($value -match 'Patch') { $value = ($value -split ' Patch')[0] }
    if ($value -match 'GA') { $value = ($value -split 'GA ')[0] }

    $value = ($value -replace '\s{2,}', ' ').Trim()

    # Avoid using Dell image/profile names such as Dell-ESXi as the ESXi release.
    if ($value -notmatch '\d+\.\d+') { return '' }

    return $value
}

function Compare-DriFTVersionForReport {
<#
.SYNOPSIS
    Formats installed version for report output.

.DESCRIPTION
    If InstalledVersion is older than AvailableVersion, this function prefixes an
    internal marker. The HTML renderer converts that marker into a red cell style
    without displaying *** to the user.
#>
    [CmdletBinding()]
    param(
        [AllowNull()][string]$InstalledVersion,
        [AllowNull()][string]$AvailableVersion
    )

    $installed = Get-DriFTFirstNonEmpty $InstalledVersion
    $available = Get-DriFTFirstNonEmpty $AvailableVersion

    if ([string]::IsNullOrWhiteSpace($installed)) { return 'NA' }

    $installed = $installed -replace '^OSC_', ''

    if ([string]::IsNullOrWhiteSpace($available)) { return $installed }

    try {
        if ([version]$installed -lt [version]$available) {
            return "__DRIFT_OUTDATED__$installed"
        }
        return $installed
    }
    catch {
        if ($installed -lt $available) {
            return "__DRIFT_OUTDATED__$installed"
        }
        return $installed
    }
}

function Get-DriFT17GFqddVariants {
<#
.SYNOPSIS
    Builds FQDD variants for 17G Redfish identity correlation.
#>
    [CmdletBinding()]
    param([AllowNull()][string]$Value)

    $variants = New-Object System.Collections.Generic.List[string]
    if ([string]::IsNullOrWhiteSpace($Value)) { return @() }

    $base = $Value.Trim()
    [void]$variants.Add($base)

    $clean = $base `
        -replace '^DCIM[:_]CURRENT_0x23_', '' `
        -replace '^DCIM[:_]INSTALLED_0x23_', '' `
        -replace '^DCIM[:_]PREVIOUS_0x23_', '' `
        -replace '^DCIM_CURRENT#', '' `
        -replace '^DCIM_INSTALLED#', '' `
        -replace '^DCIM_PREVIOUS#', '' `
        -replace '^DCIM_CURRENT_', '' `
        -replace '^DCIM_INSTALLED_', '' `
        -replace '^DCIM_PREVIOUS_', ''

    if ($clean) { [void]$variants.Add($clean) }

    # DellSoftwareInventory IDs often look like:
    #   DCIM:INSTALLED_0x23_701__NIC.Slot.5-1-1
    # After the DCIM prefix is removed, strip the numeric inventory class prefix
    # so correlation can find the actual FQDD.
    $cleanFqdd = $clean -replace '^\d+__', ''
    if ($cleanFqdd -and $cleanFqdd -ne $clean) { [void]$variants.Add($cleanFqdd) }

    if ($clean -match '/') {
        [void]$variants.Add(($clean.TrimEnd('/') -split '/')[-1])
    }

    if ($clean -match '^(NIC\.Slot\.\d+)-') { [void]$variants.Add($Matches[1]) }
    if ($clean -match '^(RAID\.[^-\/]+\.\d+(?:-\d+)?)') { [void]$variants.Add($Matches[1]) }
    if ($clean -match '^(Disk\.[^-\/]+)') { [void]$variants.Add($Matches[1]) }
    if ($clean -match '^(PSU\.Slot\.\d+)') { [void]$variants.Add($Matches[1]) }

    return @($variants) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique
}

function Get-DriFT17GObjectKeys {
<#
.SYNOPSIS
    Builds searchable keys for a Redfish JSON object.
#>
    [CmdletBinding()]
    param(
        [AllowNull()]$JsonObject,
        [AllowNull()][string]$FilePath
    )

    $keys = New-Object System.Collections.Generic.List[string]

    foreach ($candidate in @(
        $JsonObject.'@odata.id',
        $JsonObject.Id,
        $JsonObject.Name,
        $JsonObject.FQDD,
        $JsonObject.SoftwareId,
        $JsonObject.SoftwareID,
        $JsonObject.DeviceId,
        $JsonObject.DeviceID,
        $JsonObject.FunctionId,
        $JsonObject.FunctionID,
        $JsonObject.Oem.Dell.DellNIC.FQDD,
        $JsonObject.Oem.Dell.DellNIC.Id,
        $JsonObject.Oem.Dell.DellNIC.InstanceID,
        $JsonObject.Oem.Dell.DellNIC.ProductName,
        $JsonObject.Oem.Dell.DellNIC.DeviceDescription,
        $JsonObject.Oem.Dell.DellPCIeFunction.FQDD,
        $JsonObject.Oem.Dell.DellPCIeFunction.Id,
        $JsonObject.Oem.Dell.DellPCIeFunction.InstanceID,
        $JsonObject.Oem.Dell.DellPCIeFunction.DeviceDescription
    )) {
        if ($candidate) { [void]$keys.Add(([string]$candidate).Trim()) }
    }

    if ($FilePath) {
        try {
            $parent = Split-Path -Path (Split-Path -Path $FilePath -Parent) -Leaf
            if ($parent) { [void]$keys.Add($parent) }
        } catch {}
    }

    foreach ($key in @($keys.ToArray())) {
        if ($key -match '/') { [void]$keys.Add(($key.TrimEnd('/') -split '/')[-1]) }
        if ($key -match '#') { [void]$keys.Add(($key -split '#')[-1]) }

        foreach ($variant in Get-DriFT17GFqddVariants -Value $key) {
            if ($variant) { [void]$keys.Add($variant) }
        }
    }

    return @($keys) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique
}

function Get-DriFT17GPreferredFqdd {
<#
.SYNOPSIS
    Chooses the best FQDD from a 17G Redfish JSON object.
#>
    [CmdletBinding()]
    param(
        [AllowNull()]$JsonObject,
        [AllowNull()][string]$FilePath
    )

    $idVariants = @(Get-DriFT17GFqddVariants -Value ([string]$JsonObject.Id))
    $folderVariants = @()

    if ($FilePath) {
        try {
            $folderVariants = @(Get-DriFT17GFqddVariants -Value (Split-Path -Path (Split-Path -Path $FilePath -Parent) -Leaf))
        } catch {}
    }

    $all = @(
        $JsonObject.FQDD,
        $JsonObject.Oem.Dell.FQDD,
        $JsonObject.Oem.Dell.DellNIC.FQDD,
        $JsonObject.Oem.Dell.DellNIC.Id,
        $JsonObject.Oem.Dell.DellPCIeFunction.FQDD,
        $JsonObject.Oem.Dell.DellPCIeFunction.Id
    ) + $idVariants + $folderVariants

    $preferred = @($all | Where-Object {
        $_ -match '^(NIC\.Slot\.\d+(?:-\d+-\d+)?)$' -or
        $_ -match '^(RAID\.[^\/]+)$' -or
        $_ -match '^(Disk\.[^\/]+)$' -or
        $_ -match '^(PSU\.Slot\.\d+)$'
    } | Sort-Object { $_.Length } -Descending | Select-Object -First 1)

    if ($preferred) { return $preferred }

    return Get-DriFTFirstNonEmpty $JsonObject.FQDD $JsonObject.Oem.Dell.FQDD $idVariants $folderVariants
}

#endregion Generic Helpers