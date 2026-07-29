#Requires -RunAsAdministrator

<#
.SYNOPSIS
    Standalone Dell SDDC diagnostic capture script.

.DESCRIPTION
    Converted from GetDellSDDC module (Invoke-GetDellSDDC)
    into a single self-contained .ps1 exposing the function Invoke-GetDellSDDC.

    Original: Microsoft Corporation, (c) 2016
    Module GUID: 7e0bc824-c371-4936-98e6-b7216ba5f348

.NOTES
    Usage:
      . .\Invoke-GetDellSDDC.ps1        # dot-source
      Invoke-GetDellSDDC                 # run with defaults
      Invoke-GetDellSDDC -ClusterName MyCluster -TemporaryPath C:\Temp\Diag

    Alternatively, uncomment the auto-execute block at the bottom.
#>

# CONVERSION: replaced $Module with script-scope variables
$script:ModuleName   = 'GetDellSDDC'
$script:ScriptVersion = '1.0.0'

###############################################################################
# region CommonFuncBlock — helpers shared with child jobs/sessions
###############################################################################

$CommonFuncBlock = {

    if (Get-Module -ListAvailable FailoverClusters) { Import-Module FailoverClusters }
    if (Get-Module -ListAvailable DcbQos)           { Import-Module DcbQos }
    if (Get-Module -ListAvailable Hyper-V)          { Import-Module Hyper-V }
    Import-Module CimCmdlets
    Import-Module NetAdapter
    Import-Module NetQos
    Import-Module SmbShare
    Import-Module SmbWitness
    Import-Module Storage

    function Copy-DirContentFromNode {
        param (
            [string[]] $Nodes,
            [string]   $PathOnNode,
            [string]   $SearchFilter = "*",
            [string]   $LocalDest
        )
        foreach ($NodeName in $Nodes) {
            $remotePath = "\\$NodeName\$(($PathOnNode -replace ':', '$'))"
            try {
                $items = Get-ChildItem -Path $remotePath -Filter $SearchFilter -Recurse -Depth 3 -Directory -ErrorAction Stop |
                    Sort-Object LastWriteTime -Descending | Select-Object -First 1
            } catch {
                Show-Warning "[$NodeName] Unable to access $($remotePath): $_"
                continue
            }
            if ($items) {
                $destPath = Join-Path -Path $LocalDest -ChildPath "Node_$NodeName\$($items.Name)"
                New-Item -Path $destPath -ItemType Directory -Force | Out-Null
                Show-Update "[$NodeName] Copying $($items.FullName) to $destPath"
                try {
                    Copy-Item -Path "$($items.FullName)\*" -Destination $destPath -Recurse -Force
                } catch {
                    Show-Warning "[$NodeName] Copy failed: $_"
                }
            } else {
                Show-Warning "[$NodeName] No matching folders found with filter '$SearchFilter'"
            }
        }
    }

    function Show-Error(
        [string] $Message,
        [System.Management.Automation.ErrorRecord] $e = $null
    ) {
        $Message = "$(get-date -format 's') : $Message - cmdlet was cancelled"
        if ($e) {
            Write-Error $Message
            throw $e
        } else {
            Write-Error $Message -ErrorAction Stop
        }
    }

    function Show-Warning(
        [string] $Message
    ) {
        Write-Warning "$(get-date -format 's') : $Message"
    }

    function Show-Update(
        [string] $Message,
        [System.ConsoleColor] $ForegroundColor = [System.ConsoleColor]::White
    ) {
        Write-Host -ForegroundColor $ForegroundColor "$(get-date -format 's') : $Message"
    }

    function Import-ClixmlIf(
        [string] $Path,
        [string] $MessageIf = $null
    ) {
        if (Test-Path $path) {
            Import-Clixml $path
        } else {
            $m = "$Path not present"
            if ($MessageIf) { $m = "$MessageIf : " + $m }
            Show-Warning $m
            $null
        }
    }

    function TimespanToString {
        param( [timespan] $TimeSpan )
        if ($TimeSpan.TotalDays -ge 1) {
            $TimeSpan.ToString("dd\d\.hh\h\:mm\m\:ss\.f\s")
        } elseif ($TimeSpan.TotalHours -ge 1) {
            $TimeSpan.ToString("hh\h\:mm\m\:ss\.f\s")
        } elseif ($TimeSpan.TotalMinutes -ge 1) {
            $TimeSpan.ToString("mm\m\:ss\.f\s")
        } else {
            $TimeSpan.ToString("ss\.f\s")
        }
    }

    function Show-JobRuntime(
        [object[]] $jobs,
        [hashtable] $namehash,
        [switch] $IncludeDone = $true,
        [switch] $IncludeRunning = $true
    ) {
        $job_running = @()
        $job_done = @()
        $jobs | sort Name,Location |% {
            $this = $_
            switch ($_.GetType().Name) {
                'PSRemotingJob' {
                    $jobname = $this.Name
                    $j = $this.ChildJobs | sort Location
                }
                'PSRemotingChildJob' {
                    if ($namehash.ContainsKey($this.Id)) {
                        $jobname = $namehash[$this.Id]
                    } else {
                        $jobname = ""
                    }
                    $j = $this
                }
                default { throw "unexpected job type $_" }
            }
            if ($IncludeDone) {
                $j |? State -ne Running |% {
                    $job_done += "$($_.State): $($jobname) [$($_.Name) $($_.Location)]: $(TimespanToString ($_.PSEndTime - $_.PSBeginTime)) : Start $($_.PSBeginTime.ToString('s')) - Stop $($_.PSEndTime.ToString('s'))"
                }
            }
            if ($IncludeRunning) {
                $t = get-date
                $j |? State -eq Running |% {
                    $job_running += "Running: $($jobname) [$($_.Name) $($_.Location)]: $(TimespanToString ($t - $_.PSBeginTime)) : Start $($_.PSBeginTime.ToString('s'))"
                    if (($t - $_.PSBeginTime).TotalMinutes -gt 60) {
                        Stop-Job -Name "$($_.Name)"
                        Write-Host "Job $jobname exceeded time limit" -ForegroundColor Yellow
                    }
                }
            }
        }
        if ($job_running.Count) { $job_running |% { Show-Update $_ } }
        if ($job_done.Count)    { $job_done |% { Show-Update $_ } }
    }

    function Show-WaitChildJob(
        [object[]] $jobs,
        [int] $tick = 5
    ) {
        $jhash = @{}
        $jobs |% {
            $j = $_
            $j.ChildJobs |% { $jhash[$_.Id] = $j.Name }
        }
        $tout_c = $tick
        $ttick = get-date
        $jdone = @()
        $jwait = $jobs.ChildJobs
        $jtimeout = $false
        do {
            $jdone_c = $jwait | wait-job -any -timeout $tout_c
            $td = (get-date) - $ttick
            if ($jdone_c) {
                Show-JobRuntime $jdone_c $jhash
                $tout_c = [int] ($tick - $td.TotalSeconds)
                if ($tout_c -lt 1) { $tout_c = 1 }
                $jdone += $jdone_c
                $jwait = $jwait |? { $_ -notin $jdone_c }
            } else {
                $jtimeout = $true
                write-host ("-"*20)
                $ttick = get-date
                $tout_c = $tick
                Show-JobRuntime $jwait $jhash -IncludeDone:$false
            }
        } while ($jwait)
        $null = Wait-Job $jobs
        if ($jtimeout) {
            write-host "Job Summary" -ForegroundColor Green
            Show-JobRuntime $jobs
        }
    }

    function Get-AdminSharePathFromLocal(
        [string] $node,
        [string] $local
    ) {
        "\\"+$node+"\"+$local[0]+"$\"+$local.Substring(3,$local.Length-3)
    }

    function Get-NodePath(
        [string] $Path,
        [string] $node
    ) {
        Join-Path $Path "Node_$node"
    }

    function NCount {
        Param ([object] $Item)
        if ($null -eq $Item) {
            $Result = 0
        } else {
            if ($Item.GetType().BaseType.Name -eq "Array") {
                $Result = ($Item).Count
            } else {
                $Result = 1
            }
        }
        return $Result
    }

    function Get-SddcCapturedEvents (
        [string] $Path,
        [int] $Hours
    ) {
        $QTime = $null
        if ($Hours -ne -1) {
            $MSecs = $Hours * 60 * 60 * 1000
            $QTime = "*[System[TimeCreated[timediff(@SystemTime) <= "+$MSecs+"]]]"
        }
        $LogToExclude = 'Microsoft-Windows-FailoverClustering/Diagnostic',
            'Microsoft-Windows-FailoverClustering/DiagnosticVerbose',
            'Microsoft-Windows-FailoverClustering-Client/Diagnostic',
            'Microsoft-Windows-Health/Diagnostic',
            'Microsoft-Windows-Health/DiagnosticVerbose',
            'Microsoft-Windows-PowerShell/Operational',
            'Microsoft-Windows-StorageReplica/Performance',
            'Microsoft-Windows-StorageSpaces-Driver/Performance',
            'Microsoft-Windows-SystemDataArchiver/Diagnostic',
            'Security'
        $providers = Get-WinEvent -ListLog * -ErrorAction Ignore -WarningAction Ignore
        $TxtPath = Join-Path $Path "GetWinEvent.txt"
        $XmlPath = Join-Path $Path "GetWinEvent.xml"
        $providers > $TxtPath
        $providers | Export-Clixml $XmlPath
        Write-Output $TxtPath
        Write-Output $XmlPath
        $cs = Get-CimInstance win32_computersystem
        $jobsMax = 0
        if ($null -ne $cs) {
            $jobsMax = [int] ($cs.NumberOfLogicalProcessors / 2)
        }
        if ($jobsMax -lt 10) { $jobsMax = 10 }
        $jobs = @{}
        $completions = @()
        function ConsumeJobs {
            param( [switch] $Any )
            if ($Any) {
                $jobsComplete = $jobs.Values | Wait-Job -Any
            } else {
                $jobsComplete = $jobs.Values | Wait-Job
            }
            $newCompletions = $jobsComplete | Receive-Job
            $jobsComplete | Remove-Job
            $jobsComplete |% { $jobs.Remove($_.Id) }
            return $newCompletions
        }
        foreach ($p in $providers) {
            if ($p.LogType -in @('Analytical','Debug') -and $p.IsEnabled) {
                $directChannel = $true
            } else {
                $directChannel = $false
            }
            if ($LogToExclude -contains $p.LogName -or (($p.RecordCount -eq 0 -or $null -eq $p.RecordCount) -and $directChannel -eq $false)) {
                continue
            }
            $EventFile = Join-Path $Path ($p.LogName.Replace("/","-")+".EVTX")
            if ($jobs.Count -ge $jobsMax) {
                $completions += ConsumeJobs -Any
            }
            $j = Start-Job -ArgumentList ($p, $EventFile, $QTime, $directChannel) {
                param( $p, $EventFile, $QTime, $directChannel )
                if ($directChannel) { wevtutil sl /e:false $p.LogName }
                $tepl = (Get-Date)
                if ($QTime) {
                    wevtutil epl $p.LogName $EventFile /q:$QTime /ow:true
                } else {
                    wevtutil epl $p.LogName $EventFile /ow:true
                }
                $tepl = (Get-Date) - $tepl
                if ($directChannel -eq $true) {
                    echo y | wevtutil sl /e:true $p.LogName | out-null
                }
                $tal = (Get-Date)
                wevtutil al $EventFile /l:$PSCulture
                $tal = (Get-Date) - $tal
                [pscustomobject] @{
                    EventFile   = $EventFile
                    LogName     = $p.LogName
                    RecordCount = $p.RecordCount
                    Direct      = $directChannel
                    Time        = $tepl + $tal
                    TimeExport  = $tepl
                    TimeArchive = $tal
                }
            }
            $jobs[$j.Id] = $j
        }
        $completions += ConsumeJobs
        if ($null -ne $completions) { $completions.EventFile }
        $XmlPath = Join-Path $Path "GetWinEvent-Timing.xml"
        $completions | Export-Clixml -Path $XmlPath
        Write-Output $XmlPath
        $t = Get-Date
        sleep 5
        dir $env:WINDIR\ServiceProfiles\LocalService\AppData\Local\Temp |? {
            ($_.Name -like 'MSG*.tmp' -or $_.Name -like 'EVT*.tmp' -or $_.Name -like 'PUB*.tmp') -and $_.CreateTime -lt $t
        } | del -Force -ErrorAction SilentlyContinue
    }

    function Format-SddcDateTime( [datetime] $d ) {
        $d.ToString('yyyyMMdd-HHmm')
    }

    function Test-SddcModulePresence {
        $Module = 'GetDellSDDC'
        $m = Get-Module $Module
        if (-not $m) {
            Write-Warning "Node $($env:COMPUTERNAME) does not have the $Module module installed for Sddc Diagnostic Archive. Please 'Install-SddcDiagnosticModule -Node $($env:COMPUTERNAME)' to address."
            $false
        } else {
            $true
        }
    }

    function Get-FilterXpath {
        [CmdletBinding(PositionalBinding=$false)]
        param (
            [int[]] $Event = @(),
            [datetime] $TimeBase,
            [ValidateScript({$_ -gt 0})]
            [int] $TimeDeltaMs = -1,
            [hashtable] $DataAnd = @{},
            [hashtable] $DataOr = @{}
        )
        $systemclauses = @()
        if ($Event.Count) {
            $c = $Event |% { "EventID = $_" }
            if ($Event.Count -gt 1) {
                $systemclauses += "(" + ($c -join " or ") + ")"
            } else {
                $systemclauses += $c
            }
        }
        if ($TimeDeltaMs -gt 0) {
            if ($TimeBase) {
                $t = $TimeBase.ToUniversalTime().ToString('s')
                $systemclauses += "(TimeCreated[timediff(@SystemTime,'$($t)') <= $($TimeDeltaMs)])"
            } else {
                $systemclauses += "(TimeCreated[timediff(@SystemTime) <= $($TimeDeltaMs)])"
            }
        }
        $systemclause = "System[" + ($systemclauses -join " and ") + "]"
        $cAnd = @($DataAnd.Keys | sort |% {
            "Data[@Name = '" + $_ + "'] " + $DataAnd[$_]
        }) -join " and "
        $cOr = @($DataOr.Keys | sort |% {
            "Data[@Name = '" + $_ + "'] " + $DataOr[$_]
        }) -join " or "
        if ($cAnd.Length -and $cOr.Length) {
            $dataclause = "EventData[$cAnd and ($cOr)]"
        } elseif ($cAnd.Length -or $cOr.Length) {
            $dataclause = "EventData[$($cAnd + $cOr)]"
        } else {
            $dataclause = $null
        }
        if ($dataclause) {
            $xpath = "*[$systemclause and $dataclause]"
        } else {
            $xpath = "*[$systemclause]"
        }
        $xpath
    }

    function Parse-SemicolonKVData {
        BEGIN { $o = new-object psobject }
        PROCESS {
            $null = $_ -match '^([^:]+)\s*:\s*(.*)$'
            $k = $matches[1]
            $v = $matches[2]
            if ($k -like '*time') {
                $o | Add-Member -NotePropertyName $matches[1] -NotePropertyValue ([datetime]$matches[2])
            } else {
                $o | Add-Member -NotePropertyName $matches[1] -NotePropertyValue $matches[2]
            }
        }
        END { $o }
    }

    function Count-EventLog {
        param(
            [string] $path,
            [string] $xpath
        )
        $f = New-TemporaryFile
        try {
            wevtutil epl /lf:true $path $f /q:$xpath /ow:true
            $gli = wevtutil gli /lf:true $f | Parse-SemicolonKVData
            $gli.numberOfLogRecords
        } finally {
            del -Force $f
        }
    }

    function Get-EventDataHash {
        param(
            [System.Diagnostics.Eventing.Reader.EventLogRecord] $event
        )
        $xh = @{}
        $x = ([xml]$event.ToXml()).Event.EventData.Data
        $x |% { $xh[$_.Name] = $_.'#text' }
        $xh
    }
}

# Compress the common function block for session passing
$CommonFunc = [scriptblock]::Create($(
    ((([string]$CommonFuncBlock) -split "`n") |? {
        $_ -notmatch '^\s*#'
    } |? {
        $_ -notmatch '^\s*$'
    }) -replace '^\s+','' -join "`n"
))

# Evaluate into the main session
. $CommonFunc

# endregion CommonFuncBlock

###############################################################################
# region Main-session-only helper functions
###############################################################################

function Test-PrefixFilePath( [ref] $path ) {
    $p = $path.Value
    $elements = @($p -split '\\')
    $lastempty = $elements[-1] -notmatch '\S'
    $islocabs = $elements[0].Length -and $elements[0][1] -eq ':'
    $isunc = $p -like '\\*'
    if ($lastempty -or ($islocabs -and $elements.Count -eq 1) -or ($isunc -and $elements.Count -lt 5)) {
        return $false
    }
    if (-not ($islocabs -or $isunc)) {
        if ($p[0] -eq '\') {
            $p = Join-Path ((Get-Location).Path.SubString(0,2)) $p
            if ($elements.Count -eq 2) {
                $path.Value = $p
                return $true
            }
            $elements = @($p -split '\\')
        } else {
            $p = Join-Path (Get-Location).Path $p
            if ($elements.Count -eq 1) {
                $path.Value = $p
                return $true
            }
            $elements = @($path.Value -split '\\')
        }
    }
    $tp = $elements[0..($elements.Count-2)] -join '\'
    if (Test-Path $tp) {
        $path.Value = $p
        $true
    } else {
        $false
    }
}

function Check-ExtractZip( [string] $Path ) {
    if (-not $Path.ToUpper().EndsWith(".ZIP")) { return $Path }
    $ExtractToPath = $Path.Substring(0, $Path.Length - 4)
    $f = gi $ExtractToPath -ErrorAction SilentlyContinue
    if ($f) { return $f.FullName }
    Show-Update "Extracting $Path -> $ExtractToPath"
    if (-not (New-Item -ItemType Directory -ErrorAction SilentlyContinue $ExtractToPath)) {
        Show-Error("Can't create directory for extraction")
    }
    compact /c $ExtractToPath | Out-Null
    try {
        Add-Type -Assembly System.IO.Compression.FileSystem
        [System.IO.Compression.ZipFile]::ExtractToDirectory($Path, $ExtractToPath)
    } catch {
        Show-Error("Can't extract results as Zip file from '$Path' to '$ExtractToPath'")
    }
    return $ExtractToPath
}

function Start-CopyJob(
    [string] $Path,
    [switch] $Delete,
    [object[]] $j
) {
    $j |% {
        $parent = $_
        $parent.ChildJobs |% {
            $logs = Receive-Job $_
            if (@($logs).Count) {
                $Destination = (Get-NodePath $Path $_.Location)
                if (Get-Member -InputObject $_ -Name Destination) {
                    $Destination = Join-Path $Destination $_.Destination
                    if (-not (Test-Path $Destination)) {
                        $null = md $Destination -Force -ErrorAction Continue
                    }
                }
                start-job -Name "Copy $($parent.Name) $($_.Location)" -ArgumentList $logs,$Destination,$Delete {
                    param($logs,$Destination,$Delete)
                    $logs |% {
                        Copy-Item -Recurse $_ $Destination -Force -ErrorAction Continue
                        if ($Delete) {
                            Remove-Item -Recurse $_ -Force -ErrorAction Continue
                        }
                    }
                }
            }
        }
    }
}

function Invoke-SddcCommonCommand (
    [string[]] $ClusterNodes = @(),
    [string] $JobName,
    [scriptblock] $InitBlock,
    [scriptblock] $ScriptBlock,
    [string] $SessionConfigurationName,
    [Object[]] $ArgumentList
) {
    $Job = @()
    $Sessions = @()
    $SessionIds = @()
    if ($ClusterNodes.Count -eq 0) {
        $Sessions = New-PSSession -Cn localhost -EnableNetworkAccess -ConfigurationName $SessionConfigurationName
    } else {
        $Sessions = New-PSSession -ComputerName $ClusterNodes -ConfigurationName $SessionConfigurationName
    }
    Invoke-Command -Session $Sessions $InitBlock
    $Job = Invoke-Command -Session $Sessions -AsJob -JobName $JobName -ScriptBlock $ScriptBlock -ArgumentList $ArgumentList
    $SessionIds = $Sessions.Id
    $Job | Add-Member -NotePropertyName ActiveSessions -NotePropertyValue $SessionIds
    return $Job
}

function Get-ClusterAccessNode( $Nodes ) {
    for ($i = 0; $i -lt $Nodes.count; $i++) {
        $Cluster = Get-Cluster $Nodes[$i].Name -ErrorAction SilentlyContinue
        if ($Cluster -ne $null) {
            return $Nodes[$i].Name
        }
    }
}

function Get-NodeList(
    [string] $Cluster,
    [string[]] $Nodes = @(),
    [switch] $Filter
) {
    $NodesToPing = @()
    $SuccesfullyPingedNodes = @()
    $NodesToReturn = @()
    if ($Nodes.Count) {
        $NodesToPing += $Nodes |% {
            New-Object -TypeName PSObject -Property @{
                "Name" = $_; "State" = "Down"; "Type" = "ManuallySpecifiedMachine"
            }
        }
    }
    $ClusterNodes = $null
    if ($Cluster -ne "" -and $Cluster -ne $null) {
        $ClusterNodes = Get-ClusterNode -Cluster $Cluster -ErrorAction SilentlyContinue
    }
    $NodeIdx = 0
    while ($ClusterNodes -eq $null -and $NodeIdx -lt $NodesToPing.Count) {
        $ClusterNodes = Get-ClusterNode -Cluster $NodesToPing[$NodeIdx].Name -ErrorAction SilentlyContinue
        $NodeIdx++
    }
    if ($ClusterNodes -ne $null) {
        if ($Nodes.Count) {
            for ($i = 0; $i -lt $ClusterNodes.Count; $i++) {
                $found = $false
                for ($j = 0; $j -lt $NodesToPing.Count; $j++) {
                    if ($NodesToPing[$j].Name -eq $ClusterNodes[$i].Name) {
                        $NodesToPing[$j] = $ClusterNodes[$i]
                        $found = $true
                        break
                    }
                }
                if ($found -ne $true) {
                    $NodesToPing += @($ClusterNodes[$i])
                }
            }
        } else {
            $NodesToPing = @($ClusterNodes)
        }
    }
    if ($NodesToPing.Count) {
        $PingResults = @()
        $j = $NodesToPing |% {
            Start-Job -ArgumentList $_ {
                param( $Node )
                if (Test-Connection -ComputerName $Node.Name -Quiet) { $Node }
            }
        }
        $null = Wait-Job $j
        $PingResults += $j | Receive-Job
        $j | Remove-Job
        for ($i = 0; $i -lt $PingResults.Count; $i++) {
            for ($j = 0; $j -lt $NodesToPing.Count; $j++) {
                if ($NodesToPing[$j].Name -eq $PingResults[$i].Name) {
                    $SuccesfullyPingedNodes += @($NodesToPing[$j])
                }
            }
        }
    }
    if ($Filter) {
        $NodesToReturn = $SuccesfullyPingedNodes
    } else {
        $NodesToReturn = $NodesToPing
    }
    return $NodesToReturn
}

# CONVERSION: Stub for Show-SddcDiagnosticReport — called by Invoke-GetDellSDDC
#             for summary generation. The full reporting functions are below.
#             This dispatcher calls the individual Get-*Report functions.
function Show-SddcDiagnosticReport {
    [CmdletBinding()]
    param(
        [parameter(Mandatory=$false)]
        [ReportType] $Report = [ReportType]::Summary,

        [parameter(Mandatory=$false)]
        [ReportLevelType] $ReportLevel = [ReportLevelType]::Standard,

        [parameter(Position=0, Mandatory=$true)]
        [string] $Path
    )

    $Path = Check-ExtractZip $Path
    if (-not (Test-Path $Path)) {
        Show-Error "Report path not found: $Path"
        return
    }

    # Ensure trailing slash
    if (-not $Path.EndsWith("\")) { $Path = $Path + "\" }

    $reports = @()
    if ($Report -eq [ReportType]::All -or $Report -eq [ReportType]::Summary) {
        $reports += 'Summary'
    }
    # Add other report types as needed when -Report All is specified
    if ($Report -eq [ReportType]::All) {
        $reports += 'StorageBusCache','StorageBusConnectivity','StorageLatency',
                    'StorageFirmware','SmbConnectivity','LSIEvent'
    } elseif ($Report -ne [ReportType]::Summary) {
        $reports += $Report.ToString()
    }

    foreach ($r in $reports) {
        switch ($r) {
            'Summary'                { try { Get-SummaryReport $Path -ReportLevel $ReportLevel } catch { Show-Warning "Summary report error: $_" } }
            'StorageBusCache'        { try { Get-StorageBusCacheReport $Path -ReportLevel $ReportLevel } catch { Show-Warning "SBC report error: $_" } }
            'StorageBusConnectivity' { try { Get-StorageBusConnectivityReport $Path -ReportLevel $ReportLevel } catch { Show-Warning "SBConn report error: $_" } }
            'StorageLatency'         { try { Get-StorageLatencyReport $Path -ReportLevel $ReportLevel } catch { Show-Warning "Latency report error: $_" } }
            'StorageFirmware'        { try { Get-StorageFirmwareReport $Path -ReportLevel $ReportLevel } catch { Show-Warning "Firmware report error: $_" } }
            'SmbConnectivity'        { try { Get-SmbConnectivityReport $Path -ReportLevel $ReportLevel } catch { Show-Warning "SMB report error: $_" } }
            'LSIEvent'               { try { Get-LsiEventReport $Path -ReportLevel $ReportLevel } catch { Show-Warning "LSI report error: $_" } }
        }
    }
}

# CONVERSION: Stub for archive parameter query — called in the gather path
function Get-SddcDiagnosticArchiveJobParameters {
    param(
        [parameter(Mandatory=$false)]
        [string] $Cluster = '.',
        [parameter(Mandatory=$false)]
        [ref] $Days,
        [parameter(Mandatory=$false)]
        [ref] $Path,
        [parameter(Mandatory=$false)]
        [ref] $Size,
        [parameter(Mandatory=$false)]
        [ref] $At
    )
    $c = Get-Cluster -Name $Cluster -ErrorAction Stop
    if ($PSBoundParameters.ContainsKey('Days')) {
        try { $Days.Value = ($c | Get-ClusterParameter -Name SddcDiagnosticArchiveDays -ErrorAction Stop).Value }
        catch { $Days.Value = 60 }
    }
    if ($PSBoundParameters.ContainsKey('Path')) {
        try { $Path.Value = ($c | Get-ClusterParameter -Name SddcDiagnosticArchivePath -ErrorAction Stop).Value }
        catch { $Path.Value = Join-Path $env:SystemRoot "SddcDiagnosticArchive" }
    }
    if ($PSBoundParameters.ContainsKey('Size')) {
        try { $Size.Value = ($c | Get-ClusterParameter -Name SddcDiagnosticArchiveSize -ErrorAction Stop).Value }
        catch { $Size.Value = 500MB }
    }
    if ($PSBoundParameters.ContainsKey('At')) {
        try {
            $Task = Get-ClusteredScheduledTask -Cluster $c.Name -TaskName SddcDiagnosticArchive -ErrorAction Stop
            $At.Value = [datetime] ($Task.TaskDefinition.Triggers[0].StartBoundary)
        } catch { $At.Value = [datetime] '3AM' }
    }
}

# CONVERSION: Stub for Show-SddcDiagnosticArchiveJob — referenced in gather
function Show-SddcDiagnosticArchiveJob {
    param(
        [parameter(Mandatory=$false)]
        [string] $Cluster = '.'
    )
    $c = Get-Cluster -Name $Cluster -ErrorAction Stop
    if (-not (Get-ClusteredScheduledTask -Cluster $c.Name |? TaskName -eq SddcDiagnosticArchive)) {
        Show-Error "SddcDiagnosticArchive job not currently registered"
    }
    $Days = $null; $Path = $null; $Size = $null; $At = $null
    Get-SddcDiagnosticArchiveJobParameters -Cluster $c.Name -Days ([ref] $Days) -Path ([ref] $Path) -Size ([ref] $Size) -At ([ref] $At)
    Write-Output "Target archive size per node : $('{0:0.00} MiB' -f ($Size/1MB))"
    Write-Output "Target days of archive       : $Days"
    Write-Output "Capture to path              : $Path"
    Write-Output "Capture at                   : $($At.ToString("h:mm tt"))"
    $Nodes = Get-NodeList -Cluster $Cluster -Filter
    Write-Output "$('-'*20)`nPer Node Report"
    $j = $Nodes | sort Name |% {
        icm $_.Name -AsJob {
            Import-Module $using:ModuleName -ErrorAction SilentlyContinue
            . ([scriptblock]::Create($using:CommonFunc))
            if (Test-SddcModulePresence) {
                $Path = $null
                Get-SddcDiagnosticArchiveJobParameters -Path ([ref] $Path)
                dir $Path\*.ZIP -ErrorAction SilentlyContinue | measure -Sum Length
            }
        }
    }
    $null = $j | Wait-Job
    $j | sort Location |% {
        $m = Receive-Job $_
        Remove-Job $_
        if ($m) { Write-Output "Node $($_.Location): $($m.Count) ZIPs which are $('{0:0.00} MiB' -f ($m.Sum/1MB))" }
    }
}

# CONVERSION: Stub for Confirm-SddcDiagnosticModule — referenced in gather
function Confirm-SddcDiagnosticModule {
    [CmdletBinding()]
    param(
        [parameter(ParameterSetName="Cluster", Mandatory=$false)]
        [string] $Cluster = '.',
        [parameter(ParameterSetName="Node", Mandatory=$true)]
        [string[]] $Node
    )
    switch ($psCmdlet.ParameterSetName) {
        "Cluster" { $Nodes = Get-NodeList -Cluster $Cluster -Filter }
        "Node"    { $Nodes = Get-NodeList -Nodes $Node -Filter }
    }
    $thisVersion = $script:ScriptVersion
    $clusterModules = icm $Nodes.Name {
        $null = Import-Module -Force $using:ModuleName -ErrorAction SilentlyContinue
        Get-Module $using:ModuleName
    }
    $Nodes.Name |? { $_ -notin $clusterModules.PsComputerName } |% {
        Write-Warning "Node $_ does not have the $($script:ModuleName) module."
    }
    $clusterModules
}

# endregion Main-session helpers

###############################################################################
# region Report enums and functions
###############################################################################

enum ReportLevelType {
    Summary  = 0
    Standard
    Full
}

enum ReportType {
    All = 0
    Summary
    SmbConnectivity
    StorageBusCache
    StorageBusConnectivity
    StorageLatency
    StorageFirmware
    LSIEvent
}

function Get-ClusterLogDataSource {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidUsingCmdletAliases", "")]
    param( [string] $logname )
    BEGIN {
        $csvf = New-TemporaryFile
        $sr = [System.IO.StreamReader](gi $logname).FullName
        $datasource = @{}
    }
    PROCESS {
        $firstline = $false
        $in = $false
        $section = $null
        do {
            $l = $sr.ReadLine()
            if ($in) {
                if ($firstline) {
                    $firstline = $false
                    if (($l -split ',').count -lt 4) {
                        $in = $false
                    } else {
                        if ($section -eq 'Resources' -and $l -match '^(.*?)(_embeddedFailureAction)(.*)$') {
                            $l = $matches[1]+"ignore"+$matches[3]
                        }
                        $n = 0
                        while ($l -match '^(.*?)(,ignore,)(.*)$') {
                            $l = $matches[1]+",ignore$n,"+$matches[3]
                            $n += 1
                        }
                        $l | out-file -Encoding ascii -Width 9999 $csvf
                    }
                } else {
                    if ($l -notmatch '^\s*$') {
                        $l | out-file -Append -Encoding ascii -Width 9999 $csvf
                    } else {
                        $datasource[$section] = import-csv $csvf
                        $in = $false
                        $section = $null
                    }
                }
            } elseif ($l -match '^\[===\s(.*)\s===\]') {
                if ($matches[1] -eq 'System') { break }
                $section = $matches[1]
                $in = $true
                $firstline = $true
            }
        } while (-not $sr.EndOfStream)
    }
    END {
        $datasource
        $sr.Close()
        del $csvf
    }
}

function Format-StorageBusCacheDiskState( [string] $DiskState ) {
    $DiskState -replace 'CacheDiskState',''
}

function Get-StorageBusCacheReport {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidUsingCmdletAliases", "")]
    param(
        [parameter(Position=0, Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string] $Path,
        [parameter(Mandatory=$true)]
        [ReportLevelType] $ReportLevel
    )
    dir $Path\*cluster.log | sort -Property BaseName |% {
        $node = ""
        if ($_.BaseName -match "^(.*)_cluster$") { $node = $matches[1] }
        Write-Output ("-"*40)
        "Node: $node"
        $data = Get-ClusterLogDataSource $_.FullName
        $d = $data['SBL Disks']
        if ($d) {
            $idmap = @{}
            $d |% { $idmap[$_.DiskId] = $_.DeviceNumber }
            if ($ReportLevel -eq [ReportLevelType]::Full) {
                $d | sort IsSblCacheDevice,CacheDeviceId,DiskState | ft -AutoSize @{
                    Label = 'DiskState'; Expression = { Format-StorageBusCacheDiskState $_.DiskState }},
                    DiskId,ProductId,Serial,
                    @{ Label = 'Device#'; Expression = {$_.DeviceNumber} },
                    @{ Label = 'CacheDevice#'; Expression = {
                        if ($_.IsSblCacheDevice -eq 'true') { '= cache' }
                        elseif ($idmap.ContainsKey($_.CacheDeviceId)) { $idmap[$_.CacheDeviceId] }
                        elseif ($_.CacheDeviceId -eq '{00000000-0000-0000-0000-000000000000}') { "= unbound" }
                        else { "= not present $($_.CacheDeviceId)" }
                    }},
                    @{ Label = 'SeekPenalty'; Expression = {$_.HasSeekPenalty} },
                    PathId,BindingAttributes,DirtyPages
            }
            $dcache = $d |? IsSblCacheDevice -eq 'true'
            $dcap = $d |? IsSblCacheDevice -ne 'true'
            Write-Output "Device counts: cache $($dcache.count) capacity $($dcap.count)"
            if ($dcache) {
                $uneven = $false
                if ($dcap.count % $dcache.count) {
                    $uneven = $true
                    Write-Warning "Capacity device count does not evenly distribute to cache devices"
                }
                $unbound = $dcap |? CacheDeviceId -eq '{00000000-0000-0000-0000-000000000000}'
                if ($unbound) { Write-Warning "There are $(@($unbound).count) unbound capacity device(s)" }
                if (-not $uneven -and ($dcap.count - @($unbound).count) % $dcache.count) { $uneven = $true }
                $gdev = $dcap |? DiskState -eq 'CacheDiskStateInitializedAndBound' | group -property CacheDeviceId
                if (@($gdev).count -ne $dcache.count) { Write-Warning "Not all cache devices in use" }
                $gdist = $gdev |% { $_.count } | group
                if (@($gdist).count -eq 1) {
                    Write-Output "Binding ratio is even: 1:$($gdist.name)"
                } else {
                    $delta = [math]::Abs([int]$gdist[0].name - [int]$gdist[1].name)
                    if ($delta -eq 1 -and $uneven) {
                        Write-Output "Binding ratios are as expected for uneven device ratios"
                    } else {
                        Write-Warning "Binding ratios are uneven"
                    }
                    $s = $($gdist |% { "1:$($_.name) ($($_.count) total)" }) -join ", "
                    Write-Output "Groups: $s"
                }
            }
            $g = $d | group -property DiskState
            if (@($g).count -ne 1) {
                write-output "Disk State Summary:"
                $g | sort -property Name | ft @{ Label = 'DiskState'; Expression = { Format-StorageBusCacheDiskState $_.Name}},
                    @{ Label = "Number of Disks"; Expression = { $_.Count }}
            } else {
                write-output "All disks are in $(Format-StorageBusCacheDiskState $g.name)"
            }
        }
    }
}

function Get-StorageBusConnectivityReport {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidUsingCmdletAliases", "")]
    param(
        [parameter(Position=0, Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string] $Path,
        [parameter(Mandatory=$true)]
        [ReportLevelType] $ReportLevel
    )
    function Show-SSBConnectivity($node) {
        BEGIN { $disks = 0; $enc = 0; $ssu = 0 }
        PROCESS {
            switch ($_.DeviceType) {
                0 { $disks += 1 }
                1 { $enc += 1 }
                2 { $ssu += 1 }
            }
        }
        END { "$node has $disks disks, $enc enclosures, and $ssu scaleunit" }
    }
    dir $path\Node_*\ClusPort.xml | sort -Property FullName |% {
        $file = $_.FullName
        $node = ""
        if ($file -match "Node_([^\\]+)\\") { $node = $matches[1] }
        Import-ClixmlIf $_ | Show-SSBConnectivity $node
    }
}

function Get-StorageLatencyReport {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidUsingCmdletAliases", "")]
    param(
        [parameter(Position=0, Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string] $Path,
        [parameter(Mandatory=$true)]
        [ReportLevelType] $ReportLevel,
        [int] $CutoffMs = 0,
        [datetime] $TimeBase,
        [int] $HoursOfEvents = -1
    )
    if ($CutoffMs) {
        Write-Output "Latency Cutoff: report limited to IO of $($CutoffMs)ms and higher"
    } else {
        Write-Output "Latency Cutoff: none"
    }
    if ($HoursOfEvents -eq -1) {
        Write-Output "Time Cutoff   : none"
    } else {
        Write-Output "Time Cutoff   : from $($TimeBase.ToString()) for the prior $HoursOfEvents hours"
    }
    if (-not $CutoffMs -and $HoursOfEvents -eq -1) {
        write-output "NOTE: time/latency cutoff limits may significantly speed up reporting"
    }
    $j = @()
    dir $Path\Node_*\Microsoft-Windows-Storage-Storport-Operational.EVTX | sort -Property FullName |% {
        $file = $_.FullName
        $node = ""
        if ($file -match "Node_([^\\]+)\\") { $node = $matches[1] }
        $j += Invoke-SddcCommonCommand -InitBlock $CommonFunc -JobName $node -SessionConfigurationName $null -ScriptBlock {
            $dofull = $false
            if ($using:ReportLevel -eq "Full") { $dofull = $true }
            function Get-Bucket {
                param( [int] $i, [int] $max, [string[]] $s )
                $i .. $max |% { $l = $_; $s |% { "BucketIo$_$l" } }
            }
            $buckhash = @{}
            $bucklabels = $null
            $buckvalueschema = $null
            $cutoffbuck = 1
            $evs = @()
            $e = Get-WinEvent -Path $using:file -FilterXPath (Get-FilterXpath -Event 505) -ErrorAction SilentlyContinue -MaxEvents 1
            if ($e) {
                $xh = Get-EventDataHash $e
                $bucklabels = $xh['IoLatencyBuckets'] -split ',\s+'
                if ($xh.ContainsKey("BucketIoSuccess1")) {
                    $schemasplit = $true
                    $buckvalueschema = "^BucketIo(Success|Failed)(\d+)$"
                } else {
                    $schemasplit = $false
                    $buckvalueschema = "^BucketIo(Count)(\d+)$"
                }
                $DataOr = @{}
                if ($using:CutoffMs) {
                    $CutoffUs = $using:CutoffMs * 1000
                    $a = $xh['IoLatencyBuckets'] -split ',\s+' |% {
                        switch -Regex ($_) {
                            "^(\d+)us$"   { [int] $matches[1] }
                            "^(\d+)ms$"   { ([int] $matches[1]) * 1000 }
                            "^(\d+)\+ms$" { [int]::MaxValue }
                            default { throw "misparsed bucket label $_" }
                        }
                    }
                    foreach ($i in 0..($a.Count - 1)) {
                        if ($CutoffUs -lt $a[$i]) { $cutoffbuck = $i+1; break }
                    }
                    if ($schemasplit) {
                        $buck = Get-Bucket $cutoffbuck $a.Count 'Success','Failed'
                    } else {
                        $buck = Get-Bucket $cutoffbuck $a.Count 'Count'
                    }
                    $DataOr = @{}
                    $buck |% { $DataOr[$_] = "> 0" }
                }
                if ($cutoffbuck -eq 1) {
                    $bucklabels[0] = "0-" + $bucklabels[0]
                } else {
                    $bucklabels[$cutoffbuck - 1] = $bucklabels[$cutoffbuck - 2] + "-" + $bucklabels[$cutoffbuck - 1]
                    $bucklabels = $bucklabels[($cutoffbuck - 1) .. ($bucklabels.Count - 1)]
                }
                if ($using:HoursOfEvents -ne -1) {
                    $xpath = Get-FilterXpath -Event 505 -TimeBase $using:TimeBase -TimeDeltaMs ($using:HoursOfEvents * 60 * 60 * 1000) -DataOr $DataOr
                } else {
                    $xpath = Get-FilterXpath -Event 505 -DataOr $DataOr
                }
                Get-WinEvent -Path $using:file -FilterXPath $xpath |% {
                    $xh = Get-EventDataHash $_
                    $dev = [string] $xh['ClassDeviceGuid']
                    if ($dev -match '{(.*)}') { $dev = $matches[1] }
                    $buckvalues = @($null) * $bucklabels.length
                    $xh.Keys |% {
                        if ($_ -match $buckvalueschema) {
                            $thisbuck = [int] $matches[2]
                            if ($thisbuck -ge $cutoffbuck) {
                                $buckvalues[$thisbuck - $cutoffbuck] += [int] $xh[$_]
                            }
                        }
                    }
                    if ($buckvalues -contains $null) { throw "misparsed 505 event latency buckets" }
                    if (-not $buckhash.ContainsKey($dev)) {
                        $buckhash[$dev] = $buckvalues |% { if ($_) { 1 } else { 0 }}
                    } else {
                        foreach ($i in 0..($buckvalues.count - 1)) {
                            if ($buckvalues[$i]) { $buckhash[$dev][$i] += 1 }
                        }
                    }
                    if ($dofull -and ($buckvalues[-1] -ne 0 -or $cutoffbuck -ne 1)) {
                        $evs += $(
                            $o = New-Object psobject -Property @{ 'Time' = $_.TimeCreated; 'Device' = [string] $_.Properties[4].Value }
                            foreach ($i in 0..($bucklabels.count -1)) {
                                $o | Add-Member -NotePropertyName $bucklabels[$i] -NotePropertyValue $buckvalues[$i]
                            }
                            $o
                        )
                    }
                }
                ,$bucklabels
                $buckhash
                $evs
            }
        }
    }
    $PhysicalDisks = Import-ClixmlIf (Join-Path $Path "GetPhysicalDisk.XML")
    $PhysicalDisksTable = @{}
    $PhysicalDisks |% {
        if ($_.ObjectId -match 'PD:{(.*)}') { $PhysicalDisksTable[$matches[1]] = $_ }
    }
    $pdattrs_tab = @{ Label = 'FriendlyName'; Expression = { $PhysicalDisksTable[$_.Device].FriendlyName }},
        @{ Label = 'SerialNumber'; Expression = { $PhysicalDisksTable[$_.Device].SerialNumber }},
        @{ Label = 'Firmware'; Expression = { $PhysicalDisksTable[$_.Device].FirmwareVersion }},
        @{ Label = 'Media'; Expression = { $PhysicalDisksTable[$_.Device].MediaType }},
        @{ Label = 'Usage'; Expression = { $PhysicalDisksTable[$_.Device].Usage }},
        @{ Label = 'OpStat'; Expression = { $PhysicalDisksTable[$_.Device].OperationalStatus }},
        @{ Label = 'HealthStat'; Expression = { $PhysicalDisksTable[$_.Device].HealthStatus }}
    $pdattrs_ev = @{ Label = 'FriendlyName'; Expression = { $PhysicalDisksTable[$_.Device].FriendlyName }},
        @{ Label = 'SerialNumber'; Expression = { $PhysicalDisksTable[$_.Device].SerialNumber }},
        @{ Label = 'Media'; Expression = { $PhysicalDisksTable[$_.Device].MediaType }},
        @{ Label = 'Usage'; Expression = { $PhysicalDisksTable[$_.Device].Usage }}
    $j | Wait-Job | sort name |% {
        ($bucklabels, $buckhash, $evs) = receive-job $_
        $node = $_.Name
        remove-job $_
        Write-Output ("-"*40),"Node: $node","`nSample Period Count Report"
        if ($buckhash.Count -eq 0) {
            Write-Warning "Node $node is not reporting latency information."
        } else {
            $buckhash.Keys |? { $PhysicalDisksTable.ContainsKey($_) } |% {
                $dev = $_
                $vprop = @{}
                $weight = 0
                foreach ($i in 0..($bucklabels.count - 1)) {
                    $v = $buckhash[$_][$i]
                    if ($v) { $weight = $i; $weightval = $v; $vprop[$bucklabels[$i]] = $v }
                }
                $vprop['Device'] = $dev
                $vprop['Weight'] = $weight
                $vprop['WeightVal'] = $weightval
                new-object psobject -Property $vprop
            } | sort Weight,@{ Expression = {$PhysicalDisksTable[$_.Device].Usage}},WeightVal |
                ft -AutoSize (,'Device' + $pdattrs_tab + $bucklabels)
            if ($ReportLevel -eq [ReportLevelType]::Full) {
                Write-Output "`nHigh Latency Events"
                $n = 0
                if ($null -ne $evs) {
                    $evs |? { $PhysicalDisksTable.ContainsKey($_.Device) } |% { $n += 1; $_ } |
                        sort Time -Descending | ft -AutoSize ('Time','Device' + $pdattrs_ev + $bucklabels)
                }
                if ($n -eq 0) { Write-Output "-> No Events" }
            }
        }
    }
}

function Get-StorageFirmwareReport {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidUsingCmdletAliases", "")]
    param(
        [parameter(Position=0, Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string] $Path,
        [parameter(Mandatory=$true)]
        [ReportLevelType] $ReportLevel
    )
    $PhysicalDisks = Import-ClixmlIf (Join-Path $Path "GetPhysicalDisk.XML") |? Usage -ne Retired
    Write-Output "Total Firmware Report"
    $PhysicalDisks | group -Property Manufacturer,Model,FirmwareVersion | sort Name | ft @{
        Label = 'Number'; Expression = { $_.Count }},
        @{ Label = 'Manufacturer'; Expression = { $_.Group[0].Manufacturer }},
        @{ Label = 'Model'; Expression = { $_.Group[0].Model }},
        @{ Label = 'Firmware'; Expression = { $_.Group[0].FirmwareVersion }},
        @{ Label = 'Media'; Expression = { $_.Group[0].MediaType }},
        @{ Label = 'Usage'; Expression = { $_.Group[0].Usage }}
    Write-Output "Per Unit Firmware Report`n"
    $good = @()
    $PhysicalDisks | group -Property Manufacturer,Model | sort Name |% {
        $fwg = $_.Group | group -Property FirmwareVersion | sort -Property Count
        if (($fwg | measure).Count -ne 1) {
            Write-Output "$($_.Group[0].Manufacturer) $($_.Group[0].Model): varying firmware found - $($fwg.Name -join ' ')"
            Write-Output "Majority Devices: $($fwg[-1].Count) are at firmware version $($fwg[-1].Group[0].FirmwareVersion)"
            Write-Output "Minority Devices:"
            $fwg | select -SkipLast 1 |% {
                Write-Output "Firmware Version $($_.Name) - Total $($_.Count)"
                $_.Group | ft @{ Label = 'SerialNumber'; Expression = {
                    if ($_.BusType -eq 'NVME') { $_.AdapterSerialNumber } else { $_.SerialNumber }
                }}, @{ Label = "Media"; Expression = { $_.MediaType }}, Usage
            }
        } else {
            $good += "$($_.Group[0].Manufacturer) $($_.Group[0].Model): all devices are on firmware version $($_.Group[0].FirmwareVersion)"
        }
    }
    Write-Output $good
}

function Get-LsiEventReport {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidUsingCmdletAliases", "")]
    param(
        [parameter(Position=0, Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string] $Path,
        [parameter(Mandatory=$true)]
        [ReportLevelType] $ReportLevel
    )
    dir $Path\Node_*\System.EVTX | sort -Property FullName |% {
        $node = ""
        if ($_.FullName -match "Node_([^\\]+)\\") { $node = $matches[1] }
        Write-Output ("-"*40)
        "Node: $node"
        $ev = Get-WinEvent -Path $_ -FilterXPath '*[System[(EventID=11)]]' -ErrorAction SilentlyContinue |? ProviderName -match "lsi" |% {
            new-object psobject -Property @{
                'Time' = $_.TimeCreated
                'Provider Name' = $_.ProviderName
                'LSI Error'= (($_.Properties[1].Value[19..16] |% { '{0:X2}' -f $_ }) -join '')
            }
        }
        if (-not $ev) {
            Write-Output "No LSI events present"
        } else {
            Write-Output "Summary of LSI Event 11 error codes"
            $ev | group -Property 'LSI Error' -NoElement | sort -Property Name | ft -AutoSize Count,@{
                Label = 'LSI Error'; Expression = { $_.Name }}
            if ($ReportLevel -eq [ReportLevelType]::Full) {
                Write-Output "LSI Event 11 errors by time"
                $ev | ft Time,'LSI Error'
            }
        }
    }
}

function Get-SmbConnectivityReport {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidUsingCmdletAliases", "")]
    param(
        [parameter(Position=0, Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string] $Path,
        [parameter(Mandatory=$true)]
        [ReportLevelType] $ReportLevel
    )
    $ReportTableBlock = {
        param( [string[]] $paths, [int] $ev, [datetime] $timebase, [System.ConsoleColor] $warncol, [string] $warn )
        $r = $paths |% {
            $node = ""
            if ($_ -match "Node_([^\\]+)\\") { $node = $matches[1] }
            $last5    = (1000*60*5)
            $lasthour = (1000*60*60)
            $lastday  = (1000*60*60*24)
            New-Object psobject -Property @{
                'ComputerName'  = $node
                'RDMA Last5Min' = Count-EventLog -path $_ -xpath $(Get-FilterXpath -Event $ev -TimeBase $timebase -TimeDeltaMs $last5    -DataAnd @{'ConnectionType'='=2'})
                'RDMA LastHour' = Count-EventLog -path $_ -xpath $(Get-FilterXpath -Event $ev -TimeBase $timebase -TimeDeltaMs $lasthour -DataAnd @{'ConnectionType'='=2'})
                'RDMA LastDay'  = Count-EventLog -path $_ -xpath $(Get-FilterXpath -Event $ev -TimeBase $timebase -TimeDeltaMs $lastday  -DataAnd @{'ConnectionType'='=2'})
                'TCP Last5Min'  = Count-EventLog -path $_ -xpath $(Get-FilterXpath -Event $ev -TimeBase $timebase -TimeDeltaMs $last5    -DataAnd @{'ConnectionType'='=1'})
                'TCP LastHour'  = Count-EventLog -path $_ -xpath $(Get-FilterXpath -Event $ev -TimeBase $timebase -TimeDeltaMs $lasthour -DataAnd @{'ConnectionType'='=1'})
                'TCP LastDay'   = Count-EventLog -path $_ -xpath $(Get-FilterXpath -Event $ev -TimeBase $timebase -TimeDeltaMs $lastday  -DataAnd @{'ConnectionType'='=1'})
            }
        }
        $hdr = 'ComputerName','RDMA Last5Min','RDMA LastHour','RDMA LastDay','TCP Last5Min','TCP LastHour','TCP LastDay'
        $rdmafail = ($r |% { $row = $_; $hdr |? {$_ -like 'RDMA*' } |% { $row.$_ }} | measure -sum).sum -ne 0
        if ($rdmafail) { Write-Host -ForegroundColor $warncol $warn }
        $r | sort -Property ComputerName | ft -Property $hdr
    }
    $Parameters = Import-ClixmlIf (Join-Path $Path "GetParameters.XML")
    $CaptureDate = $Parameters.TodayDate
    Write-Host "This report is relative to the time of data capture: $($CaptureDate)"
    $eventlogs = (dir $Path\Node_*\Microsoft-Windows-SmbClient-Connectivity.EVTX).FullName
    $j = @()
    $w = @"
WARNING: the SMB Client is receiving RDMA disconnects. This is an error whose root
`t cause may be PFC/CoS misconfiguration (if RoCE) on hosts or switches, physical
`t issues (ex: bad cable), switch or NIC firmware issues.
"@
    $j += Start-Job -name 'SMB Connectivity Error Check - Disconnect Failures (Event 30804)' -InitializationScript $CommonFunc -ScriptBlock $ReportTableBlock -ArgumentList $eventlogs,30804,$CaptureDate,([ConsoleColor]'Red'),$w
    $w = @"
WARNING: the SMB Client is receiving RDMA connect errors. This is an error whose root
`t cause may be actual lack of connectivity or fundamental problems with the RDMA
`t network fabric.
"@
    $j += Start-Job -name 'SMB Connectivity Error Check - Connect Failures (Event 30803)' -InitializationScript $CommonFunc -ScriptBlock $ReportTableBlock -ArgumentList $eventlogs,30803,$CaptureDate,([ConsoleColor]'Yellow'),$w
    $null = $j | Wait-Job
    $j | sort Name |% {
        Write-Host -ForegroundColor Cyan $_.Name
        Receive-Job $_
    }
    $j | Remove-Job
}

function Get-SummaryReport {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidUsingCmdletAliases", "")]
    param(
        [parameter(Position=0, Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string] $Path,
        [parameter(Mandatory=$true)]
        [ReportLevelType] $ReportLevel
    )
    $Parameters = Import-ClixmlIf (Join-Path $Path "GetParameters.XML")
    $TodayDate             = $Parameters.TodayDate
    $ExpectedNodes         = $Parameters.ExpectedNodes
    $ExpectedNetworks      = $Parameters.ExpectedNetworks
    $ExpectedVolumes       = $Parameters.ExpectedVolumes
    $ExpectedDedupVolumes  = $Parameters.ExpectedDedupVolumes
    $ExpectedPhysicalDisks = $Parameters.ExpectedPhysicalDisks
    $ExpectedPools         = $Parameters.ExpectedPools
    $ExpectedEnclosures    = $Parameters.ExpectedEnclosures
    $HoursOfEvents         = $Parameters.HoursOfEvents

    # CONVERSION: references to module version replaced
    Show-Update "Gathered with  : $($Parameters.Version)"
    Show-Update "Report created : $($script:ScriptVersion)"

    Show-Update "<<< Phase 1 - Health Overview >>>`n" -ForegroundColor Cyan
    Write-Host ("Date of capture : " + $TodayDate)
    $ClusterNodes = Import-ClixmlIf (Join-Path $Path "GetClusterNode.XML")
    $Cluster = Import-ClixmlIf (Join-Path $Path "GetCluster.XML")
    if ($Cluster) {
        $ClusterName = $Cluster.Name + "." + $Cluster.Domain
        $S2DEnabled = $Cluster.S2DEnabled
        $ClusterDomain = $Cluster.Domain
        Write-Host "Cluster Name    : $ClusterName"
        Write-Host "S2D Enabled     : $S2DEnabled"
    } else {
        Write-Host "Cluster Name    : Cluster was unavailable"
        Write-Host "S2D Enabled     : Cluster was unavailable"
    }
    $f = Join-Path $Path SddcDiagnosticArchiveJob.txt
    if (gi -ErrorAction SilentlyContinue $f) {
        Write-Host "$("-"*3)`nSddc Diagnostic Archive Status`n"
        gc $f
        $f = Join-Path $Path SddcDiagnosticArchiveJobWarn.txt
        if ((gi $f).Length) { gc $f |% { Show-Warning $_ } }
        Write-Host $("-"*3)
    }
    if ($Cluster) {
        $ClusterGroups = Import-ClixmlIf (Join-Path $Path "GetClusterGroup.XML")
        $ScaleOutServers = $ClusterGroups |? GroupType -like "ScaleOut*"
        if ($null -eq $ScaleOutServers) {
            if ($S2DEnabled -ne $true) { Show-Warning "No Scale-Out File Server cluster roles found" }
        } else {
            $ScaleOutName = $ScaleOutServers[0].Name + "." + $ClusterDomain
            Write-Host "Scale-Out File Server Name : $ScaleOutName"
        }
        $NodesTotal   = NCount($ClusterNodes)
        $NodesHealthy = NCount($ClusterNodes |? {$_.State -like "Paused" -or $_.State -like "Up"})
        Write-Host "Cluster Nodes up            : $NodesHealthy / $NodesTotal"
        if ($NodesTotal -lt $ExpectedNodes)   { Show-Warning "Fewer nodes than the $ExpectedNodes expected" }
        if ($NodesHealthy -lt $NodesTotal)    { Show-Warning "Unhealthy nodes detected" }
        $ClusterNetworks = Import-ClixmlIf (Join-Path $Path "GetClusterNetwork.XML")
        $NetsTotal   = NCount($ClusterNetworks)
        $NetsHealthy = NCount($ClusterNetworks |? {$_.State -like "Up"})
        Write-Host "Cluster Networks up         : $NetsHealthy / $NetsTotal"
        if ($NetsTotal -lt $ExpectedNetworks) { Show-Warning "Fewer cluster networks than the $ExpectedNetworks expected" }
        if ($NetsHealthy -lt $NetsTotal)      { Show-Warning "Unhealthy cluster networks detected" }
        $ClusterResources = Import-ClixmlIf (Join-Path $Path "GetClusterResource.XML")
        $ClusterResourceParameters = Import-ClixmlIf (Join-Path $Path "GetClusterResourceParameters.XML")
        $ResTotal   = NCount($ClusterResources)
        $ResHealthy = NCount($ClusterResources |? State -like "Online")
        Write-Host "Cluster Resources Online    : $ResHealthy / $ResTotal "
        if ($ResHealthy -lt $ResTotal) { Show-Warning "Unhealthy cluster resources detected" }
        if ($S2DEnabled) {
            $HealthProviders = $ClusterResourceParameters |? { $_.ClusterObject -like 'Health' -and $_.Name -eq 'Providers' }
            $HealthProviderCount = $HealthProviders.Value.Count
            if ($HealthProviderCount) {
                Write-Host "Health Resource             : $HealthProviderCount health providers registered"
            } else {
                Show-Warning "Health Resource providers not registered"
            }
        }
    } else {
        Show-Warning "Skipping Cluster status since it was unavailable"
    }
    $Subsystem = Import-ClixmlIf (Join-Path $Path "GetStorageSubsystem.XML")
    $SubsystemUnhealthy = $false
    if ($Subsystem -eq $null) {
        Show-Warning "No clustered storage subsystem present"
    } elseif ($Subsystem.HealthStatus -notlike "Healthy") {
        $SubsystemUnhealthy = $true
        Show-Warning "Clustered storage subsystem '$($Subsystem.FriendlyName)' is in health state $($Subsystem.HealthStatus)"
    } else {
        Write-Host "Clustered storage subsystem '$($Subsystem.FriendlyName)' is healthy"
    }
    $VerifiedNodes = @()
    foreach ($node in $ClusterNodes.Name) {
        $f = Join-Path (Get-NodePath $Path $node) "verifier-query.txt"
        $o = @(gc $f)
        if (-not ($o.Count -eq 1 -and $o[0] -eq 'No drivers are currently verified.')) { $VerifiedNodes += $node }
    }
    if ($VerifiedNodes.Count -ne 0) {
        Show-Warning "The following $($VerifiedNodes.Count) node(s) have system verification (verifier.exe) active."
        $VerifiedNodes |% { Write-Host "`t$_" }
    } else {
        Write-Host "No nodes currently under the system verifier."
    }
    $StorageJobs = Import-ClixmlIf (Join-Path $Path "GetStorageJob.XML")
    if ($StorageJobs -eq $null) {
        Write-Host "No storage jobs were present at the time of the gather"
    } else {
        Show-Warning "Storage jobs were present"
        $StorageJobs | ft -AutoSize
    }
    Write-Host "`nHealthy Components count: [SMBShare -> CSV -> VirtualDisk -> StoragePool -> PhysicalDisk -> StorageEnclosure]"
    $ShareStatus = Import-ClixmlIf (Join-Path $Path "ShareStatus.XML")
    $ShTotal   = NCount($ShareStatus)
    $ShHealthy = NCount($ShareStatus |? Health -like "Accessible")
    "SMB CA Shares Accessible        : $ShHealthy / $ShTotal"
    if ($ShHealthy -lt $ShTotal) { Show-Warning "Inaccessible CA shares detected" }
    $SmbOpenFiles = Import-ClixmlIf (Join-Path $Path "GetSmbOpenFile.XML")
    $FileTotal = NCount( $SmbOpenFiles | Group-Object ClientComputerName)
    Write-Host "Users with Open Files           : $FileTotal"
    if ($FileTotal -eq 0) { Show-Warning "No users with open files" }
    $SmbWitness = Import-ClixmlIf (Join-Path $Path "GetSmbWitness.XML")
    $WitTotal = NCount($SmbWitness |? State -eq RequestedNotifications | Group-Object ClientName)
    Write-Host "Users with a Witness            : $WitTotal"
    if ($FileTotal -ne 0 -and $WitTotal -eq 0) { Show-Warning "No users with a Witness" }
    if ($Cluster) {
        $CSV = Import-ClixmlIf (Join-Path $Path "GetClusterSharedVolume.XML")
        $CSVTotal   = NCount($CSV)
        $CSVHealthy = NCount($CSV |? State -like "Online")
        Write-Host "Cluster Shared Volumes Online   : $CSVHealthy / $CSVTotal"
        if ($CSVHealthy -lt $CSVTotal) { Show-Warning "Offline cluster shared volumes detected" }
    } else {
        Show-Warning "Skipping CSV status since cluster was unavailable"
    }
    $Volumes = Import-ClixmlIf (Join-Path $Path "GetVolume.XML")
    $VolsTotal   = NCount($Volumes |? FileSystem -eq CSVFS)
    $VolsHealthy = NCount($Volumes |? FileSystem -eq CSVFS |? { ($_.HealthStatus -like "Healthy") -or ($_.HealthStatus -eq 0) })
    Write-Host "Cluster Shared Volumes Healthy  : $VolsHealthy / $VolsTotal "
    $DedupEnabled = $false
    if (Test-Path (Join-Path $Path "GetDedupVolume.XML")) {
        $DedupEnabled = $true
        $DedupVolumes = Import-ClixmlIf (Join-Path $Path "GetDedupVolume.XML")
        $DedupTotal   = NCount($DedupVolumes)
        $DedupHealthy = NCount($DedupVolumes |? LastOptimizationResult -eq 0)
        if ($DedupTotal) {
            Write-Host "Dedup Volumes Healthy           : $DedupHealthy / $DedupTotal "
            if ($DedupHealthy -lt $DedupTotal) { Show-Warning "Unhealthy Dedup volumes detected" }
        } else { $DedupHealthy = 0 }
        if ($DedupTotal -lt $ExpectedDedupVolumes) { Show-Warning "Fewer Dedup volumes than the $ExpectedDedupVolumes expected" }
    }
    $VirtualDisks = Import-ClixmlIf (Join-Path $Path "GetVirtualDisk.XML")
    $VDsTotal   = NCount($VirtualDisks)
    $VDsHealthy = NCount($VirtualDisks |? { ($_.HealthStatus -like "Healthy") -or ($_.HealthStatus -eq 0) })
    Write-Host "Virtual Disks Healthy           : $VDsHealthy / $VDsTotal"
    if ($VDsHealthy -lt $VDsTotal) { Show-Warning "Unhealthy virtual disks detected" }
    $StoragePools = @(Import-ClixmlIf (Join-Path $Path "GetStoragePool.XML"))
    $PoolsTotal   = NCount($StoragePools)
    $PoolsHealthy = NCount($StoragePools |? { ($_.HealthStatus -like "Healthy") -or ($_.HealthStatus -eq 0) })
    Write-Host "Storage Pools Healthy           : $PoolsHealthy / $PoolsTotal "
    if ($S2DEnabled -and $StoragePools.Count -ne 1) { Show-Warning "S2D is enabled but pool count $($StoragePools.Count) != 1" }
    if ($PoolsTotal -lt $ExpectedPools) { Show-Warning "Fewer storage pools than the $ExpectedPools expected" }
    if ($PoolsHealthy -lt $PoolsTotal)  { Show-Warning "Unhealthy storage pools detected" }
    $PhysicalDisks = Import-ClixmlIf (Join-Path $Path "GetPhysicalDisk.XML")
    $PhysicalDiskSNV = Import-ClixmlIf (Join-Path $Path "GetPhysicalDiskSNV.XML")
    $PDsTotal   = NCount($PhysicalDisks)
    $PDsHealthy = NCount($PhysicalDisks |? { ($_.HealthStatus -like "Healthy") -or ($_.HealthStatus -eq 0) })
    Write-Host "Physical Disks Healthy          : $PDsHealthy / $PDsTotal"
    if ($PDsTotal -lt $ExpectedPhysicalDisks) { Show-Warning "Fewer physical disks than the $ExpectedPhysicalDisks expected" }
    if ($PDsHealthy -lt $PDsTotal) { Show-Warning "$($PDsTotal - $PDsHealthy) unhealthy physical disks detected" }
    $StorageEnclosures = Import-ClixmlIf (Join-Path $Path "GetStorageEnclosure.XML")
    $EncsTotal   = NCount($StorageEnclosures)
    $EncsHealthy = NCount($StorageEnclosures |? { ($_.HealthStatus -like "Healthy") -or ($_.HealthStatus -eq 0) })
    Write-Host "Storage Enclosures Healthy      : $EncsHealthy / $EncsTotal "
    if ($EncsTotal -lt $ExpectedEnclosures) { Show-Warning "Fewer storage enclosures than the $ExpectedEnclosures expected" }
    if ($EncsHealthy -lt $EncsTotal) { Show-Warning "Unhealthy storage enclosures detected" }
    if (-not (Test-Path (Join-Path $Path "GetReliabilityCounter.XML"))) {
        Write-Host "`nNOTE: storage device reliability counters not gathered for this capture."
    }

    Show-Update "<<< Phase 2 - Unhealthy Component Detail >>>`n" -ForegroundColor Cyan
    $Failed = $False
    if ($Cluster) {
        if ($NodesTotal -ne $NodesHealthy) {
            $Failed = $true; Write-Host "Cluster Nodes:"
            $ClusterNodes |? State -ne "Up" | Format-Table -AutoSize
        }
        if ($NetsTotal -ne $NetsHealthy) {
            $Failed = $true; Write-Host "Cluster Networks:"
            $ClusterNetworks |? State -ne "Up" | Format-Table -AutoSize
        }
        if ($ResTotal -ne $ResHealthy) {
            $Failed = $true; Write-Host "Cluster Resources:"
            $ClusterResources |? State -notlike "Online" | Format-Table Name,@{
                Label = 'State'; Expression = { $_.State.Value }}, OwnerGroup, ResourceType
        }
    } else {
        Show-Warning "Skipping cluster node, network and resource reporting since cluster was not available"
    }
    if ($SubsystemUnhealthy) {
        Write-Host "Clustered storage subsystem '$($Subsystem.FriendlyName)' not healthy:"
        Import-ClixmlIf (Join-Path $Path "DebugStorageSubsystem.XML") -MessageIf "Expected if cluster not available" | ft -AutoSize
    }
    if ($Cluster) {
        if ($CSVTotal -ne $CSVHealthy) {
            $Failed = $true; Write-Host "Cluster Shared Volumes not Online:"
            $CSV |? State -ne "Online" | Format-Table -AutoSize
        }
    }
    if ($VolsTotal -ne $VolsHealthy) {
        $Failed = $true; Write-Host "Cluster Shared Volumes not Healthy:"
        $Volumes |? { ($_.HealthStatus -notlike "Healthy") -and ($_.HealthStatus -ne 0) } | Format-Table Path,HealthStatus -AutoSize
    }
    if ($DedupEnabled -and $DedupTotal -ne $DedupHealthy) {
        $Failed = $true; Write-Host "Volumes:"
        $DedupVolumes |? LastOptimizationResult -eq 0 | Format-Table Volume,Capacity,SavingsRate,LastOptimizationResultMessage -AutoSize
    }
    if ($VDsTotal -ne $VDsHealthy) {
        $Failed = $true; Write-Host "Virtual Disks:"
        $VirtualDisks |? { ($_.HealthStatus -notlike "Healthy") -and ($_.HealthStatus -ne 0) } |
            Format-Table FriendlyName,HealthStatus,OperationalStatus,ResiliencySettingName,IsManualAttach -AutoSize
    }
    if ($PoolsTotal -ne $PoolsHealthy) {
        $Failed = $true; Write-Host "Storage Pools:"
        $StoragePools |? { ($_.HealthStatus -notlike "Healthy") -and ($_.HealthStatus -ne 0) } |
            Format-Table FriendlyName,HealthStatus,OperationalStatus,IsReadOnly -AutoSize
    }
    if ($PDsTotal -ne $PDsHealthy) {
        $Failed = $true; Write-Host "Physical Disks:"
        $PhysicalDisks |? { ($_.HealthStatus -notlike "Healthy") -and ($_.HealthStatus -ne 0) } |
            Format-Table FriendlyName,EnclosureNumber,SlotNumber,HealthStatus,OperationalStatus,Usage -AutoSize
    }
    if ($EncsTotal -ne $EncsHealthy) {
        $Failed = $true; Write-Host "Enclosures:"
        $StorageEnclosures |? { ($_.HealthStatus -notlike "Healthy") -and ($_.HealthStatus -ne 0) } |
            Format-Table FriendlyName,HealthStatus,ElementTypesInError -AutoSize
    }
    if ($ShTotal -ne $ShHealthy) {
        $Failed = $true; Write-Host "CA Shares:"
        $ShareStatus |? Health -notlike "Healthy" | Format-Table -AutoSize
    }
    if (-not $Failed) { "No unhealthy components`n" }

    Show-Update "<<< Phase 3 - Firmware and drivers >>>`n" -ForegroundColor Cyan
    foreach ($node in $ClusterNodes.Name) {
        "`nCluster Node: $node"
        Import-ClixmlIf (Join-Path (Get-NodePath $Path $node) "GetDrivers.XML") |? {
            ($_.DeviceCLass -eq 'SCSIADAPTER') -or ($_.DeviceCLass -eq 'NET')
        } | Group-Object DeviceName,DriverVersion | Sort Name | ft -AutoSize Count,
            @{ Expression = { $_.Group[0].DeviceName }; Label = "DeviceName" },
            @{ Expression = { $_.Group[0].DriverVersion }; Label = "DriverVersion" },
            @{ Expression = { $_.Group[0].DriverDate }; Label = "DriverDate" }
    }
    Write-Host "`nPhysical disks by Media Type, Model and Firmware Version"
    $PhysicalDisks | Group-Object MediaType,Model,FirmwareVersion | ft -AutoSize Count,
        @{ Expression = { $_.Group[0].Model }; Label="Model" },
        @{ Expression = { $_.Group[0].FirmwareVersion }; Label="FirmwareVersion" },
        @{ Expression = { $_.Group[0].MediaType }; Label="MediaType" }
    Write-Host "Storage Enclosures by Model and Firmware Version"
    $StorageEnclosures | Group-Object Model,FirmwareVersion | ft -AutoSize Count,
        @{ Expression = { $_.Group[0].Model }; Label="Model" },
        @{ Expression = { $_.Group[0].FirmwareVersion }; Label="FirmwareVersion" }
}

# endregion Report functions

###############################################################################
# region Invoke-GetDellSDDC — the main function (renamed from Invoke-GetDellSDDC)
###############################################################################

function Invoke-GetDellSDDC {

    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidUsingCmdletAliases", "")]
    [CmdletBinding(DefaultParameterSetName="WriteC")]
    [OutputType([String])]
    param(
        [parameter(ParameterSetName="WriteC", Position=0, Mandatory=$false)]
        [parameter(ParameterSetName="WriteN", Position=0, Mandatory=$false)]
        [alias("WriteToPath")]
        [ValidateNotNullOrEmpty()]
        [string] $TemporaryPath = $($env:userprofile + "\HealthTest\"),

        [parameter(ParameterSetName="M", Position=1, Mandatory=$false)]
        [parameter(ParameterSetName="WriteC", Position=1, Mandatory=$false)]
        [string] $ClusterName = ".",

        [parameter(ParameterSetName="WriteN", Position=1, Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string[]] $Nodelist = @(),

        [parameter(ParameterSetName="Read", Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string] $ReadFromPath = "",

        [parameter(ParameterSetName="WriteC", Mandatory=$false)]
        [parameter(ParameterSetName="WriteN", Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
        [bool] $IncludePerformance = $true,

        [parameter(ParameterSetName="WriteC", Mandatory=$false)]
        [parameter(ParameterSetName="WriteN", Mandatory=$false)]
        [ValidateRange(1,3600)]
        [int] $PerfSamples = 30,

        [parameter(ParameterSetName="WriteC", Mandatory=$false)]
        [parameter(ParameterSetName="WriteN", Mandatory=$false)]
        [switch] $ProcessCounter,

        [parameter(ParameterSetName="M", Mandatory=$true)]
        [switch] $MonitoringMode,

        [parameter(ParameterSetName="WriteC", Mandatory=$false)]
        [parameter(ParameterSetName="WriteN", Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
        [int] $HoursOfEvents = 168,

        [parameter(ParameterSetName="WriteC", Mandatory=$false)]
        [parameter(ParameterSetName="WriteN", Mandatory=$false)]
        [ValidateRange(-1,365)]
        [int] $DaysOfArchive = 8,

        [parameter(ParameterSetName="WriteC", Position=2, Mandatory=$false)]
        [parameter(ParameterSetName="WriteN", Position=2, Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
        [string] $ZipPrefix = $($env:userprofile + "\HealthTest"),

        [parameter(ParameterSetName="WriteC", Mandatory=$false)]
        [parameter(ParameterSetName="WriteN", Mandatory=$false)]
        [ValidateRange(1,1000)]
        [int] $ExpectedNodes,

        [parameter(ParameterSetName="WriteC", Mandatory=$false)]
        [parameter(ParameterSetName="WriteN", Mandatory=$false)]
        [ValidateRange(1,1000)]
        [int] $ExpectedNetworks,

        [parameter(ParameterSetName="WriteC", Mandatory=$false)]
        [parameter(ParameterSetName="WriteN", Mandatory=$false)]
        [ValidateRange(0,1000)]
        [int] $ExpectedVolumes,

        [parameter(ParameterSetName="WriteC", Mandatory=$false)]
        [parameter(ParameterSetName="WriteN", Mandatory=$false)]
        [ValidateRange(0,1000)]
        [int] $ExpectedDedupVolumes,

        [parameter(ParameterSetName="WriteC", Mandatory=$false)]
        [parameter(ParameterSetName="WriteN", Mandatory=$false)]
        [ValidateRange(1,10000)]
        [int] $ExpectedPhysicalDisks,

        [parameter(ParameterSetName="WriteC", Mandatory=$false)]
        [parameter(ParameterSetName="WriteN", Mandatory=$false)]
        [ValidateRange(1,1000)]
        [int] $ExpectedPools,

        [parameter(ParameterSetName="WriteC", Mandatory=$false)]
        [parameter(ParameterSetName="WriteN", Mandatory=$false)]
        [ValidateRange(1,10000)]
        [int] $ExpectedEnclosures,

        [parameter(ParameterSetName="WriteC", Mandatory=$false)]
        [parameter(ParameterSetName="WriteN", Mandatory=$false)]
        [switch] $IncludeAssociations,

        [parameter(ParameterSetName="WriteC", Mandatory=$false)]
        [parameter(ParameterSetName="WriteN", Mandatory=$false)]
        [switch] $IncludeDumps,

        [parameter(ParameterSetName="WriteC", Mandatory=$false)]
        [parameter(ParameterSetName="WriteN", Mandatory=$false)]
        [switch] $IncludeGetNetView,

        [parameter(ParameterSetName="WriteC", Mandatory=$false)]
        [parameter(ParameterSetName="WriteN", Mandatory=$false)]
        [switch] $SkipVM,

        [parameter(ParameterSetName="WriteC", Mandatory=$false)]
        [parameter(ParameterSetName="WriteN", Mandatory=$false)]
        [switch] $IncludeHealthReport,

        [parameter(ParameterSetName="WriteC", Mandatory=$false)]
        [parameter(ParameterSetName="WriteN", Mandatory=$false)]
        [switch] $IncludeClusterPerformanceHistory,

        [parameter(ParameterSetName="WriteC", Mandatory=$false)]
        [parameter(ParameterSetName="WriteN", Mandatory=$false)]
        [ValidateSet('LastHour','LastDay','LastWeek','LastMonth','LastYear')]
        [string] $PerformanceHistoryTimeFrame = "LastDay",

        [parameter(ParameterSetName="WriteC", Mandatory=$false)]
        [parameter(ParameterSetName="WriteN", Mandatory=$false)]
        [switch] $IncludeLiveDump,

        [parameter(ParameterSetName="WriteC", Mandatory=$false)]
        [parameter(ParameterSetName="WriteN", Mandatory=$false)]
        [switch] $IncludeStorDiag,

        [parameter(ParameterSetName="WriteC", Mandatory=$false)]
        [parameter(ParameterSetName="WriteN", Mandatory=$false)]
        [switch] $IncludeProcessDump,

        [parameter(ParameterSetName="WriteC", Mandatory=$false)]
        [parameter(ParameterSetName="WriteN", Mandatory=$false)]
        [string] $Processlists = "",

        [parameter(ParameterSetName="WriteC", Mandatory=$false)]
        [parameter(ParameterSetName="WriteN", Mandatory=$false)]
        [switch] $IncludeReliabilityCounters,

        [parameter(ParameterSetName="WriteC", Mandatory=$false)]
        [parameter(ParameterSetName="WriteN", Mandatory=$false)]
        [switch] $RunCluChk,

        [parameter(ParameterSetName="WriteC", Mandatory=$false)]
        [parameter(ParameterSetName="WriteN", Mandatory=$false)]
        [string] $SessionConfigurationName = $null
    )

    Set-StrictMode -Version Latest

    # CONVERSION: use script-scope variables instead of Get-Module
    $Module = $script:ModuleName

    #region Inner helper functions (Volume/Association lookups)
    function VolumeToPath    { Param ([String] $Volume) if ($null -eq $Associations) { Show-Error("No device associations present.") } $Result = ""; $Associations |% { if ($_.VolumeID -eq $Volume) { $Result = $_.CSVPath } }; return $Result }
    function VolumeToCSV     { Param ([String] $Volume) if ($null -eq $Associations) { Show-Error("No device associations present.") } $Result = ""; $Associations |% { if ($_.VolumeID -eq $Volume) { $Result = $_.CSVVolume } }; return $Result }
    function VolumeToVD      { Param ([String] $Volume) if ($null -eq $Associations) { Show-Error("No device associations present.") } $Result = ""; $Associations |% { if ($_.VolumeID -eq $Volume) { $Result = $_.FriendlyName } }; return $Result }
    function VolumeToShare   { Param ([String] $Volume) if ($null -eq $Associations) { Show-Error("No device associations present.") } $Result = ""; $Associations |% { if ($_.VolumeID -eq $Volume) { $Result = $_.ShareName } }; return $Result }
    function VolumeToResiliency { Param ([String] $Volume) if ($null -eq $Associations) { Show-Error("No device associations present.") } $Result = ""; $Associations |% { if ($_.VolumeID -eq $Volume) { $Result = $_.VDResiliency+","+$_.VDCopies; if ($_.VDEAware) { $Result += ",E" } else { $Result += ",NE" } } }; return $Result }
    function VolumeToColumns { Param ([String] $Volume) if ($null -eq $Associations) { Show-Error("No device associations present.") } $Result = ""; $Associations |% { if ($_.VolumeID -eq $Volume) { $Result = $_.VDColumns } }; return $Result }
    function CSVToShare      { Param ([String] $Volume) if ($null -eq $Associations) { Show-Error("No device associations present.") } $Result = ""; $Associations |% { if ($_.CSVVolume -eq $Volume) { $Result = $_.ShareName } }; return $Result }
    function VolumeToPool    { Param ([String] $Volume) if ($null -eq $Associations) { Show-Error("No device associations present.") } $Result = ""; $Associations |% { if ($_.VolumeId -eq $Volume) { $Result = $_.PoolName } }; return $Result }
    function CSVToVD         { Param ([String] $Volume) if ($null -eq $Associations) { Show-Error("No device associations present.") } $Result = ""; $Associations |% { if ($_.CSVVolume -eq $Volume) { $Result = $_.FriendlyName } }; return $Result }
    function CSVToPool       { Param ([String] $Volume) if ($null -eq $Associations) { Show-Error("No device associations present.") } $Result = ""; $Associations |% { if ($_.CSVVolume -eq $Volume) { $Result = $_.PoolName } }; return $Result }
    function CSVToNode       { Param ([String] $Volume) if ($null -eq $Associations) { Show-Error("No device associations present.") } $Result = ""; $Associations |% { if ($_.CSVVolume -eq $Volume) { $Result = $_.CSVNode } }; return $Result }
    function VolumeToCSVName { Param ([String] $Volume) if ($null -eq $Associations) { Show-Error("No device associations present.") } $Result = ""; $Associations |% { if ($_.VolumeId -eq $Volume) { $Result = $_.CSVName } }; return $Result }
    function CSVStatus       { Param ([String] $Volume) if ($null -eq $Associations) { Show-Error("No device associations present.") } $Result = ""; $Associations |% { if ($_.VolumeId -eq $Volume) { $Result = $_.CSVStatus.Value } }; return $Result }
    function PoolOperationalStatus { Param ([String] $Volume) if ($null -eq $Associations) { Show-Error("No device associations present.") } $Result = ""; $Associations |% { if ($_.VolumeId -eq $Volume) { $Result = $_.PoolOpStatus } }; return $Result }
    function PoolHealthStatus { Param ([String] $Volume) if ($null -eq $Associations) { Show-Error("No device associations present.") } $Result = ""; $Associations |% { if ($_.VolumeId -eq $Volume) { $Result = $_.PoolHealthStatus } }; return $Result }
    function PoolHealthyPDs  { Param ([String] $PoolName) $healthyPDs = ""; if ($PoolName) { $totalPDs = (Get-StoragePool -FriendlyName $PoolName -CimSession $ClusterName -ErrorAction SilentlyContinue | Get-PhysicalDisk).Count; $healthyPDs = (Get-StoragePool -FriendlyName $PoolName -CimSession $ClusterName -ErrorAction SilentlyContinue | Get-PhysicalDisk |? HealthStatus -eq "Healthy" ).Count } else { Show-Error("No storage pool specified") }; return "$totalPDs/$healthyPDs" }
    function VDOperationalStatus { Param ([String] $Volume) if ($null -eq $Associations) { Show-Error("No device associations present.") } $Result = ""; $Associations |% { if ($_.VolumeId -eq $Volume) { $Result = $_.OperationalStatus } }; return $Result }
    function VDHealthStatus  { Param ([String] $Volume) if ($null -eq $Associations) { Show-Error("No device associations present.") } $Result = ""; $Associations |% { if ($_.VolumeId -eq $Volume) { $Result = $_.HealthStatus } }; return $Result }
    #endregion Inner helper functions

    # do not allow running in a remote powershell session
    If (Get-Variable PSSenderInfo -ErrorAction SilentlyContinue) {
        Show-Error "This function is not supported using a remote powershell session. Please run locally"
    }

    $OS = Get-CimInstance -ClassName Win32_OperatingSystem
    $S2DEnabled = $false
    if ([uint64]$OS.BuildNumber -lt 14393) {
        Show-Error("Wrong OS Version - Need at least Windows Server 2016. BuildNumber $($OS.BuildNumber)")
    }
    if (-not (Get-Module -ListAvailable FailoverClusters)) {
        Show-Error("Cluster PowerShell not available. Download the Windows Failover Clustering RSAT tools.")
    }

    function StartMonitoring {
        Show-Update "Entered continuous monitoring mode." -ForegroundColor Yellow
        Show-Update "Press Ctrl + C to stop monitoring" -ForegroundColor Yellow
        try { $ClusterName = (Get-Cluster -Name $ClusterName).Name }
        catch { Show-Error("Cluster could not be contacted. `nError="+$_.Exception.Message) }
        $NodeList = Get-NodeList -Cluster $ClusterName -Filter
        $AccessNode = Get-ClusterAccessNode @($NodeList)
        if ($AccessNode -ne $null) { $AccessNode = $AccessNode + "." + (Get-Cluster -Name $AccessNode).Domain }
        try { $Volumes = Get-Volume -CimSession $AccessNode }
        catch { Show-Error("Unable to get Volumes. `nError="+$_.Exception.Message) }
        $AssocJob = Start-Job -ArgumentList $AccessNode,$ClusterName {
            param($AccessNode,$ClusterName)
            $SmbShares = Get-SmbShare -CimSession $AccessNode
            $Associations = Get-VirtualDisk -CimSession $AccessNode |% {
                $o = $_ | Select-Object FriendlyName, CSVName, CSVNode, CSVPath, CSVFS,CSVVolume, ShareName, SharePath, VolumeID, PoolName, VDResiliency, VDCopies, VDColumns, VDEAware
                $AssocCSV = $_ | Get-ClusterSharedVolume -Cluster $ClusterName
                if ($AssocCSV) {
                    $o.CSVName = $AssocCSV.Name; $o.CSVNode = $AssocCSV.OwnerNode.Name
                    $o.CSVPath = $AssocCSV.SharedVolumeInfo.FriendlyVolumeName
                    $o.CSVFS = ($_ | Get-Disk | Get-Partition | Get-Volume).FileSystemType
                    if ($o.CSVPath.Length -ne 0) { $o.CSVVolume = $o.CSVPath.Split("\")[2] }
                    $AssocLike = $o.CSVPath+"\*"
                    $AssocShares = $SmbShares |? Path -like $AssocLike
                    $AssocShare = $AssocShares | Select-Object -First 1
                    if ($AssocShare) {
                        $o.ShareName = $AssocShare.Name; $o.SharePath = $AssocShare.Path; $o.VolumeID = $AssocShare.Volume
                        if ($AssocShares.Count -gt 1) { $o.ShareName += "*" }
                    }
                }
                Write-Output $o
            }
            $AssocPool = Get-StoragePool -CimSession $AccessNode -ErrorAction SilentlyContinue
            $AssocPool |% {
                $AssocPName = $_.FriendlyName
                Get-StoragePool -CimSession $AccessNode -FriendlyName $AssocPName | Get-VirtualDisk -CimSession $AccessNode |% {
                    $AssocVD = $_
                    $Associations |% {
                        if ($_.FriendlyName -eq $AssocVD.FriendlyName) {
                            $_.PoolName = $AssocPName; $_.VDResiliency = $AssocVD.ResiliencySettingName
                            $_.VDCopies = $AssocVD.NumberofDataCopies; $_.VDColumns = $AssocVD.NumberofColumns
                            $_.VDEAware = $AssocVD.IsEnclosureAware
                        }
                    }
                }
            }
            Write-Output $Associations
        }
        $Associations = $AssocJob | Wait-Job | Receive-Job
        $AssocJob | Remove-Job
        [System.Console]::Clear()
        $Volumes |? FileSystem -eq CSVFS | Sort-Object SizeRemaining | Format-Table -AutoSize `
            @{Expression={$poolName = VolumeToPool($_.Path); "[$(PoolOperationalStatus($_.Path))/$(PoolHealthStatus($_.Path))] " + $poolName};Label="[OpStatus/Health] Pool"},
            @{Expression={(PoolHealthyPDs(VolumeToPool($_.Path)))};Label="HealthyPhysicalDisks"; Align="Center"},
            @{Expression={$vd = VolumeToVD($_.Path); "[$(VDOperationalStatus($_.Path))/$(VDHealthStatus($_.Path))] "+$vd};Label="[OpStatus/Health] VirtualDisk"},
            @{Expression={$csvVolume = VolumeToCSV($_.Path); "[" + $_.HealthStatus + "] " + $csvVolume};Label="[Health] CSV Volume"},
            @{Expression={$csvName = VolumeToCSVName($_.Path); $csvStatus = CSVStatus($_.Path); " [$csvStatus] " + $csvName};Label="[Status] CSV Name"},
            @{Expression={CSVToNode(VolumeToCSV($_.Path))};Label="Volume Owner"},
            @{Expression={VolumeToShare($_.Path)};Label="Share Name"},
            @{Expression={$VolResiliency = VolumeToResiliency($_.Path); $volColumns = VolumeToColumns($_.Path); "$VolResiliency,$volColumns" +"Col" };Label="Volume Configuration"},
            @{Expression={"{0:N2}" -f ($_.Size/1GB)};Label="Total Size";Width=11;Align="Right"},
            @{Expression={"{0:N2}" -f ($_.SizeRemaining/$_.Size*100)};Label="Avail%";Width=11;Align="Right"}
        StartMonitoring
    }

    If (-not (([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))) {
        Write-Error "Please run this function as an administrator"; exit
    }

    if ($MonitoringMode) { StartMonitoring }

    if (-not (Test-PrefixFilePath ([ref] $ZipPrefix))) {
        Write-Error "$ZipPrefix is not a valid prefix for ZIP: $ZipPrefix.ZIP must be creatable"
        return
    }

    if ($ReadFromPath -ne "") {
        $Path = $ReadFromPath
        $Read = $true
    } else {
        $Path = $TemporaryPath
        $Read = $false
    }

    if ($Read) {
        $Path = Check-ExtractZip $Path
    } else {
        Remove-Item -Path $Path -ErrorAction SilentlyContinue -Recurse | Out-Null
        New-Item -ItemType Directory -ErrorAction SilentlyContinue $Path | Out-Null
        compact /c $Path | Out-Null
    }

    $PathObject = Get-Item $Path
    if ($null -eq $PathObject) { Show-Error ("Path not found: $Path") }
    $Path = $PathObject.FullName
    if (-not $Path.EndsWith("\")) { $Path = $Path + "\" }

    if ($Read) {
        Show-SddcDiagnosticReport -Report Summary -ReportLevel Full $Path
        return
    }

    # Start Transcript
    $transcriptFile = Join-Path $Path "0_CloudHealthGatherTranscript.log"
    try {
        Start-Transcript -Path $transcriptFile -Force
        Write-host "Dell SDDC Version"
    } catch {
        Show-Error "Unable to start transcript at $transcriptFile" $_
        throw $_
    }

    try {

    Show-Update "Temporary write path : $Path"

    $Parameters = "" | Select-Object TodayDate, ExpectedNodes, ExpectedNetworks, ExpectedVolumes,
        ExpectedPhysicalDisks, ExpectedPools, ExpectedEnclosures, ExpectedDedupVolumes, HoursOfEvents, Version
    $TodayDate = Get-Date
    $Parameters.TodayDate = $TodayDate
    $Parameters.ExpectedNodes = $ExpectedNodes
    $Parameters.ExpectedNetworks = $ExpectedNetworks
    $Parameters.ExpectedVolumes = $ExpectedVolumes
    $Parameters.ExpectedDedupVolumes = $ExpectedDedupVolumes
    $Parameters.ExpectedPhysicalDisks = $ExpectedPhysicalDisks
    $Parameters.ExpectedPools = $ExpectedPools
    $Parameters.ExpectedEnclosures = $ExpectedEnclosures
    $Parameters.HoursOfEvents = $HoursOfEvents

    # CONVERSION: replaced (Get-Module $Module).Version.ToString() with script variable
    $Parameters.Version = $script:ScriptVersion

    $Parameters | Export-Clixml ($Path + "GetParameters.XML")

    Show-Update "Invoke-GetDellSDDC v $($Parameters.Version)"

    Show-Update "<<< Phase 1 - Data Gather >>>`n" -ForegroundColor Cyan

    try {
        $ClusterNodes = Get-NodeList -Cluster $ClusterName -Nodes $Nodelist
    } catch {
        Show-Error "Unable to get Cluster Nodes for reporting" $_
    }
    $ClusterNodes | Export-Clixml ($Path + "GetClusterNode.XML")

    try {
        $ClusterNodes = Get-NodeList -Cluster $ClusterName -Nodes $Nodelist -Filter
    } catch {
        Show-Error "Unable to get filtered Cluster Nodes for gathering" $_
    }

    $AccessNode = Get-ClusterAccessNode @($ClusterNodes)

    try {
        if ($ClusterName -eq ".") {
            foreach ($cn in $ClusterNodes) {
                $Cluster = Get-Cluster -Name $cn.Name -ErrorAction SilentlyContinue
                if ($Cluster -ne $null) { break }
            }
        } else {
            $Cluster = Get-Cluster -Name $ClusterName
        }
    } catch {
        Show-Error("Cluster could not be contacted. `nError="+$_.Exception.Message)
    }

    if ($Cluster -ne $null) {
        $Cluster | Export-Clixml ($Path + "GetCluster.XML")
        $ClusterName = $Cluster.Name + "." + $Cluster.Domain
        $S2DEnabled = $Cluster.S2DEnabled
        $ClusterDomain = $Cluster.Domain
        Write-Host "Cluster name         : $ClusterName"
    } else {
        Show-Warning "Cluster service was not running on any node, some information will be unavailable"
        $ClusterName = ''
        $ClusterDomain = ''
        Write-Host "Cluster name         : Unavailable, Cluster is not online on any node"
    }

    Write-Host ("Accessible Node List : " + [string]::Join(", ",$ClusterNodes.name))
    Write-Host "Access node          : $AccessNode`n"

    $ClusterNodes.Name |% { md (Get-NodePath $Path $_) | Out-Null }

    $DedupEnabled = $true
    if ($(Invoke-Command -ComputerName $AccessNode -ConfigurationName $SessionConfigurationName {(-not (Get-Command -Module Deduplication))} )) {
        $DedupEnabled = $false
    }

    $JobStatic = @()
    $JobCopyOut = @()
    $JobCopyOutNoDelete = @()

    # Sddc Diagnostic Archive capture
    if ($Cluster -and (Get-ClusteredScheduledTask -Cluster $Cluster -TaskName SddcDiagnosticArchive)) {
        if ($DaysOfArchive -gt 0) {
            Show-Update "Start gather of Sddc Diagnostic Archives ..."
            $JobStatic += Start-Job -Name 'Sddc Diagnostic Archive Report' {
                Import-Module $using:Module -ErrorAction SilentlyContinue
                $o = (Join-Path $using:Path SddcDiagnosticArchiveJob.txt)
                Show-SddcDiagnosticArchiveJob -Cluster $using:Cluster > $o
                $o = (Join-Path $using:Path SddcDiagnosticArchiveJobWarn.txt)
                $null = Confirm-SddcDiagnosticModule -Cluster $using:Cluster 3> $o
            }
            $j = Invoke-SddcCommonCommand -ClusterNodes $ClusterNodes.Name -JobName SddcDiagnosticArchive -SessionConfigurationName $SessionConfigurationName -InitBlock $CommonFunc {
                Import-Module $using:Module -ErrorAction SilentlyContinue
                if (Test-SddcModulePresence) {
                    $Path = $null
                    Get-SddcDiagnosticArchiveJobParameters -Path ([ref] $Path)
                    & {
                        if ($using:DaysOfArchive -ne -1) {
                            $Archive = dir $Path\*.ZIP | sort -Descending
                            if ($Archive.Count -gt $using:DaysOfArchive) {
                                $Archive = $Archive[0..$($using:DaysOfArchive - 1)]
                            }
                            $Archive.FullName
                            (dir $Path\*.log).FullName
                        } else {
                            Join-Path (gi $Path).FullName "*"
                        }
                    } |% { Get-AdminSharePathFromLocal $env:COMPUTERNAME $_ }
                }
            }
            $j.ChildJobs |% { $_ | Add-Member -NotePropertyName Destination -NotePropertyValue SddcDiagnosticArchive }
            $JobCopyOutNoDelete += $j
        }
    }

    if ($AccessNode) {
        Show-Update "Start gather of cluster configuration ..."

        $JobStatic += start-job -Name ClusterGroup {
            try { $o = Get-ClusterGroup -Cluster $using:AccessNode; $o | Export-Clixml ($using:Path + "GetClusterGroup.XML") }
            catch { Write-Warning "Unable to get Cluster Groups. `nError=$($_.Exception.Message)" }
        }
        $JobStatic += start-job -Name ClusterNetwork {
            try { $o = Get-ClusterNetwork -Cluster $using:AccessNode; $o | Export-Clixml ($using:Path + "GetClusterNetwork.XML") }
            catch { Write-Warning "Could not get Cluster Networks. `nError=$($_.Exception.Message)" }
        }
        $JobStatic += start-job -Name ClusterNetworkLiveMigrationInformation {
            try { $o = Get-ClusterResourceType -Name 'Virtual Machine' -Cluster $using:AccessNode | Get-ClusterParameter; $o | Export-Clixml ($using:Path + "ClusterNetworkLiveMigration.XML") }
            catch { Write-Warning "Could not get Cluster Network Live Migration Information. `nError=$($_.Exception.Message)" }
        }
        $JobStatic += start-job -Name ClusterResource {
            try { $o = Get-ClusterResource -Cluster $using:AccessNode; $o | Export-Clixml ($using:Path + "GetClusterResource.XML") }
            catch { Write-Warning "Unable to get Cluster Resources. `nError=$($_.Exception.Message)" }
        }
        $JobStatic += start-job -Name ClusterResourceParameter {
            try { $o = Get-ClusterResource -Cluster $using:AccessNode | Get-ClusterParameter; $o | Export-Clixml ($using:Path + "GetClusterResourceParameters.XML") }
            catch { Write-Warning "Unable to get Cluster Resource Parameters. `nError=$($_.Exception.Message)" }
        }
        $JobStatic += start-job -Name ClusterSharedVolume {
            try { $o = Get-VirtualDisk | %{$vd=$_; $vd | Get-ClusterSharedVolume -Cluster $using:AccessNode | Select *,@{L="VDID";E={$vd.UniqueId}}}; $o | Export-Clixml ($using:Path + "GetClusterSharedVolume.XML") }
            catch { Write-Warning "Unable to get Cluster Shared Volumes. `nError=$($_.Exception.Message)" }
        }
        $JobStatic += start-job -Name GetClusterFaultDomain {
            try { $o = Get-ClusterFaultDomain | select name,type,parentname,childrennames,location; $o | Export-Clixml ($using:Path + "GetClusterFaultDomain.XML") }
            catch { Write-Warning "Unable to get ClusterFaultDomain. `nError=$($_.Exception.Message)" }
        }
        $JobStatic += start-job -Name GetClusterAffinityRule {
            try { $o = Get-ClusterAffinityRule | select *; $o | Export-Clixml ($using:Path + "GetClusterAffinityRule.XML") }
            catch { Write-Warning "Unable to get ClusterAffinityRule. `nError=$($_.Exception.Message)" }
        }
        $JobStatic += start-job -Name ClusterNodeSupportedVersion {
            try { $o = Get-ClusterNodeSupportedVersion; $o | Export-Clixml ($using:Path + "GetClusterNodeSupportedVersion.XML") }
            catch { Write-Warning "Unable to get Cluster Node Supported Version `nError=$($_.Exception.Message)" }
        }
        $JobStatic += start-job -Name GetCauClusterRole {
            try { $o = Get-CauClusterRole; $o | Export-Clixml ($using:Path + "GetCauClusterRole.XML") }
            catch { Write-Warning "Unable to get CAU Cluster Role `nError=$($_.Exception.Message)" }
        }

        Show-Update "Start gather of Network ATC information..."
        $NetworkATC = $False
        $NetworkATC = try {(Get-WindowsFeature NetworkATC).installed} catch {$False}
        if ($NetworkATC) {
            $JobStatic += start-job -Name NetIntentStatus {
                try { $o = Get-NetIntentStatus -ClusterName $using:AccessNode; $o | Export-Clixml ($using:Path + "GetNetIntentStatus.XML") }
                catch { Write-Warning "Unable to get NetIntentStatus. `nError=$($_.Exception.Message)" }
            }
            $JobStatic += start-job -Name NetIntentStatusGlobalOverrides {
                try { $o = Get-NetIntentStatus -GlobalOverrides -ClusterName $using:AccessNode; $o | Export-Clixml ($using:Path + "GetNetIntentStatusGlobalOverrides.XML") }
                catch { Write-Warning "Unable to get NetIntentStatus -GlobalOverrides. `nError=$($_.Exception.Message)" }
            }
            $JobStatic += start-job -Name NetIntent {
                try { $o = Get-NetIntent -ClusterName $using:AccessNode; $o | Export-Clixml ($using:Path + "GetNetIntent.XML") }
                catch { Write-Warning "Unable to get NetIntent. `nError=$($_.Exception.Message)" }
            }
            $JobStatic += start-job -Name NetIntentGlobalOverrides {
                try { $o = Get-NetIntent -GlobalOverrides -ClusterName $using:AccessNode; $o | Export-Clixml ($using:Path + "GetNetIntentGlobalOverrides.XML") }
                catch { Write-Warning "Unable to get NetIntent -GlobalOverrides. `nError=$($_.Exception.Message)" }
            }
        }
    } else {
        Show-Update "... Skip gather of cluster configuration since cluster is not available"
    }

    if ($IncludeClusterPerformanceHistory) {
        Show-Update "Starting ClusterPerformanceHistory log collection ..."
        $JobStatic += start-job -Name ClusterPerformanceHistory {
            try { Get-Clusterlog -ExportClusterPerformanceHistory -Destination $using:Path -PerformanceHistoryTimeFrame $using:PerformanceHistoryTimeFrame -Node $using:ClusterNodes.Name }
            catch { Write-Warning "Could not get ClusterPerformanceHistory. `nError=$($_.Exception.Message)" }
        }
    }

    Show-Update "Start gather of driver information ..."
    $ClusterNodes.Name |% {
        $node = $_
        $JobStatic += start-job -Name "Driver Information: $node" {
            try { $o = Get-CimInstance -ClassName Win32_PnPSignedDriver -ComputerName $using:node } catch {}
            $o | Export-Clixml (Join-Path (Join-Path $using:Path "Node_$using:node") "GetDrivers.XML")
        }
    }

    Show-Update "Start gather of verifier ..."
    $JobCopyOut += Invoke-SddcCommonCommand -ClusterNodes $($ClusterNodes).Name -JobName Verifier -SessionConfigurationName $SessionConfigurationName -InitBlock $CommonFunc {
        $LocalFile = Join-Path $env:temp "verifier-query.txt"
        verifier /query > $LocalFile
        Write-Output (Get-AdminSharePathFromLocal $env:COMPUTERNAME $LocalFile)
        $LocalFile = Join-Path $env:temp "verifier-querysettings.txt"
        verifier /querysettings > $LocalFile
        Write-Output (Get-AdminSharePathFromLocal $env:COMPUTERNAME $LocalFile)
    }

    Show-Update "Start gather of filesystem filter status ..."
    $JobCopyOut += Invoke-SddcCommonCommand -ClusterNodes $($ClusterNodes).Name -JobName 'Filesystem Filter Manager' -SessionConfigurationName $SessionConfigurationName -InitBlock $CommonFunc {
        $LocalFile = Join-Path $env:temp "fltmc.txt"
        fltmc > $LocalFile
        Write-Output (Get-AdminSharePathFromLocal $env:COMPUTERNAME $LocalFile)
        $filters = fltmc | ForEach-Object {
            $line = $_.Trim() -split "\s{2,}"
            if ($line.Length -ge 4 -and $line[0] -notmatch '^-+$' -and $line[0] -ne "Filter Name") {
                [PSCustomObject]@{ FilterName = $line[0]; NumInstances = $line[1]; Altitude = $line[2]; Frame = $line[3]; WindowsDriver = $line[0] + ".sys" }
            }
        }
        $filters | ForEach-Object {
            $driverPath = "C:\Windows\System32\drivers\$($_.WindowsDriver)"
            if (Test-Path $driverPath) {
                $_ | Add-Member -MemberType NoteProperty -Name Company -Value (Get-ItemProperty $driverPath).VersionInfo.CompanyName
                $_ | Add-Member -MemberType NoteProperty -Name Description -Value (Get-ItemProperty $driverPath).VersionInfo.FileDescription
            } else { $_ | Add-Member -MemberType NoteProperty -Name Company -Value "Unknown" }
        }
        $LocalFileXml = Join-Path $env:temp "fltmc.xml"
        $filters | Export-Clixml $LocalFileXml
        Write-Output (Get-AdminSharePathFromLocal $env:COMPUTERNAME $LocalFileXml)
        $LocalFile = Join-Path $env:temp "fltmc-instances.txt"
        fltmc instances > $LocalFile
        Write-Output (Get-AdminSharePathFromLocal $env:COMPUTERNAME $LocalFile)
    }

    $JobCopyOutNoDelete += Invoke-SddcCommonCommand -ClusterNodes $($ClusterNodes).Name -JobName 'Copy WER ReportArchive' -SessionConfigurationName $SessionConfigurationName -InitBlock $CommonFunc {
        Write-Output (Get-AdminSharePathFromLocal $env:COMPUTERNAME $env:ProgramData\Microsoft\Windows\WER\ReportArchive)
    }

    if ($IncludeDumps -eq $true) {
        $JobCopyOutNoDelete += Invoke-SddcCommonCommand -ClusterNodes $($ClusterNodes).Name -JobName 'Copy ReportQueue' -SessionConfigurationName $SessionConfigurationName -InitBlock $CommonFunc {
            Write-Output (Get-AdminSharePathFromLocal $env:COMPUTERNAME $env:ProgramData\Microsoft\Windows\WER\ReportQueue)
        }
    }

    if ($IncludeProcessDump) {
        $JobCopyOut += Invoke-SddcCommonCommand -ClusterNodes $($ClusterNodes).Name -JobName ProcessDumps -SessionConfigurationName $SessionConfigurationName -InitBlock $CommonFunc -ArgumentList $ProcessLists {
            Param($ProcessLists)
            $NodePath = $env:Temp
            $Node = $env:COMPUTERNAME
            $DumpProcesses = @("vmms", "vmcompute", "vmwp", "rhs", "clussvc")
            if ($ProcessLists -ne $null) { $DumpProcesses += $ProcessLists.split(",") }
            $DumpFileFolder = Join-Path -Path $NodePath -ChildPath 'ProcessDumps'
            if (Test-Path -Path $DumpFileFolder) { Remove-Item -Path $DumpFileFolder -Recurse -Force }
            $null = New-Item -Path $DumpFileFolder -ItemType Directory
            $WER = [PSObject].Assembly.GetType('System.Management.Automation.WindowsErrorReporting')
            $NativeMethods = $WER.GetNestedType('NativeMethods', 'NonPublic')
            $MiniDump = $NativeMethods.GetMethod('MiniDumpWriteDump', ([Reflection.BindingFlags]'NonPublic, Static'))
            $MiniDumpWithFullMemory = [UInt32] 2
            $ProcessList = @{}
            foreach ($ProcessName in $DumpProcesses) {
                $ProcessIds = Get-Process -Name $ProcessName -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id
                if (-not $ProcessIds) { Show-Warning "Could not generate minidump for process $ProcessName"; continue }
                foreach ($ProcessId in $ProcessIds) {
                    if ($ProcessList[$ProcessId]) { continue }
                    $ProcessList.Add($ProcessId, $ProcessName)
                    $Process = Get-Process -Id $ProcessId
                    $ProcessHandle = $Process.Handle
                    $DumpFileName = "$($ProcessName)_$($ProcessId).dmp"
                    $DumpFilePath = Join-Path $DumpFileFolder $DumpFileName
                    $DumpFile = New-Object IO.FileStream($DumpFilePath, [IO.FileMode]::Create)
                    $Result = $MiniDump.Invoke($null, @($ProcessHandle, $ProcessId, $DumpFile.SafeFileHandle, $MiniDumpWithFullMemory, [IntPtr]::Zero, [IntPtr]::Zero, [IntPtr]::Zero))
                    $DumpFile.Close()
                    if(-not $Result) {
                        Show-Warning "Failed to write dump file for process $ProcessName with PID $ProcessId."
                        Remove-Item $DumpFilePath
                    } else {
                        Write-Output (Get-AdminSharePathFromLocal $Node $DumpFilePath)
                    }
                }
            }
        }
    }

    if ($IncludeGetNetView) {
        Show-Update "Start gather of Get-NetView ..."
        $JobCopyOut += Invoke-SddcCommonCommand -ArgumentList $SkipVm -ClusterNodes $($ClusterNodes).Name -JobName 'GetNetView' -SessionConfigurationName $SessionConfigurationName -InitBlock $CommonFunc {
            Param($SkipVM)
            $NodePath = $env:Temp
            $gnvDir = Join-Path $NodePath 'GetNetView'
            Remove-Item -Recurse -Force $gnvDir -ErrorAction SilentlyContinue
            $null = md $gnvDir -Force -ErrorAction SilentlyContinue
            $j = Start-Job -ArgumentList $gnvDir,$SkipVM {
                param($gnvDir,$SkipVM)
                $transcriptFile = Join-Path $gnvDir "0_GetNetViewGatherTranscript.log"
                Start-Transcript -Path $transcriptFile -Force
                if (Get-Command Get-NetView -ErrorAction SilentlyContinue) {
                    if ($SkipVM) { Get-NetView -OutputDirectory $gnvDir -SkipLogs -SkipVM }
                    else { Get-NetView -OutputDirectory $gnvDir -SkipLogs }
                } else { Write-Host "Get-NetView command not available" }
                Stop-Transcript
            }
            $null = $j | Wait-Job
            $j | Remove-Job
            dir $gnvDir -Directory |% { Remove-Item -Recurse -Force $_.FullName }
            Write-Output (Get-AdminSharePathFromLocal $env:COMPUTERNAME $gnvDir)
        }
    }

    #region Events, cmd, reports, et.al.
    Show-Update "Start gather of system info, cluster/netft/health logs, reports and dump files ..."

    $RPath = (Get-AdminSharePathFromLocal $env:COMPUTERNAME $Path)
    If ($HoursOfEvents -eq -1) {$ClusterLogMinutes=999999} else {$ClusterLogMinutes=$HoursOfEvents*60}

    $JobStatic += Foreach ($NodeName in ($ClusterNodes.Name)) {
        Invoke-Command -AsJob -JobName "ClusterLogs$NodeName" -ComputerName $Nodename -ScriptBlock {
            try {$null=Get-ClusterLog -UseLocalTime -TimeSpan $using:ClusterLogMinutes} catch {}
        }
    }

    if ((Get-Command Get-ClusterLog).Parameters.ContainsKey("NetFt")) {
        $JobStatic += Foreach ($NodeName in ($ClusterNodes.Name)) {
            Invoke-Command -AsJob -JobName "ClusterLogsNetft$NodeName" -ComputerName $Nodename -ScriptBlock {
                $null=Get-ClusterLog -UseLocalTime -Netft -TimeSpan $using:ClusterLogMinutes
            }
        }
    }

    if ($S2DEnabled) {
        $JobStatic += Foreach ($NodeName in ($ClusterNodes.Name)) {
            Invoke-Command -AsJob -JobName "ClusterLogsHealth$NodeName" -ComputerName $Nodename -ScriptBlock {
                $null=Get-ClusterLog -UseLocalTime -Health -TimeSpan $using:ClusterLogMinutes
            }
        }
    }

    # Run Send-DiagnosticData if HciSvc exists on the node
    Foreach ($NodeName in $ClusterNodes) {
        if (Get-Service -ComputerName $NodeName -Name 'HciSvc' -ErrorAction SilentlyContinue) {
            Show-Update "Gathering Send-DiagnosticData for $($NodeName)..."
            $LocalNodeDir = Get-NodePath $Path $NodeName
            $CopySendDiag = Join-Path (Get-AdminSharePathFromLocal $AccessNode $LocalNodeDir) $NodeName
            Write-host "AccessNode: $AccessNode"
            $JobStatic += Invoke-Command -ComputerName $NodeName -JobName "SendDiagnosticData-$NodeName" -AsJob -ScriptBlock {
                param($CopySendDiag)
                $log = @()
                Write-host "[$env:COMPUTERNAME] Checking for HciSvc..."
                $svc = Get-Service -Name 'HciSvc' -ErrorAction SilentlyContinue
                if ($svc) {
                    Write-host "[$env:COMPUTERNAME] HciSvc found. Starting diagnostic..."
                    $RPath = 'C:\SendDiags'
                    Remove-Item $RPath -Recurse -Force -ErrorAction SilentlyContinue
                    New-Item -Path $RPath -ItemType Directory -Force -ErrorAction SilentlyContinue | Out-Null
                    try {
                        Send-DiagnosticData -SaveToPath $RPath -CollectSddc $false
                        Get-ChildItem "$Rpath\MOC_ARB_*.zip" -Recurse | sort name -Descending | select -Skip 1 | Remove-Item -Force -Confirm:$false
                    } catch { Write-host "[$env:COMPUTERNAME] ERROR: $_" }
                } else { Write-host "[$env:COMPUTERNAME] HciSvc not found. Skipping." }
                $log
            } -ArgumentList $CopySendDiag
        }
    }

    $JobStatic += $ClusterNodes.Name |% {
        $NodeName = $_
        Invoke-SddcCommonCommand -JobName "System Info: $NodeName" -InitBlock $CommonFunc -SessionConfigurationName $SessionConfigurationName -ScriptBlock {
            $Node = "$using:NodeName"
            if ($using:ClusterDomain.Length) { $Node += ".$using:ClusterDomain" }
            $LocalNodeDir = Get-NodePath $using:Path $using:NodeName

            $SysInfoOut=(Join-Path (Get-NodePath $using:Path $using:NodeName) "SystemInfo.TXT")
            Start-Process -FilePath "$env:comspec" -ArgumentList "/c SystemInfo.exe /S $using:NodeName > $SysInfoOut" -WindowStyle Minimized

            $LocalFileMsInfo = (Join-Path $LocalNodeDir "\msinfo.nfo")
            $msinfo=Start-Process C:\Windows\System32\msinfo32.exe -WindowStyle Minimized -ArgumentList "/computer $using:NodeName /nfo $LocalFileMsInfo" -PassThru

            $CmdsToLog = 'Get-HotFix -ComputerName _C_',
                'Get-NetAdapter -CimSession _C_',
                'Get-NetAdapterAdvancedProperty -CimSession _C_',
                'Get-NetAdapterBinding -CimSession _C_',
                'Get-NetAdapterChecksumOffload -CimSession _C_',
                'Get-NetAdapterIPsecOffload -CimSession _C_',
                'Get-NetAdapterLso -CimSession _C_',
                'Get-NetAdapterPacketDirect -CimSession _C_',
                'Get-NetAdapterRdma -CimSession _C_',
                'Get-NetAdapterRsc -CimSession _C_',
                'Get-NetAdapterRss -CimSession _C_',
                'Get-NetAdapterVmq -CimSession _C_',
                'Get-NetAdapterStatistics -CimSession _C_',
                'Get-NetIPv4Protocol -CimSession _C_',
                'Get-NetIPv6Protocol -CimSession _C_',
                'Get-NetIpAddress -CimSession _C_',
                'Get-NetLbfoTeam -CimSession _C_',
                'Get-NetLbfoTeamMember -CimSession _C_',
                'Get-NetLbfoTeamNic -CimSession _C_',
                'Get-NetOffloadGlobalSetting -CimSession _C_',
                'Get-NetPrefixPolicy -CimSession _C_',
                'Get-NetQosPolicy -CimSession _C_',
                'Get-NetAdapterQos -CimSession _C_',
                'Get-NetRoute -CimSession _C_',
                'Get-Disk -CimSession _C_',
                'Get-NetTcpConnection -CimSession _C_',
                'Get-NetTcpSetting -CimSession _C_',
                'Get-ScheduledTask -CimSession _C_ | Get-ScheduledTaskInfo -CimSession _C_',
                'Get-SmbServerNetworkInterface -CimSession _C_',
                'Get-StorageFaultDomain -CimSession _A_ -Type StorageScaleUnit |? FriendlyName -eq _N_ | Get-StorageFaultDomain -CimSession _A_',
                'Get-NetFirewallProfile -CimSession _C_',
                'Get-NetFirewallRule -CimSession _C_',
                'Get-NetConnectionProfile -CimSession _C_',
                'Get-SmbMultichannelConnection -CimSession _C_ -SmbInstance SBL',
                'Get-SmbClientConfiguration -CimSession _C_',
                'Get-SmbServerConfiguration -CimSession _C_',
                'Get-NetIPConfiguration -CimSession _C_',
                'Invoke-Command -ComputerName _C_ {Get-ComputerInfo}',
                'Invoke-Command -ComputerName _C_ {Get-ItemProperty -Path HKLM:\SYSTEM\CurrentControlSet\Services\spacePort\Parameters}',
                'Invoke-Command -ComputerName _C_ {Echo Get-RegSpacePortParameters;Get-ItemProperty -Path HKLM:\SYSTEM\CurrentControlSet\Services\spacePort\Parameters}',
                'Invoke-Command -ComputerName _C_ {Echo Get-RegOEMInformation;IF((Get-WmiObject -Class Win32_OperatingSystem).Caption -imatch "HCI"){Get-ItemProperty -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\OEMInformation}}',
                'Invoke-Command -ComputerName _C_ {Echo Get-netsh;netsh int tcp show global}',
                'Invoke-Command -ComputerName _C_ {Echo Get-win32_networkadapter;Get-WmiObject win32_networkadapter}',
                'Invoke-Command -ComputerName _C_ {Echo Get-TcpipParametersInterfaces;Get-ItemProperty -path HKLM:\System\CurrentControlSet\services\Tcpip\Parameters\Interfaces\*}',
                'Invoke-Command -ComputerName _C_ {Echo Get-mpioParameters;IF((Get-WindowsFeature -Name "Multipath-IO").Installed -eq "True"){Get-ItemProperty -path HKLM:\SYSTEM\CurrentControlSet\Services\mpio\Parameters}}',
                'Invoke-Command -ComputerName _C_ {Echo Get-mpioSettings;IF((Get-WindowsFeature -Name "Multipath-IO").Installed -eq "True"){Get-ItemProperty -path "HKLM:\SYSTEM\CurrentControlSet\Control\Class\{4d36e97b-e325-11ce-bfc1-08002be10318}\000*"}}',
                'Invoke-Command -ComputerName _C_ {Echo Get-MSDSMSupportedHW;IF((Get-WindowsFeature -Name "Multipath-IO").Installed -eq "True"){Get-MSDSMSupportedHW -CimSession _C_}}',
                'Invoke-Command -ComputerName _C_ {Echo Get-DriverSuiteVersion;Get-ChildItem HKLM:\SOFTWARE\Dell\MUP -Recurse | Get-ItemProperty}',
                'Invoke-Command -ComputerName _C_ {Echo Get-ChipsetVersion;Get-WmiObject win32_product | ? Name -like "*chipset*"}',
                'Invoke-Command -ComputerName _C_ {Echo Get-NetFirewallRule;Get-NetFirewallRule -All}',
                'Invoke-Command -ComputerName _C_ {Echo Get-ProcessByService;$aps=GPs;$r=@();$Ass=GWmi Win32_Service;foreach($p in $aps){$ss=$Ass|?{$_.ProcessID -eq $p.Id};IF($ss){$r+=[PSCustomObject]@{Service=$ss.DisplayName;ProcessName=$p.ProcessName;ProcessID=$p.Id}}}$r}',
                'Get-NetNeighbor -CimSession _C_',
                'Invoke-Command -ComputerName _C_ {Echo Get-CurrentVersion;Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion"}',
                'Invoke-Command -ComputerName _C_ {Get-WindowsFeature | Sort-Object -Property @{Expression="Installed";Descending=$true}, @{Expression="Name";Descending=$false} | Select-Object DisplayName, Name, Installed}',
                'Get-VMNetworkAdapterIsolation -ManagementOS -CimSession _C_',
                'Invoke-Command -ComputerName _C_ {Echo Get-gpresult;gpresult /Z}',
                'Get-DnsClientServerAddress -CimSession _C_'

            if (Get-Module DcbQos -ErrorAction SilentlyContinue) {
                $CmdsToLog += 'Invoke-Command -ComputerName _C_ {Echo Get-NetQosDcbxSettingPerNic;Get-NetAdapter | Get-NetQosDcbxSetting}',
                    'Get-NetQosDcbxSetting -CimSession _C_',
                    'Get-NetQosFlowControl -CimSession _C_',
                    'Get-NetQosTrafficClass -CimSession _C_'
            }
            if (Get-Module Hyper-V -ErrorAction SilentlyContinue) {
                $CmdsToLog += 'Get-VM -CimSession _C_ -ErrorAction SilentlyContinue | Select-Object *',
                    'Invoke-Command -ComputerName _C_ {Echo Get-vmprocessor;Get-VM -CimSession _C_ | Get-VMProcessor -ErrorAction SilentlyContinue | Select-Object *}',
                    'Get-VMNetworkAdapter -All -CimSession _C_ -ErrorAction SilentlyContinue | Select-Object *',
                    'Get-VMSwitch -CimSession _C_ -ErrorAction SilentlyContinue | Select-Object *',
                    'Echo Get-VMSwitchTeam; Get-VMSwitch -CimSession _C_ | Where-Object {$_.EmbeddedTeamingEnabled -eq $true} | %{Get-VMSwitchTeam -CimSession _C_ -SwitchName $_.name | Select-Object *}',
                    'Get-VMHost -CimSession _C_ -ErrorAction SilentlyContinue | Select-Object *',
                    'Get-VMNetworkAdapterVlan -CimSession _C_ -ManagementOS -ErrorAction SilentlyContinue | Select-Object *',
                    'Get-VMNetworkAdapterTeamMapping -CimSession _C_ -ManagementOS -ErrorAction SilentlyContinue | Select-Object *'
            }
            If (Get-Module Deduplication -ErrorAction SilentlyContinue){
                $clusterCimSession = New-CimSession -ComputerName $ClusterName
                $CmdsToLog += "Get-DedupVolume -CimSession $clusterCimSession"
            }
            If ((Get-WmiObject -Class Win32_OperatingSystem).Caption -imatch "HCI"){
                $CmdsToLog += "Get-AzureStackHCI"
                $CmdsToLog += "Get-AzureStackHCIArcIntegration"
            }
            IF(Invoke-Command -ComputerName $using:NodeName {gcm Get-StampInformation -ErrorAction SilentlyContinue}){
                $CmdsToLog += 'Invoke-Command -ComputerName _C_ {Get-StampInformation}',
                    'Invoke-Command -ComputerName _C_ {Get-SolutionUpdateEnvironment | Tee-Object -Variable GSUpdE >$Null;$GSUpdE | ForEach-Object{$_;$_.HealthCheckResult}}',
                    'Invoke-Command -ComputerName _C_ {Get-SolutionDiscoveryDiagnosticInfo | Tee-Object -Variable GSDDI >$Null ;$GSDDI|ForEach-Object{$_;$_.NonApplicableUpdates}}',
                    'Invoke-Command -ComputerName _C_ {Get-SolutionUpdate | Tee-Object -Variable GSUpd >$Null;$GSUpd | ForEach-Object{$_;$_.HealthCheckResult}}'
                if (test-path "C:\Observability\OEMDiagnostics") {
                    $LocalDiagsDir = Join-Path $LocalNodeDir "OEMDiagnostics"
                    $CmdsToLog += 'Invoke-Command -ComputerName _C_ {Echo Get-ActionplanInstanceToComplete;try {(Get-ActionPlanInstances | ? Status -ne "Completed" | Sort StartDateTime | Select -last 1).ProgressAsXml} catch {}}'
                }
            }

            $nodejobs=@()
            foreach ($cmd in $CmdsToLog) {
                $LocalFile = (Join-Path $LocalNodeDir ([regex]::match(($cmd.split() | Where-Object {$_ -imatch 'Get-'}),'Get-[a-zA-Z0-9]*').value -replace "-",""))
                try {
                    $cmdex = $cmd -replace '_C_',$using:NodeName -replace '_N_',$using:NodeName -replace '_A_',$using:AccessNode
                    $cmdsb = [scriptblock]::Create("$cmdex")
                    $nodejobs+=Start-Job -Name $LocalFile -ScriptBlock $cmdsb
                } catch {}
            }

            $NodeSystemRootPath = Invoke-Command -ComputerName $using:NodeName -ConfigurationName $using:SessionConfigurationName { $env:SystemRoot }
            $NodeSystemDrivePath = Invoke-Command -ComputerName $using:NodeName -ConfigurationName $using:SessionConfigurationName { $env:SystemDrive }

            Show-Update "Gathering Mini Dump and Live Kernel Data..."
            if ($using:IncludeDumps -eq $true) {
                $NodeMinidumpsPath = Invoke-Command -ComputerName $using:NodeName -ConfigurationName $using:SessionConfigurationName {
                    (Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\CrashControl').MinidumpDir
                } -ErrorAction SilentlyContinue
                $NodeLiveKernelReportsPath = Invoke-Command -ComputerName $using:NodeName -ConfigurationName $using:SessionConfigurationName {
                    (Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\CrashControl\LiveKernelReports').LiveKernelReportsPath
                } -ErrorAction SilentlyContinue
                try {
                    if ($NodeMinidumpsPath) { $RPath = (Get-AdminSharePathFromLocal $using:NodeName "$NodeMinidumpsPath\*.dmp") }
                    else { $RPath = (Get-AdminSharePathFromLocal $using:NodeName "$NodeSystemRootPath\Minidump\*.dmp") }
                    $DmpFiles = Get-ChildItem -Path $RPath -Recurse -ErrorAction SilentlyContinue
                } catch { $DmpFiles = ""; Show-Warning "Unable to get minidump files for node $using:NodeName" }
                $DmpFiles |% { try { Copy-Item $_.FullName $LocalNodeDir } catch { Show-Warning("Could not copy minidump file $_.FullName") } }
                try {
                    if ($NodeLiveKernelReportsPath) { $RPath = (Get-AdminSharePathFromLocal $using:NodeName "$NodeLiveKernelReportsPath\*.dmp") }
                    else { $RPath = (Get-AdminSharePathFromLocal $using:NodeName "$NodeSystemRootPath\LiveKernelReports\*.dmp") }
                    $DmpFiles = Get-ChildItem -Path $RPath -Recurse -ErrorAction SilentlyContinue
                } catch { $DmpFiles = ""; Show-Warning "Unable to get LiveKernelReports files for node $using:NodeName" }
                $DmpFiles |% { try { Copy-Item $_.FullName $LocalNodeDir } catch { Show-Warning "Could not copy LiveKernelReports file $($_.FullName)" } }
            }

            Show-Update "Gathering Cluster Reports..."
            try {
                $RPath = (Get-AdminSharePathFromLocal $using:NodeName "$NodeSystemRootPath\Cluster\Reports\*.*")
                $RepFiles = Get-ChildItem -Path $RPath -Recurse -ErrorAction SilentlyContinue | Sort LastWriteTime
            } catch { $RepFiles = ""; Show-Warning "Unable to get reports for node $using:NodeName" }

            if (test-path "C:\Observability\OEMDiagnostics") {
                try {
                    $ASFiles=@()
                    $ASpath = (Get-AdminSharePathFromLocal $using:NodeName "$NodeSystemDrivePath\Observability\OEMDiagnostics")
                    $FWpath = (Get-AdminSharePathFromLocal $using:NodeName "$NodeSystemDrivePath\dell\logs\lcm")
                    $ASFiles += Get-ChildItem -Path $ASPath -Recurse -ErrorAction SilentlyContinue | Sort LastWriteTime | Select -Last 10
                    $ASFiles += Get-ChildItem -Path $FWPath -Recurse -ErrorAction SilentlyContinue | Sort LastWriteTime | Select -Last 10
                } catch { $ASFiles = ""; Show-Warning "No zipped OEMDiagnostics or FW Files available for $($using:NodeName)" }
            }

            $LocalReportDir = Join-Path $LocalNodeDir "ClusterReports"
            md $LocalReportDir -ErrorAction SilentlyContinue | Out-Null
            md $LocalDiagsDir -ErrorAction SilentlyContinue | Out-Null

            Do {
                Start-Sleep -Seconds 1
                foreach ($job in ($nodejobs | Where-Object { $_.State -eq 'Completed' -and -not $_.JobStatus })) {
                    $LocalFile = $job.Name
                    $output = Receive-Job $job
                    $output | Format-Table -AutoSize | Out-File -Width 9999 -Encoding ascii -FilePath "$LocalFile.txt"
                    $output | Export-Clixml -Path "$LocalFile.xml"
                    $job | Add-Member -MemberType NoteProperty -Name "JobStatus" -Value "JOBDONE" -Force
                    $job.Dispose()
                }
                $nodejobs | Format-List * | Out-File -FilePath (Join-Path $LocalNodeDir "GetNodeJobsStatus.txt")
            } while ($nodejobs.State -contains 'Running')

            $FailedJobs = @()
            foreach ($job in ($nodejobs | Where-Object { $_.State -ne 'Completed' })) { $FailedJobs += $job }
            $FailedJobs | Format-List * | Out-File -FilePath (Join-Path $LocalNodeDir "GetNodeJobsFailed.txt")
            $nodejobs | Remove-Job -Force

            $RepFiles |% {
                if (($_.Name -notlike "Cluster.log") -and ($_.Name -notlike "ClusterHealth.log")) {
                    try { Copy-Item $_.FullName $LocalReportDir } catch { Show-Warning "Could not copy report file $($_.FullName)" }
                }
            }
            if ($ASFiles.count -gt 0) {
                $ASFiles |% {
                    try { Copy-Item $_.FullName $LocalDiagsDir } catch { Show-Warning "Could not copy AS or FW Files file $($_.FullName)" }
                }
            }

            While (!(Test-Path $LocalFileMsInfo -ErrorAction SilentlyContinue) -and $msinfo.HasExited -ne $True) {Sleep -Milliseconds 100}
            While ($msinfo.HasExited -ne $True -and (Get-Item $LocalFileMsInfo -ErrorAction SilentlyContinue).LastWriteTime -ge (Get-Date).AddMinutes(-30)) {Sleep -Milliseconds 100}
        }
    }
    #endregion

    Show-Update "Starting export diagnostic log and live dump ..."
    $JobCopyOut += Invoke-SddcCommonCommand -ArgumentList $IncludeLiveDump,$IncludeStorDiag -ClusterNodes $AccessNode -SessionConfigurationName $SessionConfigurationName -InitBlock $CommonFunc -JobName StorageDiagnosticInfoAndLiveDump {
        Param($IncludeLiveDump,$IncludeStorDiag)
        $Node = $env:COMPUTERNAME
        $NodePath = $env:Temp
        $destinationPath = Join-Path -Path $NodePath -ChildPath 'StorageDiagnosticDump'
        if (Test-Path -Path $destinationPath) { Remove-Item -Path $destinationPath -Recurse -Force }
        $clusterSubsystem = (Get-StorageSubSystem |? Model -eq 'Clustered Windows Storage').FriendlyName
        if ($IncludeLiveDump) {
            Get-StorageDiagnosticInfo -StorageSubSystemFriendlyName $clusterSubsystem -IncludeLiveDump -DestinationPath $destinationPath
            Write-Output (Get-AdminSharePathFromLocal $Node $destinationPath)
        } elseif ($IncludeStorDiag) {
            Get-StorageDiagnosticInfo -StorageSubSystemFriendlyName $clusterSubsystem -DestinationPath $destinationPath
            Write-Output (Get-AdminSharePathFromLocal $Node $destinationPath)
        }
    }

    Show-Update "Starting export of events ..."
    $JobCopyOut += Invoke-SddcCommonCommand -ArgumentList $HoursOfEvents -ClusterNodes $($ClusterNodes).Name -SessionConfigurationName $SessionConfigurationName -InitBlock $CommonFunc -JobName Events {
        Param([int] $Hours)
        $Node = $env:COMPUTERNAME
        $NodePath = $env:Temp
        Get-SddcCapturedEvents $NodePath $Hours |% { Write-Output (Get-AdminSharePathFromLocal $Node $_) }
        Write-Output (Get-AdminSharePathFromLocal $Node (Join-Path $NodePath "LocaleMetaData"))
    }

    if ($IncludeAssociations -and $ClusterName.Length) {
        $SNVJob = Start-Job -Name 'StorageNodePhysicalDiskView' -ArgumentList $ClusterName {
            param ($ClusterName)
            $clusterCimSession = New-CimSession -ComputerName $ClusterName
            $snvInstances = Get-CimInstance -Namespace root\Microsoft\Windows\Storage -ClassName MSFT_StorageNodeToPhysicalDisk -CimSession $clusterCimSession
            $allPhysicalDisks = Get-PhysicalDisk -CimSession $clusterCimSession
            $SNV = @()
            Foreach ($phyDisk in $snvInstances) {
                $SNVObject = New-Object -TypeName System.Object
                $pdIndex = $phyDisk.PhysicalDiskObjectId.IndexOf("PD:")
                $pdLength = $phyDisk.PhysicalDiskObjectId.Length
                $pdID = $phyDisk.PhysicalDiskObjectId.Substring($pdIndex+3, $pdLength-($pdIndex+4))
                $PDUID = ($allPhysicalDisks |? ObjectID -Match $pdID).UniqueID
                $pd = $allPhysicalDisks |? UniqueID -eq $PDUID
                $nodeIndex = $phyDisk.StorageNodeObjectId.IndexOf("SN:")
                $nodeLength = $phyDisk.StorageNodeObjectId.Length
                $storageNodeName = $phyDisk.StorageNodeObjectId.Substring($nodeIndex+3, $nodeLength-($nodeIndex+4))
                $poolName = ($pd | Get-StoragePool -CimSession $clusterCimSession -ErrorAction SilentlyContinue |? IsPrimordial -eq $false).FriendlyName
                if (-not $poolName) { continue }
                $SNVObject | Add-Member -Type NoteProperty -Name PhysicalDiskUID -Value $PDUID
                $SNVObject | Add-Member -Type NoteProperty -Name StorageNode -Value $storageNodeName
                $SNVObject | Add-Member -Type NoteProperty -Name StoragePool -Value $poolName
                $SNVObject | Add-Member -Type NoteProperty -Name MPIOPolicy -Value $phyDisk.LoadBalancePolicy
                $SNVObject | Add-Member -Type NoteProperty -Name MPIOState -Value $phyDisk.IsMPIOEnabled
                $SNVObject | Add-Member -Type NoteProperty -Name StorageEnclosure -Value $pd.PhysicalLocation
                $SNVObject | Add-Member -Type NoteProperty -Name PathID -Value $phyDisk.PathID
                $SNVObject | Add-Member -Type NoteProperty -Name PathState -Value $phyDisk.PathState
                $SNV += $SNVObject
            }
            Write-Output $SNV
        }
        $AssocJob = Start-Job -Name 'StorageComponentAssociations' -ArgumentList $AccessNode,$ClusterName {
            param($AccessNode,$ClusterName)
            $SmbShares = Get-SmbShare -CimSession $AccessNode
            $Associations = Get-VirtualDisk -CimSession $AccessNode |% {
                $o = $_ | Select-Object FriendlyName, OperationalStatus, HealthStatus, CSVName, CSVStatus, CSVNode, CSVPath, CSVVolume, ShareName, SharePath, VolumeID, PoolName, PoolOpStatus, PoolHealthStatus, VDResiliency, VDCopies, VDColumns, VDEAware
                $AssocCSV = $_ | Get-ClusterSharedVolume -Cluster $ClusterName
                if ($AssocCSV) {
                    $o.CSVName = $AssocCSV.Name; $o.CSVStatus = $AssocCSV.State; $o.CSVNode = $AssocCSV.OwnerNode.Name
                    $o.CSVPath = $AssocCSV.SharedVolumeInfo.FriendlyVolumeName
                    if ($o.CSVPath.Length -ne 0) { $o.CSVVolume = $o.CSVPath.Split("\")[2] }
                    $AssocLike = $o.CSVPath+"\*"
                    $AssocShares = $SmbShares |? Path -like $AssocLike
                    $AssocShare = $AssocShares | Select-Object -First 1
                    if ($AssocShare) {
                        $o.ShareName = $AssocShare.Name; $o.SharePath = $AssocShare.Path; $o.VolumeID = $AssocShare.Volume
                        if ($AssocShares.Count -gt 1) { $o.ShareName += "*" }
                    }
                }
                Write-Output $o
            }
            $AssocPool = Get-StoragePool -CimSession $AccessNode -ErrorAction SilentlyContinue
            $AssocPool |% {
                $AssocPName = $_.FriendlyName; $AssocPOpStatus = $_.OperationalStatus; $AssocPHStatus = $_.HealthStatus
                Get-StoragePool -CimSession $AccessNode -FriendlyName $AssocPName | Get-VirtualDisk -CimSession $AccessNode |% {
                    $AssocVD = $_
                    $Associations |% {
                        if ($_.FriendlyName -eq $AssocVD.FriendlyName) {
                            $_.PoolName = $AssocPName; $_.PoolOpStatus = $AssocPOpStatus; $_.PoolHealthStatus = $AssocPHStatus
                            $_.VDResiliency = $AssocVD.ResiliencySettingName; $_.VDCopies = $AssocVD.NumberofDataCopies
                            $_.VDColumns = $AssocVD.NumberofColumns; $_.VDEAware = $AssocVD.IsEnclosureAware
                        }
                    }
                }
            }
            Write-Output $Associations
        }
    }

    # SMB share health/status
    Show-Update "SMB Shares"
    try { $SmbShares = Get-SmbShare -CimSession $AccessNode }
    catch { Show-Error("Unable to get SMB Shares. `nError="+$_.Exception.Message) }
    $ShareStatus = $SmbShares |? ContinuouslyAvailable | Select-Object ScopeName, Name, SharePath, Health
    $Count1 = 0; $Total1 = NCount($ShareStatus)
    if ($Total1 -gt 0) {
        $ShareStatus |% {
            $Progress = $Count1 / $Total1 * 100; $Count1++
            Write-Progress -Activity "Testing file share access" -PercentComplete $Progress
            if ($ClusterDomain -ne "") { $_.SharePath = "\\" + $_.ScopeName + "." + $ClusterDomain + "\" + $_.Name }
            else { $_.SharePath = "\\" + $_.ScopeName + "\" + $_.Name }
            try {
                if (Test-Path -Path $_.SharePath -ErrorAction SilentlyContinue) { $_.Health = "Accessible" }
                else { $_.Health = "Inaccessible" }
            } catch { $_.Health = "Accessible: "+$_.Exception.Message }
        }
        Write-Progress -Activity "Testing file share access" -Completed
    }
    $ShareStatus | Export-Clixml ($Path + "ShareStatus.XML")

    Show-Update "SMB Share Open Files"
    try { $o = Get-SmbOpenFile -CimSession $AccessNode; $o | Export-Clixml ($Path + "GetSmbOpenFile.XML") }
    catch { Show-Error("Unable to get Open Files. `nError="+$_.Exception.Message) }

    Show-Update "SMB Share Witness"
    try { $o = Get-SmbWitnessClient -CimSession $AccessNode; $o | Export-Clixml ($Path + "GetSmbWitness.XML") }
    catch { Show-Error("Unable to get SMB Witness. `nError="+$_.Exception.Message) }

    Show-Update "Clustered Subsystem"
    try { $Subsystem = Get-StorageSubsystem Cluster* -CimSession $AccessNode; $Subsystem | Export-Clixml ($Path + "GetStorageSubsystem.XML") }
    catch { Show-Warning("Unable to get Clustered Subsystem.`nError="+$_.Exception.Message) }

    if ($Subsystem.HealthStatus -notlike "Healthy" -and $ClusterName.Length) {
        Show-Update "Triage for Clustered Subsystem (HealthStatus = $($Subsystem.HealthStatus))"
        try {
            $cmdlet = Get-Command Get-HealthFault -ErrorAction SilentlyContinue
            if ($null -ne $cmdlet -and $cmdlet.Source -eq 'FailoverClusters') {
                Get-HealthFault -CimSession $AccessNode | Export-Clixml (Join-Path $Path "HeathFault.XML")
            } else {
                $Subsystem | Debug-StorageSubsystem -CimSession $AccessNode | Export-Clixml (Join-Path $Path "DebugStorageSubsystem.XML")
            }
        } catch { Show-Error "Unable to get Get-HealthFault or Debug-StorageSubsystem.`nError=" $_ }
    }

    Show-Update "Volumes & Virtual Disks"
    try { $Volumes = Get-Volume -CimSession $AccessNode -StorageSubSystem $Subsystem; $Volumes | Export-Clixml ($Path + "GetVolume.XML") }
    catch { Show-Error("Unable to get Volumes. `nError="+$_.Exception.Message) }

    try {
        $VirtualDisk = Get-VirtualDisk -CimSession $AccessNode -StorageSubSystem $Subsystem
        $VirtualDisk = $VirtualDisk | Select *,@{L="FileSystem";E={($_ | Get-Disk | Get-Partition | Get-Volume).FileSystemType}}
        $VirtualDisk | Export-Clixml ($Path + "GetVirtualDisk.XML")
    } catch { Show-Warning("Unable to get Virtual Disks.`nError="+$_.Exception.Message) }

    if ($DedupEnabled) {
        Show-Update "Dedup Volume Status"
        try { $DedupVolumes = Invoke-Command -ComputerName $AccessNode -ConfigurationName $SessionConfigurationName { Get-DedupStatus }; $DedupVolumes | Export-Clixml ($Path + "GetDedupVolume.XML") }
        catch { Show-Error("Unable to get Dedup Volumes.`nError="+$_.Exception.Message) }
        $DedupTotal = NCount($DedupVolumes); $DedupHealthy = NCount($DedupVolumes |? LastOptimizationResult -eq 0)
    } else { $DedupVolumes = @(); $DedupTotal = 0; $DedupHealthy = 0 }

    Show-Update "Storage Pool & Tiers"
    try { Get-StorageNode -CimSession $AccessNode | Export-Clixml ($Path + "GetStorageNode.XML") }
    catch { Show-Warning("Unable to get Storage Nodes. `nError="+$_.Exception.Message) }
    try { Get-StorageTier -CimSession $AccessNode | Export-Clixml ($Path + "GetStorageTier.XML") }
    catch { Show-Warning("Unable to get Storage Tiers. `nError="+$_.Exception.Message) }
    try { $StoragePools = @(Get-StoragePool -IsPrimordial $False -CimSession $AccessNode -StorageSubSystem $Subsystem -ErrorAction SilentlyContinue); $StoragePools | Export-Clixml ($Path + "GetStoragePool.XML") }
    catch { Show-Error("Unable to get Storage Pools. `nError="+$_.Exception.Message) }

    Show-Update "Storage Jobs"
    try { icm $AccessNode -ConfigurationName $SessionConfigurationName { Get-StorageJob } | Export-Clixml ($Path + "GetStorageJob.XML") }
    catch { Show-Warning("Unable to get Storage Jobs. `nError="+$_.Exception.Message) }

    Show-Update "Clustered PhysicalDisks and SNV"
    try { $PhysicalDisks = Get-PhysicalDisk -CimSession $AccessNode -StorageSubSystem $Subsystem; $PhysicalDisks | Export-Clixml ($Path + "GetPhysicalDisk.XML") }
    catch { Show-Error("Unable to get Physical Disks. `nError="+$_.Exception.Message) }
    try { Get-PhysicalDisk -CimSession $AccessNode -StorageSubSystem $Subsystem | Get-PhysicalDiskSNV -CimSession $AccessNode | Export-Clixml ($Path + "GetPhysicalDiskSNV.XML") }
    catch { Show-Error("Unable to get Physical Disk Storage Node View. `nError="+$_.Exception.Message) }

    if ($IncludeReliabilityCounters -eq $true) {
        Show-Update "Storage Reliability Counters"
        try { $PhysicalDisks | Get-StorageReliabilityCounter -CimSession $AccessNode | Export-Clixml ($Path + "GetReliabilityCounter.XML") }
        catch { Show-Error("Unable to get Storage Reliability Counters. `nError="+$_.Exception.Message) }
    }

    Show-Update "Storage Enclosures"
    try { Get-StorageEnclosure -CimSession $AccessNode -StorageSubSystem $Subsystem | Export-Clixml ($Path + "GetStorageEnclosure.XML") }
    catch { Show-Error("Unable to get Enclosures. `nError="+$_.Exception.Message) }

    if ($S2DEnabled) {
        Show-Update "Pooled Disks"
        try { if ($StoragePools.Count -eq 1) { $StoragePools | Get-PhysicalDisk -CimSession $AccessNode | Export-Clixml (Join-Path $Path "GetPhysicalDisk_Pool.xml") } }
        catch { Show-Error "Not able to query pooled disks" $_ }

        Show-Update "Storage Scale Units"
        try { $Subsystem | Get-StorageFaultDomain -CimSession $AccessNode -Type StorageScaleUnit | Export-Clixml (Join-Path $Path "GetStorageFaultDomain_SSU.xml") }
        catch { Show-Error "Not able to query Storage Scale Units" $_ }

        Show-Update "S2D Connectivity"
        try {
            $JobStatic += $ClusterNodes |% {
                $node = $_.Name
                start-job -Name "S2D Connectivity: $node" {
                    Get-CimInstance -Namespace root\wmi -ClassName ClusPortDeviceInformation -ComputerName $using:node | Export-Clixml (Join-Path (Join-Path $using:Path "Node_$using:node") "ClusPort.xml")
                    Get-CimInstance -Namespace root\wmi -ClassName ClusBfltDeviceInformation -ComputerName $using:node | Export-Clixml (Join-Path (Join-Path $using:Path "Node_$using:node") "ClusBflt.xml")
                }
            }
        } catch { Show-Warning "Gathering S2D connectivity failed" }

        Show-Update "Cluster Performance History"
        try {
            $JobStatic += start-job -Name "Cluster Performance History" {
                $Cluster = Get-cluster
                $ClusterNodes = Get-ClusterNode -Cluster $Cluster -ErrorAction SilentlyContinue
                $Output = $ClusterNodes | ForEach-Object {
                    $_ | Get-ClusterPerf -ClusterNodeSeriesName "ClusterNode.Cpu.Usage" -TimeFrame "LastWeek" -ErrorAction SilentlyContinue
                }
                $Output | Sort-Object ClusterNode | Export-Clixml ($using:Path + "CPUIseeyou.xml")

                try {
                    $o = Invoke-Command $ClusterNodes.Name {
                        Function Format-Latency { Param ($RawValue); $i = 0; $Labels = ("s","ms","$([char]956)s","ns"); Do { $RawValue *= 1000; $i++ } While ($RawValue -Lt 1); [String][Math]::Round($RawValue,2)+" "+$Labels[$i] }
                        Function Format-StandardDeviation { Param ($RawValue); If ($RawValue -Gt 0){$Sign="+"}Else{$Sign="-"}; $Sign+[String][Math]::Round([Math]::Abs($RawValue),2) }
                        $HDD = Get-StorageNode |?{$ENV:COMPUTERNAME -imatch ($_.name -split '\.')[0]} | Get-PhysicalDisk -PhysicallyConnected
                        $Output = $HDD | ForEach-Object {
                            $Iops = $_ | Get-ClusterPerf -PhysicalDiskSeriesName "PhysicalDisk.Iops.Total" -TimeFrame "LastWeek"
                            $AvgIops = ($Iops | Measure-Object -Property Value -Average).Average
                            If ($AvgIops -Gt 0) {
                                $Latency = $_ | Get-ClusterPerf -PhysicalDiskSeriesName "PhysicalDisk.Latency.Average" -TimeFrame "LastWeek"
                                $AvgLatency = ($Latency | Measure-Object -Property Value -Average).Average
                                [PsCustomObject]@{ "FriendlyName"=$_.FriendlyName;"SerialNumber"=$_.SerialNumber;"MediaType"=$_.MediaType;"AvgLatencyPopulation"=$null;"AvgLatencyThisHDD"=Format-Latency $AvgLatency;"RawAvgLatencyThisHDD"=$AvgLatency;"Deviation"=$null;"RawDeviation"=$null }
                            }
                        }
                        If ($Output.Length -Ge 3) {
                            $u = ($Output | Measure-Object -Property RawAvgLatencyThisHDD -Average).Average
                            $d = $Output | ForEach-Object { ($_.RawAvgLatencyThisHDD - $u) * ($_.RawAvgLatencyThisHDD - $u) }
                            $o2 = [Math]::Sqrt(($d | Measure-Object -Sum).Sum / $Output.Length)
                            $Output | ForEach-Object { $Deviation = ($_.RawAvgLatencyThisHDD - $u)/$o2; $_.AvgLatencyPopulation = Format-Latency $u; $_.Deviation = Format-StandardDeviation $Deviation; $_.RawDeviation = $Deviation }
                        }
                        $Output
                    }
                    $o | Sort-Object PsComputerName | Export-Clixml ($using:Path + "latencyoutlier.xml")
                } catch {}

                try {
                    $o = Invoke-Command $ClusterNodes.Name {
                        Function Format-Iops { Param ($RawValue); $i=0;$Labels=(" ","K","M","B","T"); Do{if($RawValue -Gt 1000){$RawValue/=1000;$i++}}While($RawValue -Gt 1000); [String][Math]::Round($RawValue)+" "+$Labels[$i] }
                        Get-VM | ForEach-Object {
                            $IopsTotal = $_ | Get-ClusterPerf -VMSeriesName "VHD.Iops.Total"
                            $IopsRead  = $_ | Get-ClusterPerf -VMSeriesName "VHD.Iops.Read"
                            $IopsWrite = $_ | Get-ClusterPerf -VMSeriesName "VHD.Iops.Write"
                            [PsCustomObject]@{ "VM"=$_.Name;"IopsTotal"=Format-Iops $IopsTotal.Value;"IopsRead"=Format-Iops $IopsRead.Value;"IopsWrite"=Format-Iops $IopsWrite.Value;"RawIopsTotal"=$IopsTotal.Value }
                        }
                    }
                    $o | Sort-Object RawIopsTotal -Descending | Select-Object -First 10 | Export-Clixml ($using:Path + "Noisyneighbor.xml")
                } catch {}

                try {
                    $o = Invoke-Command $ClusterNodes.Name {
                        Function Format-BitsPerSec { Param ($RawValue); $i=0;$Labels=("bps","kbps","Mbps","Gbps","Tbps","Pbps"); Do{$RawValue/=1000;$i++}While($RawValue -Gt 1000); [String][Math]::Round($RawValue)+" "+$Labels[$i] }
                        Get-NetAdapter | ForEach-Object {
                            $Inbound  = $_ | Get-ClusterPerf -NetAdapterSeriesName "NetAdapter.Bandwidth.Inbound" -TimeFrame "LastDay"
                            $Outbound = $_ | Get-ClusterPerf -NetAdapterSeriesName "NetAdapter.Bandwidth.Outbound" -TimeFrame "LastDay"
                            If ($Inbound -Or $Outbound) {
                                $MeasureInbound  = $Inbound  | Measure-Object -Property Value -Maximum
                                $MeasureOutbound = $Outbound | Measure-Object -Property Value -Maximum
                                $Saturated = $False
                                If (($MeasureInbound.Maximum -Gt (0.90 * $_.Speed)) -Or ($MeasureOutbound.Maximum -Gt (0.90 * $_.Speed))) { $Saturated = $True }
                                [PsCustomObject]@{ "NetAdapter"=$_.InterfaceDescription;"LinkSpeed"=$_.LinkSpeed;"MaxInbound"=Format-BitsPerSec $MeasureInbound.Maximum;"MaxOutbound"=Format-BitsPerSec $MeasureOutbound.Maximum;"Saturated"=$Saturated }
                            }
                        }
                    }
                    $o | Sort-Object PsComputerName, InterfaceDescription | Export-Clixml ($using:Path + "25gigisthenew10gig.xml")
                } catch {}

                try {
                    Get-Volume | Where-Object FileSystem -Like "*CSV*" |
                        %{$_ | Get-ClusterPerf -VolumeSeriesName "Volume.Size.Available" -TimeFrame "LastYear" | Sort-Object Time | Select-Object -Last 14} |
                        Sort-Object ClusterNode | Export-Clixml ($using:Path + "trendyagain.xml")
                } catch {}

                try {
                    $Output = Invoke-Command (Get-ClusterNode).Name {
                        Function Format-Bytes { Param ($RawValue); $i=0;$Labels=("B","KB","MB","GB","TB","PB","EB","ZB","YB"); Do{if($RawValue -Gt 1024){$RawValue/=1024;$i++}}While($RawValue -Gt 1024); [String][Math]::Round($RawValue)+" "+$Labels[$i] }
                        Get-VM | ForEach-Object {
                            $Data = $_ | Get-ClusterPerf -VMSeriesName "VM.Memory.Assigned" -TimeFrame "LastMonth"
                            If ($Data) {
                                $AvgMemoryUsage = ($Data | Measure-Object -Property Value -Average).Average
                                [PsCustomObject]@{ "VM"=$_.Name;"AvgMemoryUsage"=Format-Bytes $AvgMemoryUsage;"RawAvgMemoryUsage"=$AvgMemoryUsage }
                            }
                        }
                    }
                    $Output | Sort-Object RawAvgMemoryUsage -Descending | Select-Object -First 10 | Export-Clixml ($using:Path + "Memoryhog.xml")
                } catch {}
            }
        } catch { Show-Warning "Gathering Cluster Performance History failed" }
    }

    Show-Update "AzureStack HCI info"
    try {
        If ((Get-WmiObject -Class Win32_OperatingSystem).Caption -imatch "HCI"){
            Get-AzureStackHCI | Export-Clixml ($Path + "GetAzureStackHCI.xml")
            Get-AzureStackHCIArcIntegration | Export-Clixml ($Path + "AzureStackHCIArcIntegration.xml")
        }
    } catch { Show-Warning("Unable to get AzureStack HCI info. `nError="+$_.Exception.Message) }

    Show-Update "Start gather of Cluster Performance information..."

    # Remote copyout jobs
    if ($JobCopyOut.Count -or $JobCopyOutNoDelete.Count) {
        Show-Update "Completing jobs with remote copyout ..." -ForegroundColor Green
        Show-WaitChildJob ($JobCopyOut + $JobCopyOutNoDelete) 120
        Show-Update "Starting remote copyout ..."
        $JobCopy = @()
        if ($JobCopyOut.Count) { $JobCopy += Start-CopyJob $Path -Delete $JobCopyOut }
        if ($JobCopyOutNoDelete.Count) { $JobCopy += Start-CopyJob $Path $JobCopyOutNoDelete }
        Show-WaitChildJob $JobCopy 30
        Receive-Job $JobCopy
        Remove-Job ($JobCopyOut + $JobCopyOutNoDelete)
        Remove-Job $JobCopy
        if (Get-Member -InputObject $JobCopyOut ActiveSessions) { Remove-PSSession -Id $JobCopyOut.ActiveSessions }
    }
    Show-Update "All remote copyout complete" -ForegroundColor Green

    # Static jobs
    Show-Update "Completing background gathers ..." -ForegroundColor Green
    Show-Update "Start monitoring $($PerfSamples)s" -ForegroundColor Green

    $PerfProc = Start-Process -WindowStyle Hidden -FilePath "powershell.exe" -ArgumentList @("-Command", """& {Get-Counter -Counter (Get-Counter -ListSet 'Cluster Storage*','Cluster CSV*','Storage Spaces*','Storage Replica*','Refs','Cluster Disk Counters','PhysicalDisk','RDMA*','Mellanox WinOF-2 Port Traffic*','Mellanox WinOF-2 Congestion Control*','Mellanox WinOF-2 Diagnostics Ext 1*','Marvell*','Hyper-V Hypervisor Virtual Processor','Hyper-V Hypervisor Logical Processor','Hyper-V Hypervisor Root Virtual Processor' -ComputerName (Get-ClusterNode).Name -ErrorAction SilentlyContinue).paths -SampleInterval 1 -MaxSamples $PerfSamples -ErrorAction Ignore -WarningAction Ignore | Export-counter -Path ('$Path' + '\GetCounters.blg') -Force -FileFormat BLG}""") -Passthru

    Show-WaitChildJob $JobStatic 30
    $JobStatic |% { if ($_.Name -ne "Cluster Performance History" -and $_.Name -notlike "ClusterLogs*") { $o=Receive-Job $_; If ($o) {Write-Host "Job $($_.Name) Output:";$o} } }
    Remove-Job $JobStatic

    # Collect Send-DiagnosticData
    IF($ClusterNodes -eq $Null){ $ClusterNodes = Get-ClusterNode | Select-Object -ExpandProperty Name }
    if (Get-Service 'HciSvc' -ErrorAction SilentlyContinue) {
        Show-Update "Copy-DirContentFromNode -Nodes $ClusterNodes -PathOnNode 'C:\SendDiags' -SearchFilter 'DiagLogs-*' -LocalRoot $($env:userprofile + "\HealthTest\")"
        Copy-DirContentFromNode -Nodes $ClusterNodes -PathOnNode 'C:\SendDiags' -SearchFilter 'DiagLogs-*' -LocalDest $($env:userprofile + "\HealthTest\")
    }

    Show-Update "Copying cluster logs."
    Foreach ($NodeName in ((Get-ClusterNode).Name)) {
        $NodeSystemRootPath = Invoke-Command -ComputerName $NodeName { $env:SystemRoot }
        try {
            $RPath = (Get-AdminSharePathFromLocal $NodeName "$NodeSystemRootPath\Cluster\Reports\cluster*.log")
            $RepFiles = Get-ChildItem -Path $RPath -Recurse -ErrorAction SilentlyContinue | Sort LastWriteTime
        } catch { $RepFiles = ""; Show-Warning "Unable to get reports for node $NodeName" }
        $RepFiles |% {
            $DestPath=(Join-Path $Path $NodeName)+"_$($_.Name)"
            If (($_.Name -eq "Cluster.log" -or $_.Name -eq "ClusterHealth.log") -and -not (Test-Path $DestPath)) {
                try { Copy-Item $_.FullName $DestPath } catch { Show-Warning "Could not copy report file $($_.FullName)" }
            }
        }
    }

    if (Get-Member -InputObject $JobStatic ActiveSessions) { Remove-PSSession -Id $JobStatic.ActiveSessions }

    Remove-Variable JobCopyOut
    Remove-Variable JobStatic

    # Phase 2 Prep
    Show-Update "<<< Phase 2 - Pool, Physical Disk and Volume Details >>>" -ForegroundColor Cyan

    if ($IncludeAssociations) {
        if ($Read) {
            $Associations = Import-ClixmlIf ($Path + "GetAssociations.XML")
            $SNVView = Import-ClixmlIf ($Path + "GetStorageNodeView.XML")
        } else {
            "`nCollecting device associations..."
            try {
                $Associations = $AssocJob | Wait-Job | Receive-Job
                $AssocJob | Remove-Job
                if ($null -eq $Associations) { Show-Warning "Unable to get object associations" }
                $Associations | Export-Clixml ($Path + "GetAssociations.XML")
                "`nCollecting storage view associations..."
                $SNVView = $SNVJob | Wait-Job | Receive-Job
                $SNVJob | Remove-Job
                if ($null -eq $SNVView) { Show-Warning "Unable to get nodes storage view associations" }
                $SNVView | Export-Clixml ($Path + "GetStorageNodeView.XML")
            } catch { Show-Warning "Not able to query associations.." }
        }
    }

    # Phase 2 - Health Report (optional)
    if ($IncludeHealthReport) {
        "`n[Health Report]"
        "`nVolumes with status, total size and available size, sorted by Available Size"
        "Notes: Sizes shown in gigabytes (GB). * means multiple shares on that volume"
        $Volumes |? FileSystem -eq CSVFS | Sort-Object SizeRemaining | Format-Table -AutoSize `
            @{Expression={$poolName = VolumeToPool($_.Path); "[$(PoolOperationalStatus($_.Path))/$(PoolHealthStatus($_.Path))] " + $poolName};Label="[OpStatus/Health] Pool"},
            @{Expression={(PoolHealthyPDs(VolumeToPool($_.Path)))};Label="HealthyPhysicalDisks"; Align="Center"},
            @{Expression={$vd = VolumeToVD($_.Path); "[$(VDOperationalStatus($_.Path))/$(VDHealthStatus($_.Path))] "+$vd};Label="[OpStatus/Health] VirtualDisk"},
            @{Expression={$csvVolume = VolumeToCSV($_.Path); "[" + $_.HealthStatus + "] " + $csvVolume};Label="[Health] CSV Volume"},
            @{Expression={$csvName = VolumeToCSVName($_.Path); $csvStatus = CSVStatus($_.Path); " [$csvStatus] " + $csvName};Label="[Status] CSV Name"},
            @{Expression={CSVToNode(VolumeToCSV($_.Path))};Label="Volume Owner"},
            @{Expression={VolumeToShare($_.Path)};Label="Share Name"},
            @{Expression={$VolResiliency = VolumeToResiliency($_.Path); $volColumns = VolumeToColumns($_.Path); "$VolResiliency,$volColumns" +"Col" };Label="Volume Configuration"},
            @{Expression={"{0:N2}" -f ($_.Size/1GB)};Label="Total Size";Width=11;Align="Right"},
            @{Expression={"{0:N2}" -f ($_.SizeRemaining/$_.Size*100)};Label="Avail%";Width=11;Align="Right"}

        if ($DedupEnabled -and ($DedupTotal -gt 0)) {
            "Dedup Volumes"
            $DedupVolumes | Sort-Object SavingsRate -Descending | Format-Table -AutoSize `
                @{Expression={$poolName = VolumeToPool($_.VolumeId); "[$(PoolOperationalStatus($_.VolumeId))/$(PoolHealthStatus($_.VolumeId))] " + $poolName};Label="[OpStatus/Health] Pool"},
                @{Expression={(PoolHealthyPDs(VolumeToPool($_.VolumeId)))};Label="HealthyPhysicalDisks"; Align="Center"},
                @{Expression={$vd = VolumeToVD($_.VolumeId); "[$(VDOperationalStatus($_.VolumeId))/$(VDHealthStatus($_.VolumeId))] "+$vd};Label="[OpStatus/Health] VirtualDisk"},
                @{Expression={VolumeToCSV($_.VolumeId)};Label="Volume "},
                @{Expression={VolumeToShare($_.VolumeId)};Label="Share"},
                @{Expression={"{0:N2}" -f ($_.Capacity/1GB)};Label="Capacity";Width=11;Align="Left"},
                @{Expression={"{0:N2}" -f ($_.UnoptimizedSize/1GB)};Label="Before";Width=11;Align="Right"},
                @{Expression={"{0:N2}" -f ($_.UsedSpace/1GB)};Label="After";Width=11;Align="Right"},
                @{Expression={"{0:N2}" -f ($_.SavingsRate)};Label="Savings%";Width=11;Align="Right"},
                @{Expression={"{0:N2}" -f ($_.FreeSpace/1GB)};Label="Free";Width=11;Align="Right"},
                @{Expression={"{0:N2}" -f ($_.FreeSpace/$_.Capacity*100)};Label="Free%";Width=11;Align="Right"},
                @{Expression={"{0:N0}" -f ($_.InPolicyFilesCount)};Label="Files";Width=11;Align="Right"}
        }

        if ($SNVView) {
            "`n[Storage Node view]"
            $SNVView | sort StorageNode,StorageEnclosure | Format-Table -AutoSize `
                @{Expression={$_.StorageNode};Label="StorageNode";Align="Left"},
                @{Expression={$_.StoragePool};Label="StoragePool";Align="Left"},
                @{Expression={$_.MPIOPolicy};Label="MPIOPolicy";Align="Left"},
                @{Expression={$_.MPIOState};Label="MPIOState";Align="Left"},
                @{Expression={$_.PathID};Label="PathID";Align="Left"},
                @{Expression={$_.PathState};Label="PathState";Align="Left"},
                @{Expression={$_.PhysicalDiskUID};Label="PhysicalDiskUID";Align="Left"},
                @{Expression={$_.StorageEnclosure};Label="StorageEnclosureLocation";Align="Left"}
        }

        "`n[Capacity Report]"
        $PDStatus = $PhysicalDisks |? EnclosureNumber -ne $null | Sort-Object EnclosureNumber, MediaType, HealthStatus |
            Group-Object EnclosureNumber, MediaType, HealthStatus | Select-Object Count, TotalSize, Unalloc,
                @{Expression={$_.Name.Split(",")[0].Trim().TrimEnd()};Label="Enc"},
                @{Expression={$_.Name.Split(",")[1].Trim().TrimEnd()};Label="Media"},
                @{Expression={$_.Name.Split(",")[2].Trim().TrimEnd()};Label="Health"}
        $PDStatus |% {
            $Current = $_; $TotalSize = 0; $Unalloc = 0
            $PDCurrent = $PhysicalDisks |? { ($_.EnclosureNumber -eq $Current.Enc) -and ($_.MediaType -eq $Current.Media) -and ($_.HealthStatus -eq $Current.Health) }
            $PDCurrent |% { $Unalloc += $_.Size - $_.AllocatedSize; $TotalSize += $_.Size }
            $Current.Unalloc = $Unalloc; $Current.TotalSize = $TotalSize
        }
        $PDStatus | Format-Table -AutoSize Enc, Media, Health, Count,
            @{Expression={"{0:N2}" -f ($_.TotalSize/$_.Count/1GB)};Label="Avg Size";Width=11;Align="Right"},
            @{Expression={"{0:N2}" -f ($_.TotalSize/1GB)};Label="Total Size";Width=11;Align="Right"},
            @{Expression={"{0:N2}" -f ($_.Unalloc/1GB)};Label="Unallocated";Width=11;Align="Right"},
            @{Expression={"{0:N2}" -f ($_.Unalloc/$_.TotalSize*100)};Label="Unalloc %";Width=11;Align="Right"}

        "Pools with health, total size and unallocated space"
        $StoragePools | Sort-Object FriendlyName | Format-Table -AutoSize `
            @{Expression={$_.FriendlyName};Label="Name"},
            @{Expression={$_.HealthStatus};Label="Health"},
            @{Expression={"{0:N2}" -f ($_.Size/1GB)};Label="Total Size";Width=11;Align="Right"},
            @{Expression={"{0:N2}" -f (($_.Size-$_.AllocatedSize)/1GB)};Label="Unallocated";Width=11;Align="Right"},
            @{Expression={"{0:N2}" -f (($_.Size-$_.AllocatedSize)/$_.Size*100)};Label="Unalloc%";Width=11;Align="Right"}
    }

    # Phase 3 - Storage Performance
    Show-Update "<<< Phase 3 - Storage Performance >>>" -ForegroundColor Cyan

    if (-not $IncludePerformance) {
        "Performance was excluded by a parameter`n"
    } else {
        Show-Update "Waiting for performance counters to complete. Timeout in 30 minutes..."
        $xb=0
        Do {Write-Host -Nonewline ".";sleep 10;$xb++} While ($PerfProc.HasExited -ne $True -and $xb -lt 180)
        Write-Host ""
        If ($xb -lt 180) { Show-Update "Performance monitoring completed" }
        else { $PerfProc | kill; Show-Warning "Performance monitoring timed out" }

        if ($ProcessCounter) {
            "Collected $PerfSamples seconds of raw performance counters. Processing...`n"
            # ProcessCounter logic omitted (deprecated) - performance data is in the .blg file
        }
    }

    if ($S2DEnabled -ne $true) {
        try {
            if ((([System.Environment]::OSVersion.Version).Major) -ge 10) {
                Show-Update "Gathering Get-StorageDiagnosticInfo"
                $deleteStorageSubsystem = $false
                if (-not (Get-StorageSubsystem -FriendlyName Clustered*)) {
                    $storageProviderName = (Get-StorageProvider -CimSession $ClusterName |? Manufacturer -match 'Microsoft').Name
                    $null = Register-StorageSubsystem -ProviderName $storageProviderName -ComputerName $ClusterName -ErrorAction SilentlyContinue
                    $deleteStorageSubsystem = $true
                    $storagesubsystemToDelete = Get-StorageSubsystem -FriendlyName Clustered*
                }
                $destinationPath = Join-Path -Path $Path -ChildPath 'StorageDiagnosticInfo'
                if (Test-Path -Path $destinationPath) { Remove-Item -Path $destinationPath -Recurse -Force }
                $null = New-Item -Path $destinationPath -ItemType Directory
                $clusterSubsystem = (Get-StorageSubSystem |? Model -eq 'Clustered Windows Storage').FriendlyName
                Stop-StorageDiagnosticLog -StorageSubSystemFriendlyName $clusterSubsystem -ErrorAction SilentlyContinue
                if ($IncludeLiveDump) {
                    Get-StorageDiagnosticInfo -StorageSubSystemFriendlyName $clusterSubsystem -IncludeLiveDump -DestinationPath $destinationPath
                } else {
                    Get-StorageDiagnosticInfo -StorageSubSystemFriendlyName $clusterSubsystem -DestinationPath $destinationPath
                }
                if ($deleteStorageSubsystem) {
                    Unregister-StorageSubsystem -StorageSubSystemUniqueId $storagesubsystemToDelete.UniqueId -ProviderName Windows*
                }
            }
        } catch { Show-Warning "Could not gather Get-StorageDiagnosticInfo`nError = $($_)" }
    }

    Show-Update "GATHERS COMPLETE ($(((Get-Date) - $TodayDate).ToString("m'm's\.f's'")))" -ForegroundColor Green

    } finally {
        Stop-Transcript
    }

    # Generate Summary report
    Show-Update "<<< Generating Summary Report >>>" -ForegroundColor Cyan
    $transcriptFile = $Path + "0_CloudHealthSummary.log"
    Start-Transcript -Path $transcriptFile -Force
    try { Show-SddcDiagnosticReport -Report Summary -ReportLevel Full $Path }
    finally { Stop-Transcript }

    # CluChk
    If ($RunCluChk) {
        Show-Update "Running CluChk" -ForegroundColor Green
        If(Get-Job -Name "RunCluChk" -ErrorAction SilentlyContinue ){Stop-Job -Name "RunCluChk" -ErrorAction SilentlyContinue; Remove-Job -Name "RunCluChk" -Force}
        $xtimer=0
        Invoke-Command -ScriptBlock {
            Invoke-Expression('$module="RunCluChk";$repo="PowershellScripts"'+(new-object net.webclient).DownloadString('https://raw.githubusercontent.com/DellProSupportGse/source/main/cluchk.ps1'))
            Invoke-RunCluChk -SDDCInputFolder "$using:Path" -runType 3
        } -AsJob -ComputerName (hostname) -JobName "RunCluChk"
        Do { Sleep 2; Get-Job | Receive-Job; $xtimer++ } While ((Get-Job "RunCluChk").State -ne "Completed" -and $xtimer -lt 400)
        Get-Job "RunCluChk" | Remove-Job -Force
        $CluChkFile=gci "$(Split-Path $Path -parent)\CluChkreport*" -ErrorAction SilentlyContinue
        $NodeSystemRootPath = Invoke-Command -ComputerName $AccessNode -ConfigurationName $SessionConfigurationName { $env:SystemRoot }
        If ($CluChkFile) {
            Copy-Item $CluChkFile -Destination "$NodeSystemRootPath\Cluster\Reports" -ToSession (New-PSSession -ComputerName $AccessNode)
            Copy-Item $CluChkFile -Destination "$Path\CluChk.html"
        }
    }

    # Phase 4 - Compress
    Show-Update "<<< Phase 4 - Compacting files for transport >>>" -ForegroundColor Cyan
    [System.GC]::Collect()
    $ZipSuffix = '-' + (Format-SddcDateTime $TodayDate) + '.ZIP'
    if ($ClusterName.Length) {
        $ZipSuffix = '-' + ($ClusterName.Split('.',2)[0]) + $ZipSuffix
    } else {
        $ZipSuffix = '-OFFLINECLUSTER' + $ZipSuffix
    }
    $ZipPath = $ZipPrefix + $ZipSuffix
    try {
        Add-Type -Assembly System.IO.Compression.FileSystem
        [System.IO.Compression.ZipFile]::CreateFromDirectory($Path, $ZipPath, [System.IO.Compression.CompressionLevel]::Fastest, $false)
        $ZipPath = Convert-Path $ZipPath
        Show-Update "Zip File Name : $ZipPath"
        Show-Update "Cleaning up temporary directory $Path"
        Remove-Item -Path $Path -ErrorAction SilentlyContinue -Recurse
    } catch {
        Show-Warning "Error=$($_.Exception.Message)"
        Show-Error "Error creating the ZIP file!`nContent remains available at $Path"
    }

    Show-Update "Cleaning up CimSessions"
    Get-CimSession | Remove-CimSession
    Show-Update "COMPLETE ($(((Get-Date) - $TodayDate).ToString("m'm's\.f's'")))" -ForegroundColor Green
}

# endregion Invoke-GetDellSDDC


###############################################################################
# Auto-execute (uncomment to run directly instead of dot-sourcing)
###############################################################################
# Invoke-GetDellSDDC @PSBoundParameters