<#
.Synopsis
   DRIFT - Driver and Firmware Tool
   Script to check for the latest Firmware and Drivers
.DESCRIPTION
   This tool compares RAW Teseract export with
   the Dell catalog to easily show drivers and
   firmware DriFt from currently available
   versions on downloads.dell.com
.CREATEDBY
    Jim Gandy
.UPDATES
    2026/05/14:v1.79DEV -  1. New Feature: JG - Added 17G support    
    2026/05/14:v1.78 -  1. Bug Fix: JG - Resolved the telemety errors and enabled again
    
    See older version for previous notes
#>
Function Invoke-RunDriFT{
    $uploadToAzure=$True
# logging
$DateTime = Get-Date -Format yyyyMMdd_HHmmss
$LogRoot = Join-Path $env:ProgramData "Dell\DriFT"
if (-not (Test-Path -Path $LogRoot -PathType Container)) {
    New-Item -Path $LogRoot -ItemType Directory -Force | Out-Null
}
$LogPath = Join-Path $LogRoot "DriFT_$DateTime.log"
Start-Transcript -NoClobber -Path $LogPath
Write-Host "Starting log: $LogPath"
IF(!($args)){
    #Variable Cleanup
    # Avoid clearing the caller/session scope. Keep variables scoped inside Invoke-RunDriFT instead.
    # Remove-Variable * -ErrorAction SilentlyContinue
}
[system.gc]::Collect()
$DriFTVer="DriFT_v1.79DEV"
$DirFTV=$DriFTVer.Split("v")
$DFTV=$DirFTV[1]

#Param ($TSRIn)
#If($TSRIn.lenght -gt 0){$args=$TSRIn}
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

Function EndScript{  
    break
}
$WhatsNew=@"
    1. Bug Fix: TP - Fixed MS latest updates by copying and converting it from CluChk
"@

If(!($args)){Clear-Host}
$text = @"
v$DFTV                                           
_______                                     
\  ___ `'.           .--.                   
 ' |--.\  \          |__|     _.._          
 | |    \  ' .-,.--. .--.   .' .._|    .|   
 | |     |  '|  .-. ||  |   | '      .' |_  
 | |     |  || |  | ||  | __| |__  .'     | 
 | |     ' .'| |  | ||  ||__   __|'--.  .-' 
 | |___.' /' | |  '- |  |   | |      |  |   
/_______.'/  | |     |__|   | |      |  |   
\_______|/   | |            | |      |  '.' 
             |_|            | |      |   /  
                            |_|      `'-'   

                             by: Jim Gandy

"@
Write-Host $text
If($args){
    IF($args -match "cluchk"){
        Write-Host "DriFT running in CluChk mode..."
        Write-Host "    $args"
        $FileNameGuid=(($args -split '\-cluchk\s')[1] -split '-input')[0].trim()
        #$FileNameGuid=$args -replace '-cluchk ',""
        Write-Host "File Name Guid:" $FileNameGuid
        Write-Host "Processing TRS File(s)"
        $TSRInputFiles=@()
        $TSRInputFiles=(($args -split '-input')[1].trim() -split ',').trim()
        $TSRLoc=$TSRInputFiles
        $TSRLoc
        $args=""
        $CluChkMode="YES"
    }Else{
        Write-Host "CMD Mode: Processing one TSR..."
        Write-Host "TSR Input File: "$args
    }
}


#Input file
Function Get-FileName($initialDirectory)
{
    Add-Type -AssemblyName System.Windows.Forms
    
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog -Property @{Multiselect = $true}
    $OpenFileDialog.Title = "Please Select One or More SupportAssist File(s)."
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "ZIP (*.zip)| *.zip"
    $OpenFileDialog.ShowDialog((New-Object System.Windows.Forms.Form -Property @{TopMost = $true })) | Out-Null
    $OpenFileDialog.filenames
}

IF(!($CluChkMode)){
    If(!($args)){
        $Title=@()
        $Title+="Welcome to DriFT (Driver and Firmware Tool)"
        Write-host $Title
        Write-host " "
        Write-Host "What's New in"$DFTV":"
        Write-Host $WhatsNew 
        Write-Host "" 
        $Run = Read-Host "Ready to run? [y/n]"
        If (($run -ieq "n")-or ($run -ieq "")){
            $OutputType="No"
            EndScript}
    };
}


# =====================================================
#region Telemetry Information
# =====================================================
$uploadToAzure=$True
IF($uploadToAzure){

    Write-Host "Logging Telemetry Information..."

    function Add-TableData {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory=$true)]
            [string]$TableName,

            [Parameter(Mandatory=$true)]
            [string]$PartitionKey,

            [Parameter(Mandatory=$true)]
            [hashtable]$Data
        )

        if (-not $uploadToAzure) { return }

        $RowKey = [guid]::NewGuid().Guid
        
        $TableSvcSasUrl = 'https://gsetools.table.core.windows.net/?SECRET REMOVED'

        $uri = "https://gsetools.table.core.windows.net/$TableName$($TableSvcSasUrl.Substring($TableSvcSasUrl.IndexOf('?')))"

        $headers = @{
            "Accept"       = "application/json;odata=nometadata"
            "Content-Type" = "application/json"
            "x-ms-version" = "2019-02-02"
        }

        $Data["PartitionKey"] = $PartitionKey
        $Data["RowKey"]       = $RowKey

        $body = $Data | ConvertTo-Json -Depth 5

        $maxRetries = 3
        $attempt = 0
        $success = $false

        while (-not $success -and $attempt -lt $maxRetries) {

            try {
                Invoke-RestMethod -Method Post -Uri $uri -Headers $headers -Body $body | Out-Null
                $success = $true
                Write-Indent "Telemetry recorded successfully" 1 Green
            }
            catch {
                $attempt++

                if ($attempt -lt $maxRetries) {
                    Write-Indent "Retrying telemetry upload ($attempt/$maxRetries)..." 1 Yellow
                    Start-Sleep -Seconds 2
                }
                else {
                    Write-Indent "Telemetry upload failed after $maxRetries attempts" 1 Yellow
                }
            }
        }
    }

    function Write-Indent {
        param(
            [string]$Message,
            [int]$Level = 1,
            [string]$Color = "Gray"
        )

        $prefix = "  " * $Level
        Write-Host "$prefix$Message" -ForegroundColor $Color
    }

    # Unique report id
    $CReportID = [guid]::NewGuid().Guid


    Write-Indent "Resolving Geo Location..."

    try {
        if (-not $global:GeoCache) {
            $global:GeoCache = Invoke-RestMethod "https://ipwho.is/" -TimeoutSec 5
        }

        $response = $global:GeoCache

        if ($response.success -eq $true) {

            $country     = $response.country
            $countryCode = $response.country_code
            $region      = $response.region
            $city        = $response.city
            $latitude    = $response.latitude
            $longitude   = $response.longitude
            $timezone    = $response.timezone.id

            Write-Indent "Country: $country" 2
            Write-Indent "Region : $region" 2
        }
    }
    catch {
        Write-Indent "WARN: ipwho lookup failed" 2 Yellow
    }

    $data = @{
        Region       = $region
        DriftVersion = $DFTV
        ReportID     = $CReportID
        country      = $country
        countryCode  = $countryCode
        geoRegion    = $region
        city         = $city
        lat          = $latitude
        lon          = $longitude
        timezone     = $timezone
        Timestamp = (Get-Date).ToUniversalTime().ToString("o")
        HostOS = [System.Environment]::OSVersion.VersionString
        PSVersion = $PSVersionTable.PSVersion.ToString()
    }

    # We use tool name for this value
    $PartitionKey = "DriFT"

    Add-TableData `
        -TableName "DriftTelemetryData" `
        -PartitionKey $PartitionKey `
        -Data $data 

}
#endregion

#Variables
#$DellURL="https://dl.dell.com/"
$DellURL="https://downloads.dell.com/"

#Get the catalog.cab
$LocCabSize=1
$DownloadFile="$env:TEMP\Catalog.cab"
#$url = "http://dl.dell.com/catalog/Catalog.cab"
$url = "https://downloads.dell.com/catalog/Catalog.cab"
#Downloading a new Catalog.cab

# Added for proxy auth 
    $browser = New-Object System.Net.WebClient
    $browser.Proxy.Credentials =[System.Net.CredentialCache]::DefaultNetworkCredentials 
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

If($LocCabSize -eq "1"){
    Write-host "Downloading Catalog.cab...."
    $CatLocNA="NO"
    #Invoke-WebRequest -Uri $url -OutFile $DownloadFile
    Try{Invoke-WebRequest -Uri $url -OutFile $DownloadFile}
    Catch{
        $CatLocNA="YES"
        Write-Host "    WARNING: Catalog Source location NOT accessible. Please provide CATALOG.CAB file."-foregroundcolor Yellow
        Write-Host "    Or manually download from:"$url -foregroundcolor Yellow}
    Finally{
        #Ask for the catalog.cab
        If($CatLocNA -eq "YES"){
            Function Get-CatFile($initialDirectory)
            {
                Add-Type -AssemblyName System.Windows.Forms
                $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog -Property @{Multiselect = $true}
                $OpenFileDialog.Title = "Please Select CATALOG.CAB file..."
                $OpenFileDialog.initialDirectory = $initialDirectory
                $OpenFileDialog.filter = "CAB (*.cab)| *.cab"
                $OpenFileDialog.ShowDialog() | Out-Null
                $OpenFileDialog.filenames
            }
            $DownloadFile=Get-CatFile("C:")
            if(!$DownloadFile){
                $OutputType="No"
                EndScript}
        }
    }
}

#Check/Create temp DIR 
# Clean build workspace. All temporary/extracted files stay under %TEMP%\DriFT.
$DriFTTempRoot = Join-Path $env:TEMP "DriFT"
$DriFTExtractRoot = Join-Path $DriFTTempRoot "Extract"
$DriFTRedfishRoot = Join-Path $DriFTTempRoot "Redfish"
$DriFTCatalogRoot = Join-Path $DriFTTempRoot "Catalog"
$DriFTWorkRoot = Join-Path $DriFTTempRoot "Work"

foreach ($DriFTPath in @($DriFTTempRoot,$DriFTExtractRoot,$DriFTRedfishRoot,$DriFTCatalogRoot,$DriFTWorkRoot)) {
    if (-not (Test-Path $DriFTPath -PathType Container)) {
        New-Item -ItemType Directory -Force -Path $DriFTPath | Out-Null
    }
}

$DriFTRunRoot = Join-Path $DriFTExtractRoot ("Run_" + ([guid]::NewGuid().Guid.Substring(0,8)))
$ExtracLoc = $DriFTRunRoot
if (!(Test-Path $ExtracLoc -PathType Container)) {New-Item -ItemType Directory -Force -Path $ExtracLoc | Out-Null}

#Extract the cab
Write-host "Extracting Catalog.xml from CAB...."
if (Test-Path "$ExtracLoc\Catalog.xml") {Remove-Item "$ExtracLoc\Catalog.xml"}

Function Expand-Cab ($SourceFile,$TargetFolder,$Item){

    $ShellObject = New-Object -com shell.application
    $zipfolder = $ShellObject.namespace($sourceFile)
    $Item = $zipfolder.parsename("$Item")
    $TargetFolder = $ShellObject.namespace("$TargetFolder")
    $TargetFolder.copyhere($Item)
}
IF(!(Test-Path "$ExtracLoc\Catalog.xml")){Expand-Cab -SourceFile $DownloadFile -TargetFolder $ExtracLoc -Item "Catalog.xml"}

# Used to extract .gz files
Function DeGZip-File{
    Param(
        $infile
        )
    $outFile = $infile.Substring(0, $infile.LastIndexOfAny('.'))
    $input = New-Object System.IO.FileStream $inFile, ([IO.FileMode]::Open), ([IO.FileAccess]::Read), ([IO.FileShare]::Read)
    $output = New-Object System.IO.FileStream $outFile, ([IO.FileMode]::Create), ([IO.FileAccess]::Write), ([IO.FileShare]::None)
    $gzipStream = New-Object System.IO.Compression.GzipStream $input, ([IO.Compression.CompressionMode]::Decompress)

    $buffer = New-Object byte[](1024)
    while($true){
        $read = $gzipstream.Read($buffer, 0, 1024)
        if ($read -le 0){break}
        $output.Write($buffer, 0, $read)
        }

    $gzipStream.Close()
    $output.Close()
    $input.Close()
}


Function Get-DriFTFirstNonEmpty {
    param(
        [Parameter(ValueFromRemainingArguments=$true)]
        [object[]]$Values
    )

    foreach ($Value in $Values) {
        if ($null -eq $Value) { continue }
        if ($Value -is [array]) {
            foreach ($Item in $Value) {
                if ($null -ne $Item -and ([string]$Item).Trim().Length -gt 0) { return [string]$Item }
            }
        }
        elseif (([string]$Value).Trim().Length -gt 0) {
            return [string]$Value
        }
    }
    return $null
}

Function Export-DriFT17GDebug {
    param(
        [string]$Name,
        [AllowNull()][object]$InputObject
    )

    # Debug exports are disabled in the clean build.
    # Keep this function as a no-op so any remaining calls do not fail.
    $DriFTTempRoot = Join-Path $env:TEMP "DriFT"
    if (-not (Test-Path $DriFTTempRoot -PathType Container)) {
        New-Item -ItemType Directory -Force -Path $DriFTTempRoot | Out-Null
    }
    return $DriFTTempRoot
}

Function Get-DriFTComparableText {
    param([AllowNull()][object[]]$Values)

    $Text = @()
    foreach ($Value in @($Values)) {
        if ($null -eq $Value) { continue }
        $StringValue = ([string]$Value)
        if ([string]::IsNullOrWhiteSpace($StringValue)) { continue }

        $Text += $StringValue
        $Text += ($StringValue `
            -replace '\s+Firmware Inventory$','' `
            -replace '\s+Firmware$','' `
            -replace '\s+Controller$','' `
            -replace '\s+Adapter$','' `
            -replace '\s+Device$','').Trim()
    }

    @($Text | Where-Object { $_ -and $_.Trim().Length -ge 3 } | Sort-Object -Unique)
}




Function Get-DriFTCatalogComponentTypeValue {
    param([AllowNull()][object]$CatalogObject)

    return Get-DriFTFirstNonEmpty `
        $CatalogObject.ComponentType.value `
        $CatalogObject.ComponentType `
        $CatalogObject.componentType.value `
        $CatalogObject.componentType
}

Function Test-DriFTDeviceIsBiosFirmware {
    param([Parameter(Mandatory=$true)]$Device)

    $DeviceComponentId = Get-DriFTFirstNonEmpty $Device.componentID $Device.ComponentID
    $DeviceDisplay     = Get-DriFTFirstNonEmpty $Device.display $Device.Display
    $DeviceElementName = Get-DriFTFirstNonEmpty $Device.ElementName $Device.Name
    $DeviceRelatedItem = Get-DriFTFirstNonEmpty $Device.RelatedItem

    # 17G Redfish BIOS currently normalizes as componentType FRMW + componentID 159.
    # Also keep the text checks so this remains safe if Dell changes the id source later.
    if ($DeviceComponentId -eq '159') { return $true }
    if ($DeviceDisplay -imatch '^(Bios|BIOS|System BIOS|BIOS\.Setup)') { return $true }
    if ($DeviceElementName -imatch '^(BIOS|System BIOS)$') { return $true }
    if ($DeviceRelatedItem -imatch '/Bios/?$') { return $true }

    return $false
}

Function Test-DriFTCatalogComponentTypeCompatible {
    param(
        [AllowNull()][object]$CatalogObject,
        [Parameter(Mandatory=$true)]$Device
    )

    # $CatalogXMLDataFiltered is already filtered to the supported system/OS.
    # For normal inventory items, still require the catalog ComponentType to match the device
    # ComponentType so firmware rows do not accidentally pull driver packages.
    # Exception: BIOS from 17G is reported as FRMW in InstalledHardwareUnique, while Catalog.xml
    # has the BIOS row as ComponentType BIOS. For BIOS only, allow componentID-only matching.

    if (Test-DriFTDeviceIsBiosFirmware -Device $Device) { return $true }

    $CatalogType = Get-DriFTCatalogComponentTypeValue -CatalogObject $CatalogObject
    $DeviceType  = Get-DriFTFirstNonEmpty $Device.componentType $Device.ComponentType

    if ([string]::IsNullOrWhiteSpace($CatalogType) -or [string]::IsNullOrWhiteSpace($DeviceType)) { return $true }
    if ($CatalogType -ieq $DeviceType) { return $true }

    return $false
}

Function Get-DriFTCatalogComponentIdValues {
    param([AllowNull()][object]$CatalogObject)

    $Values = New-Object System.Collections.Generic.List[string]
    foreach ($Obj in @($CatalogObject)) {
        if ($null -eq $Obj) { continue }

        foreach ($Candidate in @(
            $Obj.ComponentID.value,
            $Obj.componentID.value,
            $Obj.ComponentID,
            $Obj.componentID,
            $Obj.SupportedDevices.Device.ComponentID.value,
            $Obj.SupportedDevices.Device.componentID.value,
            $Obj.SupportedDevices.Device.ComponentID,
            $Obj.SupportedDevices.Device.componentID
        )) {
            foreach ($Value in @($Candidate)) {
                if ($null -ne $Value -and -not [string]::IsNullOrWhiteSpace([string]$Value)) {
                    [void]$Values.Add(([string]$Value).Trim())
                }
            }
        }
    }

    return @($Values | Sort-Object -Unique)
}

Function Test-DriFTCatalogComponentIdMatch {
    param(
        [AllowNull()][object]$CatalogObject,
        [AllowNull()][object]$ComponentId
    )

    $ComponentIdText = Get-DriFTFirstNonEmpty $ComponentId
    if ([string]::IsNullOrWhiteSpace($ComponentIdText)) { return $false }

    # 17G FirmwareInventory sometimes reports SoftwareId/componentID as 0. That is not a
    # real Dell catalog componentID and must not be used as a regex/partial match. It caused
    # false positives such as NIC.Slot.2 firmware matching BOSS/PCIe SSD catalog rows.
    if ($ComponentIdText.Trim() -eq '0') { return $false }

    foreach ($CandidateText in @(Get-DriFTCatalogComponentIdValues -CatalogObject $CatalogObject)) {
        if ($CandidateText -ieq $ComponentIdText.Trim()) { return $true }
    }

    return $false
}

Function Get-DriFTCatalogPciInfoObjects {
    param([AllowNull()][object]$CatalogObject)

    $Rows = @()
    foreach ($Obj in @($CatalogObject)) {
        if ($null -eq $Obj) { continue }

        $Rows += @($Obj.PCIInfo)
        $Rows += @($Obj.SupportedDevices.Device.PCIInfo)

        # Some Catalog XML views surface PCI values directly as S/N nodes on the object.
        if ($Obj.vendorID -or $Obj.deviceID -or $Obj.subVendorID -or $Obj.subDeviceID) {
            $Rows += [PSCustomObject]@{
                vendorID    = (Get-DriFTFirstNonEmpty $Obj.vendorID.value $Obj.vendorID)
                deviceID    = (Get-DriFTFirstNonEmpty $Obj.deviceID.value $Obj.deviceID)
                subVendorID = (Get-DriFTFirstNonEmpty $Obj.subVendorID.value $Obj.subVendorID)
                subDeviceID = (Get-DriFTFirstNonEmpty $Obj.subDeviceID.value $Obj.subDeviceID)
            }
        }
    }

    return @($Rows | Where-Object { $_ })
}

Function Test-DriFTCatalogPciIdentityMatch {
    param(
        [AllowNull()][object]$CatalogObject,
        [Parameter(Mandatory=$true)]$Device
    )

    $DeviceVendorId    = Convert-DriFT17GHexId $Device.vendorID
    $DeviceDeviceId    = Convert-DriFT17GHexId $Device.deviceID
    $DeviceSubVendorId = Convert-DriFT17GHexId $Device.subVendorID
    $DeviceSubDeviceId = Convert-DriFT17GHexId $Device.subDeviceID

    if ([string]::IsNullOrWhiteSpace($DeviceVendorId) -or [string]::IsNullOrWhiteSpace($DeviceDeviceId)) { return $false }

    foreach ($PciInfo in @(Get-DriFTCatalogPciInfoObjects -CatalogObject $CatalogObject)) {
        if ($null -eq $PciInfo) { continue }

        $CatalogVendorId    = Convert-DriFT17GHexId (Get-DriFTFirstNonEmpty $PciInfo.vendorID.value $PciInfo.vendorID)
        $CatalogDeviceId    = Convert-DriFT17GHexId (Get-DriFTFirstNonEmpty $PciInfo.deviceID.value $PciInfo.deviceID)
        $CatalogSubVendorId = Convert-DriFT17GHexId (Get-DriFTFirstNonEmpty $PciInfo.subVendorID.value $PciInfo.subVendorID)
        $CatalogSubDeviceId = Convert-DriFT17GHexId (Get-DriFTFirstNonEmpty $PciInfo.subDeviceID.value $PciInfo.subDeviceID)

        if ($CatalogVendorId -ne $DeviceVendorId) { continue }
        if ($CatalogDeviceId -ne $DeviceDeviceId) { continue }

        # If the installed device has subsystem IDs, require the catalog to match them when present.
        if ($DeviceSubVendorId -and $CatalogSubVendorId -and ($CatalogSubVendorId -ne $DeviceSubVendorId)) { continue }
        if ($DeviceSubDeviceId -and $CatalogSubDeviceId -and ($CatalogSubDeviceId -ne $DeviceSubDeviceId)) { continue }

        return $true
    }

    return $false
}


Function Test-DriFTInstalledComponentIdIsValid {
    param([AllowNull()][object]$ComponentId)

    $ComponentIdText = Get-DriFTFirstNonEmpty $ComponentId
    if ([string]::IsNullOrWhiteSpace($ComponentIdText)) { return $false }
    if ($ComponentIdText.Trim() -eq '0') { return $false }
    return $true
}

Function Test-DriFTDeviceHasPciIdentity {
    param([Parameter(Mandatory=$true)]$Device)

    $DeviceVendorId    = Convert-DriFT17GHexId $Device.vendorID
    $DeviceDeviceId    = Convert-DriFT17GHexId $Device.deviceID
    $DeviceSubVendorId = Convert-DriFT17GHexId $Device.subVendorID
    $DeviceSubDeviceId = Convert-DriFT17GHexId $Device.subDeviceID

    return (-not [string]::IsNullOrWhiteSpace($DeviceVendorId) -and
            -not [string]::IsNullOrWhiteSpace($DeviceDeviceId) -and
            -not [string]::IsNullOrWhiteSpace($DeviceSubVendorId) -and
            -not [string]::IsNullOrWhiteSpace($DeviceSubDeviceId))
}

Function Test-DriFTCatalogDeviceMatch {
    param(
        [AllowNull()][object]$CatalogDevice,
        [Parameter(Mandatory=$true)]$Device
    )

    # Matching rule requested for DriFT 17G:
    #   Match catalog SupportedDevices.Device to the installed device when either:
    #     1. componentID matches exactly, OR
    #     2. PCI identity matches: vendorID + deviceID + subDeviceID + subVendorID.
    #
    # Important: componentID "0" from 17G Redfish is not a real catalog component ID,
    # so it is ignored to avoid false positives.

    # Keep the search scoped to the already-filtered catalog row set, but prevent cross-type
    # pollution. BIOS is the only exception that can ignore ComponentType because 17G reports it
    # as FRMW while Catalog.xml represents it as BIOS.
    if (-not (Test-DriFTCatalogComponentTypeCompatible -CatalogObject $CatalogDevice -Device $Device)) {
        return $false
    }

    $ComponentMatch = $false
    if (Test-DriFTInstalledComponentIdIsValid -ComponentId $Device.componentID) {
        $ComponentMatch = Test-DriFTCatalogComponentIdMatch -CatalogObject $CatalogDevice -ComponentId $Device.componentID
    }

    $PciMatch = $false
    if (Test-DriFTDeviceHasPciIdentity -Device $Device) {
        $PciMatch = Test-DriFTCatalogPciIdentityMatch -CatalogObject $CatalogDevice -Device $Device
    }

    return ($ComponentMatch -or $PciMatch)
}

Function Find-DriFT17GCatalogMatch {
    param(
        [AllowNull()][object[]]$CatalogRows,
        [Parameter(Mandatory=$true)]$Device
    )

    # IMPORTANT:
    # CatalogRows must already be filtered for the current supported system/OS
    # ($CatalogXMLDataFiltered, $S2DCatalogXMLDataFiltered, or $SpecialCatalogXMLDataFiltered).
    # Do not search the whole Catalog.xml here.
    if (-not $CatalogRows) { return $null }

    # CatalogRows are already filtered to the supported system/OS.
    # Do NOT filter by ComponentType here. Match only by componentID exact OR full PCI identity.
    $Hit = $CatalogRows |
        Where-Object { Test-DriFTCatalogDeviceMatch -CatalogDevice $_ -Device $Device } |
        Sort-Object {[DateTime]$_.releaseDate} |
        Select-Object -Last 1

    return $Hit
}

Function Expand-DriFT17GRedfishWalk {
    param(
        [Parameter(Mandatory=$true)][string]$TarGzPath,
        [Parameter(Mandatory=$true)][string]$DestinationRoot
    )

    # Extract the Redfish walk outside the TSR tree and use a very short folder name.
    # The Redfish archive contains deeply nested paths; extracting under the TSR folder
    # can push Windows PowerShell 5.1 over MAX_PATH during later recursive searches.
    $ShortRoot = Join-Path (Join-Path $env:TEMP "DriFT") "Redfish"
    if (-not (Test-Path $ShortRoot -PathType Container)) { New-Item -ItemType Directory -Force -Path $ShortRoot | Out-Null }
    $RedfishExtractRoot = Join-Path $ShortRoot (([guid]::NewGuid().Guid).Substring(0,8))
    New-Item -ItemType Directory -Force -Path $RedfishExtractRoot | Out-Null

    $TarCommand = Get-Command tar.exe -ErrorAction SilentlyContinue
    if (-not $TarCommand) { $TarCommand = Get-Command tar -ErrorAction SilentlyContinue }

    if ($TarCommand) {
        & $TarCommand.Source -xzf $TarGzPath -C $RedfishExtractRoot 2>$null
        if ($LASTEXITCODE -eq 0) { return $RedfishExtractRoot }
    }

    throw "Unable to extract redfishidracwalk.tar.gz. tar.exe was not available or failed."
}

Function Get-DriFT17GJsonFile {
    param(
        [Parameter(Mandatory=$true)][string]$RedfishRoot,
        [Parameter(Mandatory=$true)][string]$RelativePath
    )

    $CleanPath = $RelativePath.TrimStart('/').Replace('/', [System.IO.Path]::DirectorySeparatorChar)
    $JsonPath = Join-Path $RedfishRoot $CleanPath

    # Some 17G tar files extract with an extra top-level folder. Try exact path first,
    # then fall back to a recursive suffix match.
    if (-not (Test-Path $JsonPath -PathType Leaf)) {
        $Suffix = $RelativePath.TrimStart('/') -replace '/', [regex]::Escape([System.IO.Path]::DirectorySeparatorChar)
        $JsonPath = Get-ChildItem -Path $RedfishRoot -Filter 'index.json' -File -Recurse -Force -ErrorAction SilentlyContinue |
            Where-Object { $_.FullName -match $Suffix.Replace('\\','\') + '$' } |
            Select-Object -First 1 -ExpandProperty FullName
    }

    if ($JsonPath -and (Test-Path $JsonPath -PathType Leaf)) {
        try { return (Get-Content -Raw -Path $JsonPath | ConvertFrom-Json) }
        catch { return $null }
    }
    return $null
}


Function Convert-DriFT17GHexId {
    param([AllowNull()][object]$Value)

    if ($null -eq $Value) { return '' }
    $Text = ([string]$Value).Trim()
    if ([string]::IsNullOrWhiteSpace($Text)) { return '' }

    # Normalize common Redfish/Dell PCI id formats to uppercase hex without 0x.
    if ($Text -match '0x([0-9a-fA-F]+)') { $Text = $Matches[1] }
    $Text = $Text -replace '[^0-9a-fA-F]',''
    if ($Text.Length -eq 0) { return '' }
    return $Text.ToUpper()
}


Function Get-DriFT17GFqddVariants {
    param([AllowNull()][string]$Value)

    $Variants = New-Object System.Collections.Generic.List[string]
    if ([string]::IsNullOrWhiteSpace($Value)) { return @() }

    $Base = ([string]$Value).Trim()
    foreach ($Item in @($Base)) {
        if ([string]::IsNullOrWhiteSpace($Item)) { continue }
        [void]$Variants.Add($Item)

        # DellSoftwareInventory folders/Ids commonly look like:
        # DCIM_CURRENT_0x23_NIC.Slot.5-3-1
        # DCIM_CURRENT#NIC.Slot.5-3-1
        $Clean = $Item `
            -replace '^DCIM[:_]CURRENT_0x23_', '' `
            -replace '^DCIM[:_]INSTALLED_0x23_', '' `
            -replace '^DCIM[:_]PREVIOUS_0x23_', '' `
            -replace '^DCIM_CURRENT_0x23_', '' `
            -replace '^DCIM_INSTALLED_0x23_', '' `
            -replace '^DCIM_PREVIOUS_0x23_', '' `
            -replace '^DCIM_CURRENT#', '' `
            -replace '^DCIM_INSTALLED#', '' `
            -replace '^DCIM_PREVIOUS#', '' `
            -replace '^DCIM_CURRENT_', '' `
            -replace '^DCIM_INSTALLED_', '' `
            -replace '^DCIM_PREVIOUS_', ''
        if (-not [string]::IsNullOrWhiteSpace($Clean)) { [void]$Variants.Add($Clean) }

        # Also add the last path component if this is an @odata.id/path.
        if ($Clean -match '/') {
            $Last = (($Clean.TrimEnd('/') -split '/')[-1])
            if ($Last) { [void]$Variants.Add($Last) }
        }

        # Add the base slot identity so NIC.Slot.5-3-1 can match FirmwareInventory
        # rows that only say NIC.Slot.5.
        if ($Clean -match '^(NIC\.Slot\.\d+)-') { [void]$Variants.Add($Matches[1]) }
        if ($Clean -match '^(RAID\.[^-\/]+\.\d+(?:-\d+)?)') { [void]$Variants.Add($Matches[1]) }
        if ($Clean -match '^(Disk\.[^-\/]+)') { [void]$Variants.Add($Matches[1]) }
        if ($Clean -match '^(PSU\.Slot\.\d+)') { [void]$Variants.Add($Matches[1]) }
    }

    return @($Variants | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique)
}

Function Get-DriFT17GObjectKeys {
    param(
        [AllowNull()][object]$JsonObject,
        [AllowNull()][string]$FilePath
    )

    $Keys = New-Object System.Collections.Generic.List[string]

    foreach ($Candidate in @(
        $JsonObject.'@odata.id',
        $JsonObject.Id,
        $JsonObject.Name,
        $JsonObject.FQDD,
        $JsonObject.SoftwareId,
        $JsonObject.SoftwareID,
        $JsonObject.DeviceId,
        $JsonObject.DeviceID,
        $JsonObject.FunctionId,
        $JsonObject.FunctionID
    )) {
        if ($Candidate) { [void]$Keys.Add(([string]$Candidate).Trim()) }
    }

    if ($FilePath) {
        try {
            $LeafParent = Split-Path -Path (Split-Path -Path $FilePath -Parent) -Leaf
            if ($LeafParent) { [void]$Keys.Add($LeafParent) }
        } catch {}
    }

    # Add useful variants from odata paths.
    foreach ($Key in @($Keys.ToArray())) {
        if ($Key -match '/') {
            $Last = (($Key.TrimEnd('/') -split '/')[-1])
            if ($Last) { [void]$Keys.Add($Last) }
        }
        if ($Key -match '#') {
            $HashLast = (($Key -split '#')[-1])
            if ($HashLast) { [void]$Keys.Add($HashLast) }
        }
    }

    # Add Dell 17G cleaned FQDD variants, especially from DellSoftwareInventory IDs/folders.
    foreach ($Key in @($Keys.ToArray())) {
        foreach ($Variant in @(Get-DriFT17GFqddVariants -Value $Key)) {
            if ($Variant) { [void]$Keys.Add($Variant) }
        }
    }

    return @($Keys | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique)
}

Function Get-DriFT17GPCIInventoryMap {
    param(
        [Parameter(Mandatory=$true)][object[]]$AllJsonFiles
    )

    $PciRows = @()

    foreach ($JsonFile in @($AllJsonFiles)) {
        try { $Json = Get-Content -Raw -Path $JsonFile.FullName | ConvertFrom-Json }
        catch { continue }
        if (-not $Json) { continue }

        $OdataType = [string]$Json.'@odata.type'
        $PathText  = [string]$JsonFile.FullName

        # Redfish-standard names plus common Dell OEM variations.
        $VendorId = Get-DriFTFirstNonEmpty `
            $Json.VendorId $Json.VendorID $Json.PciVendorId $Json.PCIVendorID $Json.Oem.Dell.VendorId $Json.Oem.Dell.VendorID
        $DeviceId = Get-DriFTFirstNonEmpty `
            $Json.DeviceId $Json.DeviceID $Json.PciDeviceId $Json.PCIDeviceID $Json.Oem.Dell.DeviceId $Json.Oem.Dell.DeviceID
        $SubVendorId = Get-DriFTFirstNonEmpty `
            $Json.SubsystemVendorId $Json.SubsystemVendorID $Json.SubVendorId $Json.SubVendorID $Json.PciSubVendorId $Json.PCISubVendorID $Json.Oem.Dell.SubsystemVendorId $Json.Oem.Dell.SubsystemVendorID $Json.Oem.Dell.SubVendorID
        $SubDeviceId = Get-DriFTFirstNonEmpty `
            $Json.SubsystemId $Json.SubsystemID $Json.SubsystemDeviceId $Json.SubsystemDeviceID $Json.SubDeviceId $Json.SubDeviceID $Json.PciSubDeviceId $Json.PCISubDeviceID $Json.Oem.Dell.SubsystemId $Json.Oem.Dell.SubsystemID $Json.Oem.Dell.SubDeviceID

        # Dell 17G TSRs may carry PCI identity in DellSoftwareInventory instead of PCIeFunction objects.
        # Example:
        #   IdentityInfoType  = OrgID:ComponentType:VendorID:DeviceID:SubVendorID:SubDeviceID
        #   IdentityInfoValue = DCIM:firmware:14E4:1751:14E4:5045
        $IdentityMatchFound = $false
        if ($Json.IdentityInfoType -and $Json.IdentityInfoValue) {
            $IdentityTypes  = @($Json.IdentityInfoType)
            $IdentityValues = @($Json.IdentityInfoValue)

            for ($ii = 0; $ii -lt $IdentityTypes.Count; $ii++) {
                $IdentityTypeText  = [string]$IdentityTypes[$ii]
                $IdentityValueText = [string]$IdentityValues[[Math]::Min($ii, ($IdentityValues.Count - 1))]

                if ($IdentityTypeText -imatch 'VendorID:DeviceID:SubVendorID:SubDeviceID' -and
                    -not [string]::IsNullOrWhiteSpace($IdentityValueText)) {

                    $TypeParts  = $IdentityTypeText  -split ':'
                    $ValueParts = $IdentityValueText -split ':'
                    $IdentityHash = @{}

                    for ($ip = 0; $ip -lt $TypeParts.Count -and $ip -lt $ValueParts.Count; $ip++) {
                        $IdentityHash[$TypeParts[$ip]] = $ValueParts[$ip]
                    }

                    # IMPORTANT:
                    # DellSoftwareInventory IdentityInfoValue is the authoritative catalog identity
                    # for 17G. Preserve these as HEX strings. Do not keep/convert the numeric
                    # Redfish PCIeFunction values such as 32902/22448/3, because the Dell catalog
                    # expects 8086/57B0/0003 style hex IDs.
                    $VendorId    = $IdentityHash['VendorID']
                    $DeviceId    = $IdentityHash['DeviceID']
                    $SubVendorId = $IdentityHash['SubVendorID']
                    $SubDeviceId = $IdentityHash['SubDeviceID']
                    $IdentityMatchFound = $true
                    break
                }
            }
        }

        # Skip objects that clearly are not PCI identity records.
        if (-not $VendorId -and -not $DeviceId -and -not $SubVendorId -and -not $SubDeviceId) { continue }
        if ((-not $IdentityMatchFound) -and
            ($OdataType -and $OdataType -notmatch 'PCIeFunction|PCIeDevice|DellPCIeFunction|NetworkDeviceFunction|NetworkAdapter|Storage|SoftwareInventory') -and
            ($PathText -notmatch 'PCIe|DellPCIe|NetworkAdapters|NetworkDeviceFunctions|Storage|DellSoftwareInventory')) { continue }

        $Keys = Get-DriFT17GObjectKeys -JsonObject $Json -FilePath $JsonFile.FullName
        $CleanIdVariants = @(Get-DriFT17GFqddVariants -Value ([string]$Json.Id))
        $CleanFolderVariants = @()
        try { $CleanFolderVariants = @(Get-DriFT17GFqddVariants -Value (Split-Path -Path (Split-Path -Path $JsonFile.FullName -Parent) -Leaf)) } catch {}

        # Prefer the cleaned Dell FQDD over the raw DellSoftwareInventory folder/id.
        # Example: DCIM_CURRENT_0x23_NIC.Slot.2-1-1 should store FQDD as NIC.Slot.2-1-1.
        $AllFqddVariants = @($Json.FQDD, $Json.Oem.Dell.FQDD) + $CleanIdVariants + $CleanFolderVariants
        $CleanFqdd = @($AllFqddVariants | Where-Object {
            $_ -match '^(NIC\.Slot\.\d+(?:-\d+-\d+)?)$' -or
            $_ -match '^(RAID\.[^\/]+)$' -or
            $_ -match '^(Disk\.[^\/]+)$' -or
            $_ -match '^(PSU\.Slot\.\d+)$'
        } | Sort-Object { $_.Length } -Descending | Select-Object -First 1)

        if (-not $CleanFqdd) {
            $CleanFqdd = Get-DriFTFirstNonEmpty $Json.FQDD $Json.Oem.Dell.FQDD $CleanIdVariants $CleanFolderVariants
        }

        foreach ($Variant in @($CleanIdVariants + $CleanFolderVariants)) {
            if ($Variant) { $Keys += $Variant }
        }
        $Keys = @($Keys | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique)

        $PciRows += [PSCustomObject]@{
            OdataId       = [string]$Json.'@odata.id'
            Id            = [string]$Json.Id
            Name          = [string]$Json.Name
            FQDD          = [string]$CleanFqdd
            SoftwareId    = [string]$Json.SoftwareId
            OdataType     = $OdataType
            VendorID      = Convert-DriFT17GHexId $VendorId
            DeviceID      = Convert-DriFT17GHexId $DeviceId
            SubVendorID   = Convert-DriFT17GHexId $SubVendorId
            SubDeviceID   = Convert-DriFT17GHexId $SubDeviceId
            KeyText       = (@($Keys) -join '|')
            SourceFile    = [string]$JsonFile.FullName
        }
    }

    return @($PciRows | Sort-Object OdataId,Id,FQDD,VendorID,DeviceID,SubVendorID,SubDeviceID -Unique)
}

Function Find-DriFT17GPCIRecordForFirmware {
    param(
        [Parameter(Mandatory=$true)]$FirmwareRow,
        [Parameter(Mandatory=$true)][object[]]$PciRows
    )

    if (-not $PciRows -or @($PciRows).Count -eq 0) { return $null }

    $Related = [string]$FirmwareRow.RelatedItem
    $Display = [string]$FirmwareRow.display
    $Element = [string]$FirmwareRow.ElementName
    $ComponentId = [string]$FirmwareRow.componentID

    $Needles = New-Object System.Collections.Generic.List[string]
    foreach ($Candidate in @($Related, $Display, $Element, $ComponentId)) {
        if (-not [string]::IsNullOrWhiteSpace($Candidate)) {
            [void]$Needles.Add($Candidate.Trim())
            if ($Candidate -match '/') { [void]$Needles.Add((($Candidate.TrimEnd('/') -split '/')[-1])) }
            foreach ($Variant in @(Get-DriFT17GFqddVariants -Value $Candidate)) {
                if (-not [string]::IsNullOrWhiteSpace($Variant)) { [void]$Needles.Add($Variant.Trim()) }
            }
        }
    }

    $Needles = @($Needles | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique)

    # Prefer DellSoftwareInventory over generic PCIeFunction objects.
    # Generic Redfish PCIeFunction objects may expose decimal IDs (32902/22448/3).
    # DellSoftwareInventory IdentityInfoValue exposes the catalog-ready hex IDs (8086/57B0/0003).
    $PrioritizedPciRows = @($PciRows | Sort-Object @{Expression={ if ($_.SourceFile -imatch 'DellSoftwareInventory') { 0 } else { 1 } }}, FQDD, Id)

    # Most important 17G case:
    # FirmwareInventory may only say NIC.Slot.2, while DellSoftwareInventory has
    # DCIM_CURRENT_0x23_NIC.Slot.2-1-1 with the exact PCI identity.
    # Do this before the generic KeyText match, because KeyText may contain broad base
    # slot aliases and otherwise pick the wrong NIC.Slot.* child.
    foreach ($Needle in $Needles) {
        if ($Needle -match '^NIC\.Slot\.\d+$') {
            $EscapedBase = [regex]::Escape($Needle)

            $Hit = $PrioritizedPciRows | Where-Object {
                ($_.SourceFile -imatch 'DellSoftwareInventory') -and
                ($_.VendorID -or $_.DeviceID -or $_.SubVendorID -or $_.SubDeviceID) -and
                (
                    ($_.FQDD -imatch "^$EscapedBase(?:-|$)") -or
                    ($_.Id -imatch "DCIM_CURRENT.*$EscapedBase(?:-|$)") -or
                    ($_.KeyText -imatch "(^|\|)$EscapedBase(?:-|\||$)")
                )
            } | Sort-Object `
                @{Expression={ if ($_.FQDD -ieq $Needle) { 0 } elseif ($_.FQDD -imatch "^$EscapedBase-") { 1 } else { 2 } }},
                FQDD |
                Select-Object -First 1

            if ($Hit) { return $Hit }
        }
    }

    # Also prioritize exact cleaned FQDD hits from DellSoftwareInventory for full values like NIC.Slot.2-1-1.
    foreach ($Needle in $Needles) {
        $Escaped = [regex]::Escape($Needle)
        $Hit = $PrioritizedPciRows | Where-Object {
            ($_.SourceFile -imatch 'DellSoftwareInventory') -and
            ($_.VendorID -or $_.DeviceID -or $_.SubVendorID -or $_.SubDeviceID) -and
            (
                ($_.FQDD -ieq $Needle) -or
                ($_.Id -ieq $Needle) -or
                ($_.OdataId -ieq $Needle) -or
                ($_.KeyText -imatch "(^|\|)$Escaped(\||$)")
            )
        } | Select-Object -First 1

        if ($Hit) { return $Hit }
    }

    # Exact odata/path/id/FQDD match against the broader Redfish PCI map.
    foreach ($Needle in $Needles) {
        $Escaped = [regex]::Escape($Needle)
        $Hit = $PrioritizedPciRows | Where-Object {
            ($_.OdataId -and $_.OdataId -ieq $Needle) -or
            ($_.Id -and $_.Id -ieq $Needle) -or
            ($_.SoftwareId -and $_.SoftwareId -ieq $Needle) -or
            ($_.FQDD -and $_.FQDD -ieq $Needle) -or
            ($_.KeyText -and $_.KeyText -imatch "(^|\|)$Escaped(\||$)")
        } | Select-Object -First 1
        if ($Hit) { return $Hit }
    }

    # Conservative contains match for FQDD-ish values.
    foreach ($Needle in $Needles) {
        if ($Needle.Length -lt 5) { continue }
        $Escaped = [regex]::Escape($Needle)
        $Hit = $PrioritizedPciRows | Where-Object {
            ($_.OdataId -and $_.OdataId -imatch $Escaped) -or
            ($_.Id -and $_.Id -imatch $Escaped) -or
            ($_.SoftwareId -and $_.SoftwareId -imatch $Escaped) -or
            ($_.FQDD -and $_.FQDD -imatch $Escaped) -or
            ($_.KeyText -and $_.KeyText -imatch $Escaped)
        } | Select-Object -First 1
        if ($Hit) { return $Hit }
    }

    return $null
}

Function Add-DriFT17GPciIdentityToFirmwareRows {
    param(
        [Parameter(Mandatory=$true)][object[]]$FirmwareRows,
        [Parameter(Mandatory=$true)][object[]]$PciRows
    )

    $Enriched = @()
    foreach ($Row in @($FirmwareRows)) {
        $Pci = Find-DriFT17GPCIRecordForFirmware -FirmwareRow $Row -PciRows $PciRows

        $OutVendorID    = $Row.vendorID
        $OutDeviceID    = $Row.deviceID
        $OutSubDeviceID = $Row.subDeviceID
        $OutSubVendorID = $Row.subVendorID
        $OutSource      = $Row.Source
        $OutPciMatchId  = ''
        $OutPciType     = ''

        if ($Pci) {
            if ($Pci.VendorID)    { $OutVendorID    = $Pci.VendorID }
            if ($Pci.DeviceID)    { $OutDeviceID    = $Pci.DeviceID }
            if ($Pci.SubDeviceID) { $OutSubDeviceID = $Pci.SubDeviceID }
            if ($Pci.SubVendorID) { $OutSubVendorID = $Pci.SubVendorID }
            $OutSource     = 'RedfishFirmwareInventory+PCI'
            $OutPciMatchId = Get-DriFTFirstNonEmpty $Pci.FQDD $Pci.Id $Pci.OdataId
            $OutPciType    = $Pci.OdataType
        }

        $Enriched += [PSCustomObject]@{
            componentType = $Row.componentType
            componentID   = $Row.componentID
            vendorID      = $OutVendorID
            deviceID      = $OutDeviceID
            subDeviceID   = $OutSubDeviceID
            subVendorID   = $OutSubVendorID
            version       = $Row.version
            display       = $Row.display
            ElementName   = $Row.ElementName
            RelatedItem   = $Row.RelatedItem
            Source        = $OutSource
            PciMatchId    = $OutPciMatchId
            PciMatchType  = $OutPciType
        }
    }

    return @($Enriched)
}


Function Import-DriFT17GRedfishRoot {
    param(
        [Parameter(Mandatory=$true)][string]$RedfishRoot,
        [AllowNull()][string]$WorkFolder
    )

    $SystemJson  = Get-DriFT17GJsonFile -RedfishRoot $RedfishRoot -RelativePath 'redfish/v1/Systems/System.Embedded.1/index.json'

    $AllJsonFiles = @(Get-ChildItem -Path $RedfishRoot -Filter '*.json' -File -Recurse -Force -ErrorAction SilentlyContinue)
    $DebugPath = Export-DriFT17GDebug -Name "DriFT_17G_AllJsonFiles_Debug.csv" -InputObject ($AllJsonFiles | Select-Object FullName,Length)

    # Prefer the FirmwareInventory collection index, but do not depend on it. Some extracted
    # trees have a different top-level folder or omit the collection index in the expected path.
    $FwIndex = Get-DriFT17GJsonFile -RedfishRoot $RedfishRoot -RelativePath 'redfish/v1/UpdateService/FirmwareInventory/index.json'

    $FirmwareJsonFiles = @()
    if ($FwIndex -and $FwIndex.Members) {
        foreach ($Member in @($FwIndex.Members)) {
            $MemberPath = [string]$Member.'@odata.id'
            if ([string]::IsNullOrWhiteSpace($MemberPath)) { continue }
            $MemberPath = $MemberPath.TrimStart('/')
            if ($MemberPath -notmatch 'index\.json$') { $MemberPath = $MemberPath.TrimEnd('/') + '/index.json' }
            $CleanPath = $MemberPath.Replace('/', [System.IO.Path]::DirectorySeparatorChar)
            $Candidate = Join-Path $RedfishRoot $CleanPath
            if (Test-Path $Candidate -PathType Leaf) { $FirmwareJsonFiles += Get-Item $Candidate }
        }
    }

    # Recursive fallback. Use both path-based and content-based discovery. Some 17G TSRs do not
    # unpack exactly as redfish/v1/UpdateService/FirmwareInventory/<id>/index.json.
    if (-not $FirmwareJsonFiles -or @($FirmwareJsonFiles).Count -eq 0) {
        $FirmwareJsonFiles = @($AllJsonFiles | Where-Object {
            ($_.FullName -imatch 'UpdateService.*FirmwareInventory') -and
            ($_.DirectoryName -notmatch '[\\/]FirmwareInventory$')
        })
    }

    if (-not $FirmwareJsonFiles -or @($FirmwareJsonFiles).Count -eq 0) {
        $ContentCandidates = @()
        foreach ($JsonFile in @($AllJsonFiles)) {
            try {
                $Raw = Get-Content -Raw -Path $JsonFile.FullName
                if (($Raw -imatch '"@odata\.type"\s*:\s*".*SoftwareInventory') -or
                    ($Raw -imatch '"SoftwareId"\s*:') -or
                    ($Raw -imatch '"Updateable"\s*:')) {
                    $ContentCandidates += $JsonFile
                }
            }
            catch {}
        }
        $FirmwareJsonFiles = $ContentCandidates
    }

    Export-DriFT17GDebug -Name "DriFT_17G_FirmwareCandidateFiles_Debug.csv" -InputObject ($FirmwareJsonFiles | Select-Object FullName,Length) | Out-Null

    $FirmwareRows = @()
    $RawFirmwareDebug = @()

    foreach ($FwFile in @($FirmwareJsonFiles)) {
        try { $FwItem = Get-Content -Raw -Path $FwFile.FullName | ConvertFrom-Json }
        catch { continue }

        if (-not $FwItem) { continue }

        $OdataType = [string]$FwItem.'@odata.type'
        if ($FwItem.Id -match '^Previous-') { continue }

        # Skip collection indexes.
        if ($FwItem.Members -and -not $FwItem.Version -and -not $FwItem.SoftwareId) { continue }

        # Keep only likely firmware/software inventory records.
        if (-not $FwItem.Version -and -not $FwItem.SoftwareId -and -not $FwItem.Name) { continue }
        if (($OdataType -and $OdataType -notmatch 'SoftwareInventory|FirmwareInventory') -and
            (-not $FwItem.SoftwareId) -and (-not $FwItem.Updateable)) { continue }

        $RelatedPath = $null
        if ($FwItem.RelatedItem) {
            $RelatedPath = @($FwItem.RelatedItem | ForEach-Object { $_.'@odata.id' } | Where-Object { $_ })[0]
        }

        $Display = $null
        if ($RelatedPath) { $Display = (($RelatedPath.TrimEnd('/') -split '/')[-1]) }
        if (-not $Display) { $Display = $FwItem.Id }
        if (-not $Display) { $Display = $FwItem.Name }

        $ElementName = [string]$FwItem.Name
        if ($ElementName) { $ElementName = ($ElementName -replace '\s+Firmware Inventory$','').Trim() }
        if (-not $ElementName) { $ElementName = $Display }

        $ComponentId = Get-DriFTFirstNonEmpty $FwItem.SoftwareId $FwItem.Id

        $RawFirmwareDebug += [PSCustomObject]@{
            File        = $FwFile.FullName
            OdataType   = $OdataType
            Id          = [string]$FwItem.Id
            Name        = [string]$FwItem.Name
            SoftwareId  = [string]$FwItem.SoftwareId
            Version     = [string]$FwItem.Version
            RelatedItem = [string]$RelatedPath
        }

        $FirmwareRows += [PSCustomObject]@{
            componentType = 'FRMW'
            componentID   = [string]$ComponentId
            vendorID      = ''
            deviceID      = ''
            subDeviceID   = ''
            subVendorID   = ''
            version       = [string]$FwItem.Version
            display       = [string]$Display
            ElementName   = [string]$ElementName
            RelatedItem   = [string]$RelatedPath
            Source        = 'RedfishFirmwareInventory'
        }
    }

    $FirmwareRows = @($FirmwareRows |
        Where-Object { ($_.componentID.Length -gt 0) -or ($_.version.Length -gt 0) } |
        Sort-Object componentType,componentID,vendorID,deviceID,subDeviceID,subVendorID,version,display -Unique)

    # Phase 4: build a Redfish PCI identity map and enrich firmware rows where RelatedItem/FQDD can be correlated.
    $PciInventoryRows = @(Get-DriFT17GPCIInventoryMap -AllJsonFiles $AllJsonFiles)
    Export-DriFT17GDebug -Name "DriFT_17G_PCIInventory_Debug.csv" -InputObject $PciInventoryRows | Out-Null

    if ($PciInventoryRows.Count -gt 0 -and $FirmwareRows.Count -gt 0) {
        $FirmwareRows = @(Add-DriFT17GPciIdentityToFirmwareRows -FirmwareRows $FirmwareRows -PciRows $PciInventoryRows)
    }

    $PciMatchedCount = @($FirmwareRows | Where-Object { $_.vendorID -or $_.deviceID -or $_.subVendorID -or $_.subDeviceID }).Count
    $DellSoftwareIdentityCount = @($PciInventoryRows | Where-Object { $_.SourceFile -imatch 'DellSoftwareInventory' }).Count
    Write-Host "    17G PCI identity records found: $($PciInventoryRows.Count)"
    Write-Host "    17G DellSoftwareInventory identity records found: $DellSoftwareIdentityCount"
    Write-Host "    17G firmware rows enriched with PCI IDs: $PciMatchedCount"

    $PciCorrelationDebug = @($FirmwareRows | Select-Object display,ElementName,componentID,version,RelatedItem,vendorID,deviceID,subVendorID,subDeviceID,PciMatchId,PciMatchType,Source)

    Export-DriFT17GDebug -Name "DriFT_17G_FirmwareInventory_Debug.csv" -InputObject $RawFirmwareDebug | Out-Null
    Export-DriFT17GDebug -Name "DriFT_17G_PCI_Correlation_Debug.csv" -InputObject $PciCorrelationDebug | Out-Null
    Export-DriFT17GDebug -Name "DriFT_17G_InstalledHardwareUnique_Debug.csv" -InputObject $FirmwareRows | Out-Null

    

    [PSCustomObject]@{
        ExtractedPath = $RedfishRoot
        System        = $SystemJson
        Firmware      = $FirmwareRows
        PciInventory  = $PciInventoryRows
        DebugPath     = $DebugPath
    }
}


Function Expand-DriFT17GViewerHtmlRedfishWalk {
    param(
        [Parameter(Mandatory=$true)][string]$ViewerHtmlPath
    )

    $ShortRoot = Join-Path (Join-Path $env:TEMP "DriFT") "Redfish"
    if (-not (Test-Path $ShortRoot -PathType Container)) { New-Item -ItemType Directory -Force -Path $ShortRoot | Out-Null }
    $RedfishExtractRoot = Join-Path $ShortRoot (([guid]::NewGuid().Guid).Substring(0,8))
    New-Item -ItemType Directory -Force -Path $RedfishExtractRoot | Out-Null

    $ViewerRaw = Get-Content -Raw -Path $ViewerHtmlPath
    $DebugPath = Export-DriFT17GDebug -Name "DriFT_17G_ViewerHtml_Path_Debug.txt" -InputObject $ViewerHtmlPath

    $RedfishScriptMatch = [regex]::Match(
        $ViewerRaw,
        '<script\s+[^>]*content=["'']redfish["''][^>]*>(?<body>.*?)</script>',
        [System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [System.Text.RegularExpressions.RegexOptions]::Singleline
    )

    if (-not $RedfishScriptMatch.Success) {
        throw "viewer.html did not contain a script tag with content='redfish'."
    }

    $Base64Text = $RedfishScriptMatch.Groups['body'].Value.Trim()
    $ZipPath = Join-Path $RedfishExtractRoot "viewer_redfishwalk.zip"

    try {
        [System.IO.File]::WriteAllBytes($ZipPath, [System.Convert]::FromBase64String($Base64Text))
    }
    catch {
        throw "Failed to decode viewer.html embedded redfish zip: $($_.Exception.Message)"
    }

    try {
        Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue

        # Do NOT use ZipFile.ExtractToDirectory here. The 17G viewer.html embedded ZIP
        # contains Redfish/Dell object paths with Windows-invalid filename characters
        # such as ':' in entries like:
        #   DellDrives/Disk.Bay.0:Enclosure.Internal.0-1:RAID.SL.1-1/index.json
        # ExtractToDirectory throws: "The given path's format is not supported."
        # Manually extract and sanitize only the filename/path characters so the
        # parser can still recursively scan the JSON content.
        $InvalidCharsPattern = '[<>:"|?*]'
        $ZipArchive = [System.IO.Compression.ZipFile]::OpenRead($ZipPath)
        try {
            foreach ($Entry in $ZipArchive.Entries) {
                if ([string]::IsNullOrWhiteSpace($Entry.FullName)) { continue }

                $SafeRelativePath = ($Entry.FullName -replace '/', [System.IO.Path]::DirectorySeparatorChar) -replace $InvalidCharsPattern, '_'
                $SafeRelativePath = $SafeRelativePath.TrimStart([System.IO.Path]::DirectorySeparatorChar)
                if ([string]::IsNullOrWhiteSpace($SafeRelativePath)) { continue }

                $DestinationPath = Join-Path $RedfishExtractRoot $SafeRelativePath

                # Directory entry
                if ($Entry.FullName.EndsWith('/') -or [string]::IsNullOrWhiteSpace($Entry.Name)) {
                    if (-not (Test-Path $DestinationPath -PathType Container)) {
                        New-Item -ItemType Directory -Force -Path $DestinationPath | Out-Null
                    }
                    continue
                }

                $DestinationDirectory = Split-Path -Path $DestinationPath -Parent
                if (-not (Test-Path $DestinationDirectory -PathType Container)) {
                    New-Item -ItemType Directory -Force -Path $DestinationDirectory | Out-Null
                }

                [System.IO.Compression.ZipFileExtensions]::ExtractToFile($Entry, $DestinationPath, $true)
            }
        }
        finally {
            if ($ZipArchive) { $ZipArchive.Dispose() }
        }
    }
    catch {
        throw "Failed to extract viewer.html embedded redfish zip: $($_.Exception.Message)"
    }

    return $RedfishExtractRoot
}

Function Import-DriFT17GViewerHtml {
    param(
        [Parameter(Mandatory=$true)][string]$ViewerHtmlPath,
        [AllowNull()][string]$WorkFolder
    )

    $RedfishRoot = Expand-DriFT17GViewerHtmlRedfishWalk -ViewerHtmlPath $ViewerHtmlPath
    $Result = Import-DriFT17GRedfishRoot -RedfishRoot $RedfishRoot -WorkFolder $WorkFolder

    # Keep the source visible in the debug output.
    if ($Result) {
        $Result | Add-Member -MemberType NoteProperty -Name ViewerHtmlPath -Value $ViewerHtmlPath -Force
    }
    return $Result
}

Function Import-DriFT17GRedfishWalk {
    param(
        [Parameter(Mandatory=$true)][string]$TarGzPath,
        [Parameter(Mandatory=$true)][string]$WorkFolder
    )

    $RedfishRoot = Expand-DriFT17GRedfishWalk -TarGzPath $TarGzPath -DestinationRoot $WorkFolder
    return Import-DriFT17GRedfishRoot -RedfishRoot $RedfishRoot -WorkFolder $WorkFolder
}

# =====================================================
#region DriFT normalized report engine
# =====================================================
function New-DriFTSystemInfo {
    [CmdletBinding()]
    param(
        [string]$ServiceTag,
        [string]$PowerEdge,
        [string]$OS,
        [string]$HostName,
        [string]$SystemID,
        [string]$SourceType
    )

    [PSCustomObject]@{
        ServiceTag = $ServiceTag
        PowerEdge  = $PowerEdge
        OS         = $OS
        HostName   = $HostName
        SystemID   = $SystemID
        SourceType = $SourceType
    }
}

function New-DriFTReportRow {
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

function ConvertTo-DriFTTooltipHtml {
    [CmdletBinding()]
    param([AllowNull()][string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) { return $Text }

    if ($Text -match '-') {
        $SplitPos = $Text.IndexOf('-')
        if ($SplitPos -gt 0) {
            $Summary = $Text.Substring(0,$SplitPos)
            $Detail  = $Text.Substring($SplitPos + 1)
            return "<div class='tooltip'>$Summary<span class='tooltiptext'>$Detail</span>"
        }
    }

    return $Text
}

function ConvertTo-DriFTReportTableRows {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [AllowEmptyCollection()]
        [object[]]$ReportRows
    )

    $ReportRows |
        Sort-Object ServiceTag,PowerEdge,Type,Category |
        Select-Object ServiceTag,PowerEdge,OS,Type,Category,Name,InstalledVersion,AvailableVersion,CatalogInfo,
            @{Label='Criticality';Expression={ ConvertTo-DriFTTooltipHtml -Text $_.Criticality }},
            ReleaseDate,
            @{Label='Documentation';Expression={
                if ($_.Details -and $_.Details.ToString().Length -gt 0) {
                    if ($_.Details -notmatch '<br>') { "<a href='$($_.Details)' target='_blank'>Link</a>" } else { $_.Details }
                }
            }},
            @{Label='Download Link';Expression={
                if ($_.URL -and $_.URL.ToString().Length -gt 0 -and $_.URL -inotmatch 'href') { "<a href='$($_.URL)'>$($_.URL)</a>" } else { $_.URL }
            }}
}

function New-DriFTCompareView {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [AllowEmptyCollection()]
        [object[]]$ReportRows
    )

    $NewReportView = New-Object System.Data.DataTable 'NodeCompare'
    foreach ($ColumnName in @('Type','Name','AvailableVersion','CatalogInfo','Criticality','ReleaseDate')) {
        [void]$NewReportView.Columns.Add((New-Object System.Data.DataColumn($ColumnName)))
    }

    foreach ($Node in (($ReportRows.ServiceTag | Sort-Object -Unique) -replace '\*')) {
        if (-not [string]::IsNullOrWhiteSpace($Node) -and -not $NewReportView.Columns.Contains([string]$Node)) {
            [void]$NewReportView.Columns.Add((New-Object System.Data.DataColumn([string]$Node)))
        }
    }

    $CurrentName = $null
    $Row = $null
    foreach ($Item in ($ReportRows | Sort-Object Name)) {
        if ([string]::IsNullOrWhiteSpace($Item.Name) -or $Item.Name.Length -le 10 -or $Item.Name -match 'System.__ComObject') { continue }

        if ($Item.Name -ne $CurrentName) {
            if ($null -ne $Row) { [void]$NewReportView.Rows.Add($Row) }
            $CurrentName = $Item.Name
            $Row = $NewReportView.NewRow()
            $Row['Type'] = $Item.Type
            $Row['Name'] = "<a href='$($Item.Details)' target='_blank'>$($Item.Name)</a>"
            $Row['AvailableVersion'] = "<a href='$($Item.URL)'>$($Item.AvailableVersion)</a>"
            $Row['CatalogInfo'] = $Item.CatalogInfo
            $Row['Criticality'] = ConvertTo-DriFTTooltipHtml -Text $Item.Criticality
            $Row['ReleaseDate'] = $Item.ReleaseDate
        }

        $NodeColumn = ($Item.ServiceTag -replace '\*')
        if ($NewReportView.Columns.Contains($NodeColumn)) {
            $Row[$NodeColumn] = $Item.InstalledVersion
        }
    }

    if ($null -ne $Row) { [void]$NewReportView.Rows.Add($Row) }
    return $NewReportView
}

function Get-DriFTHtmlHeader {
    @"
<style TYPE="text/css">
TABLE {border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}
TH {border-width: 1px;padding: 3px;border-style: solid;border-color: black;background-color: #6495ED;}
TD {border-width: 1px;padding: 3px;border-style: solid;border-color: black;}
TR:Nth-Child(Even) {Background-Color: #dddddd;}
TR:Hover TD {Background-Color: #C1D5F8;}
.tooltip {
    position: relative;
    display: inline-block;
    border-bottom: 1px dotted black;
  }
  .tooltip .tooltiptext {
    visibility: hidden;
    width: 120px;
    background-color: black;
    color: #fff;
    text-align: center;
    border-radius: 6px;
    padding: 5px 0;
    position: absolute;
    z-index: 1;
  }
  .tooltip:hover .tooltiptext {
    visibility: visible;
  }
</style>
<title>
DriFT Report
</title>
"@
}

function New-DriFTReportFooter {
    [CmdletBinding()]
    param(
        [AllowEmptyCollection()]
        [object[]]$ReportRows,
        [string]$ServerType,
        [string]$S2DCatalogNeeded,
        [AllowNull()][object]$NoneSupportedDevices,
        [AllowNull()][string]$CatVerInfo
    )

    if ($NoneSupportedDevices -is [array]) {
        $NoneSupportedDevices = ($NoneSupportedDevices | Where-Object { $_ } | ForEach-Object { [string]$_ }) -join ','
    }
    elseif ($null -ne $NoneSupportedDevices) {
        $NoneSupportedDevices = [string]$NoneSupportedDevices
    }

    if (($ReportRows.PowerEdge | Sort-Object -Unique) -imatch 'Precision') {
        $DownloadsDellComUrl = @()
        foreach ($PN in ($ReportRows.PowerEdge | Sort-Object -Unique)) {
            if ($PN.Length -gt 2) {
                $PrecisionNumbers = $PN -replace 'Precision' -replace 'Rack' -replace ' ' -replace 'r' -replace '7910','precision-r7910-workstation' -replace '7920','precision-7920r-workstation'
                $DownloadsDellComUrl += "<a href='https://www.dell.com/support/home/en-us/product-support/product/$PrecisionNumbers/drivers' target='_blank'>https://www.dell.com/support/home/en-us/product-support/product/$PrecisionNumbers/drivers</a>"
            }
        }
    }
    else {
        $DownloadsDellComUrl = "<a href='http://dl.dell.com/published/pages/poweredge-$ServerType.html' target='_blank'>http://dl.dell.com/published/pages/poweredge-$ServerType.html</a>"
    }

    if ($NoneSupportedDevices) {
        $NoneSupportedDevices = $NoneSupportedDevices -replace ',', '<br>'
        $Footer = '<font color="red">The following device(s) are NOT listed as supported in the CATALOG.XML for this server type: <br>'
        $Footer += $NoneSupportedDevices + '</font><br>'
        $Footer += 'More Driver and FW may be found here: <br>'
        $Footer += $DownloadsDellComUrl
        if ($S2DCatalogNeeded -eq 'YES') { $Footer += '***Storage Spaces Direct Ready Node(s) Found. Special S2D catalog used to determine certified drivers and firmware compliance.<br>' }
        $Footer += $CatVerInfo
    }
    else {
        $Footer = 'NOTES: <br>'
        if ($S2DCatalogNeeded -eq 'YES') { $Footer += '***Storage Spaces Direct Ready Node(s) Found. Special S2D catalog used to determine certified drivers and firmware compliance.<br>' }
        $Footer += 'More Driver and FW information can be found here: <br>'
        $Footer += $DownloadsDellComUrl
        $Footer += "<br><a href='https://solutions.one.dell.com/sites/NAEnterprise/SST/Communities/DRiFT/_layouts/15/start.aspx#/Lists/DriFT%20Feedback/Default.aspx' target='_blank'>Got Feedback?</a>"
    }

    return $Footer
}

function New-DriFTHtmlReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [AllowEmptyCollection()]
        [object[]]$ReportRows,
        [string]$DriFTVersion,
        [string]$DisplayVersion,
        [datetime]$ReportDate = (Get-Date),
        [string]$ServerType,
        [string]$S2DCatalogNeeded,
        [AllowNull()][object]$NoneSupportedDevices,
        [AllowNull()][string]$CatVerInfo
    )

    $Header = Get-DriFTHtmlHeader
    $OutTitle = @()
    $OutTitle += 'DriFT v' + $DisplayVersion
    $OutTitle += '<br>Date/Time: ' + $ReportDate
    $OutTitle += '<br>*A <a style="background-color:Red;color:White;">red</a> InstalledVersion indicates the InstalledVersion is less than the AvailableVersion.'
    $OutTitle += '<br>**A <a style="background-color:Yellow;">Not Available</a> InstalledVersion indicates the InstalledVersion was NOT contained in the Support Assist Collection so the latest version is shown.'

    if ($null -eq $ReportRows -or $ReportRows.Count -eq 0) {
        $ReportRows = @([PSCustomObject]@{
            ServiceTag = ''
            PowerEdge = $ServerType
            OS = ''
            Type = 'INFO'
            Category = 'No Report Rows'
            Name = 'No firmware, driver, OS, or config rows were generated for this TSR.'
            InstalledVersion = 'Not Available'
            AvailableVersion = 'Not Available'
            CatalogInfo = 'DriFT'
            Criticality = 'Informational'
            ReleaseDate = ''
            URL = ''
            Details = ''
        })
    }

    $Footer = New-DriFTReportFooter -ReportRows $ReportRows -ServerType $ServerType -S2DCatalogNeeded $S2DCatalogNeeded -NoneSupportedDevices $NoneSupportedDevices -CatVerInfo $CatVerInfo

    if (($ReportRows.ServiceTag | Sort-Object -Unique).Count -lt 2) {
        $Html = ConvertTo-DriFTReportTableRows -ReportRows $ReportRows | ConvertTo-Html -Head $Header -PreContent $OutTitle -PostContent $Footer
    }
    else {
        $CompareView = New-DriFTCompareView -ReportRows $ReportRows
        $Html = $CompareView |
            Where-Object { $_.Type -notmatch '@{ServiceTag=' } |
            Sort-Object Type,Name |
            Select-Object -Property * -ExcludeProperty RowError, RowState, Table, ItemArray, HasErrors |
            ConvertTo-Html -Head $Header -PreContent $OutTitle -PostContent $Footer
    }

    $Html -replace '&gt;','>' -replace '&lt;','<' -replace '&#39;',"'" `
        -replace '<td>INSTALLED</td>','<td style="background-color: #00ff00">INSTALLED</td>' `
        -replace '<td>MISSING</td>','<td style="color: #ffffff; background-color: #ff0000">MISSING</td>' `
        -replace '<title">hTML TABLE</title>' ,'<title"></title>' `
        -replace '<tr><th>STATUS</th><th>KB Number</th><th>LINK</th></tr>','<tr style="color: #ffffff; background-color: #0000ff"><th>STATUS</th><th>KB Number</th><th>LINK</th></tr>' `
        -replace 'td">hy','td>hy' `
        -replace [Regex]::Escape('<td>***'),'<td style="color: #ffffff; background-color: #ff0000">' `
        -replace '<td>NA</td>','<td style="background-color: #ffff00">Not Available</td>' `
        -replace '<td>Not Applicable</td>','<td style="background-color: #ffff00">Not Available</td>'
}

function Write-DriFTReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [AllowEmptyCollection()]
        [object[]]$ReportRows,
        [Parameter(Mandatory=$true)]
        [string]$Path,
        [string]$DriFTVersion,
        [string]$DisplayVersion,
        [datetime]$ReportDate = (Get-Date),
        [string]$ServerType,
        [string]$S2DCatalogNeeded,
        [AllowNull()][object]$NoneSupportedDevices,
        [AllowNull()][string]$CatVerInfo
    )

    $Html = New-DriFTHtmlReport -ReportRows $ReportRows -DriFTVersion $DriFTVersion -DisplayVersion $DisplayVersion -ReportDate $ReportDate -ServerType $ServerType -S2DCatalogNeeded $S2DCatalogNeeded -NoneSupportedDevices $NoneSupportedDevices -CatVerInfo $CatVerInfo
    if (Test-Path $Path) { Remove-Item $Path -Force }
    Out-File -FilePath $Path -InputObject $Html
    return $Path
}
#endregion DriFT normalized report engine

#import the XML
Write-host "Importing Catalog.xml...."
$CatalogXMLData = [xml](Get-Content -Path "$ExtracLoc\Catalog.xml" -Raw)
Write-host "Filtering Catalog.xml for latest PowerEdge Firmware and Drivers...."
$allArray=@()
$Files2Download=@()
$IsNewS2DCatalog="YES" #Do not change this to No Jim. :)
$SwPort2HostMapAll=@()


Do{

$NoneSupportedDevices=@()
$OutputType="HTML"
#Support Assist Data Input
IF(-not($TSRInputFiles)){
    If (-not($args)){
        $OutputType=$OutputType.ToUpper()
        Write-Host ""
        Write-Host "Please provide Support Assist Collection file from the iDRAC."
        Write-Host "    Steps to export a Support Assist Collection file:"
        Write-Host "    1. Logon to iDRAC."
        Write-Host "    2. Click on the Maintenance tab."
        Write-Host "    3. Click on the SupportAssist tab."
        Write-Host "    4. Click the Start a Collection button."
        Write-Host "    5. Click the Collect button."
        Write-Host "    6. Once completed click OK to download."
        Write-Host "    7. This is the Support Assist Collection file needed."
        Write-Host ""
        $TSRLoc=Get-FileName($env:USERPROFILE)
        if(!$TSRLoc){
            $OutputType="No"
            EndScript}
    }Else{
        IF(-not($TSRLoc)){
            $TSRLoc=$args
        }
    }
}

#Extraction temp location
# Clean workspace under %TEMP%\DriFT for each run.
$DriFTTempRoot = Join-Path $env:TEMP "DriFT"
$DriFTExtractRoot = Join-Path $DriFTTempRoot "Extract"
$DriFTRedfishRoot = Join-Path $DriFTTempRoot "Redfish"
$DriFTCatalogRoot = Join-Path $DriFTTempRoot "Catalog"
$DriFTWorkRoot = Join-Path $DriFTTempRoot "Work"

foreach ($DriFTPath in @($DriFTTempRoot,$DriFTExtractRoot,$DriFTRedfishRoot,$DriFTCatalogRoot,$DriFTWorkRoot)) {
    if (-not (Test-Path $DriFTPath -PathType Container)) {
        New-Item -ItemType Directory -Force -Path $DriFTPath | Out-Null
    }
}

$DriFTRunRoot = Join-Path $DriFTExtractRoot ("Run_" + ([guid]::NewGuid().Guid.Substring(0,8)))
$ExtracLoc = $DriFTRunRoot
if (!(Test-Path $ExtracLoc -PathType Container)) {New-Item -ItemType Directory -Force -Path $ExtracLoc | Out-Null }

#TSR unzip files
Write-Host "Unzipping TSR data files...."
function Expand-ZIPFile{
    param($file, $destination)
    $shell = new-object -com shell.application
    $zip = $shell.NameSpace($file)
    foreach($item in $zip.items())
    {
    #Removed ,1564 to allow for DSet password prompt
    #$shell.Namespace($destination).copyhere($item,1564)
    $shell.Namespace($destination).copyhere($item)
    Write-Host "$($item.path) extracted"
    "$($item.path)"
    }
}
$TFile=@()
$InnerZIP=@()
ForEach($TFile in $TSRLoc){
    $ExtFolderName=$TFile.Split('\')[-1]
    $ExtFolderName=$ExtFolderName.split('.')[0]
    $TSRDataFolder=$ExtracLoc+"\"+$ExtFolderName
    New-Item -ItemType Directory -Force -Path $TSRDataFolder | Out-Null
    Expand-ZIPFile $TFile $TSRDataFolder
    $InnerZIP=get-childitem $TSRDataFolder -filter '*.zip' -Exclude '*thermal*','*dumplog*' -Recurse
    If($InnerZIP.name){
        #$UnZipInner=$TSRDataFolder+"\"+$InnerZIP.Name
        Expand-ZIPFile $InnerZIP.fullname $TSRDataFolder
        $InnerInnerZIP=get-childitem $TSRDataFolder -filter '*.zip'
        IF($InnerInnerZIP.name -imatch $ServiceTag){
            IF($InnerInnerZIP.name -ne $InnerZIP.name){
                Expand-ZIPFile $InnerInnerZIP.fullname $TSRDataFolder
            }
        }
    }
}


$E=@()
$DriFTFolders=@()
$DriFTFolders=Get-ChildItem $ExtracLoc | Where-Object{ $_.PSIsContainer } | sort-object name

$allArrayout=@()
$MBSelLogWarnERR=@()

Foreach($E in $DriFTFolders.PSPath){
    #Support Assist Enterprise Collection from iDRAC or TSR
    $SupportAssistDataType=""
    If ($TSRDataInventory=Get-ChildItem -Path $E -Filter inventory -Directory -Recurse -Force | ForEach-Object{ $_.fullname }){
        Write-Host "SupportAssist Collection Found..."
        $SupportAssistDataType="TSR"
        #Importing TSR Data
        Write-host "Importing TSR data...."
        $CIM_BIOSAttribute=@()
        $CIM_BIOSAttribute_Instances=@()
        $DCIM_View=@()
        $DCIM_View_Instances=@()
        $DCIM_SoftwareIdentity=@()
        $DCIM_SoftwareIdentity_NAMEDINSTANCE=@()
        $TSRMetadata = $null
        $RedfishWalkData = $null
        $Is17GTSR = $false

        $MetadataPath = Get-ChildItem -Path $E -Filter 'metadata.json' -File -Recurse -Force | Select-Object -First 1 | ForEach-Object { $_.FullName }
        if ($MetadataPath) {
            try {
                $TSRMetadata = Get-Content -Raw -Path $MetadataPath | ConvertFrom-Json
                Write-Host "    Found TSR metadata.json"
            }
            catch {
                Write-Host "    WARNING: Failed to parse metadata.json: $($_.Exception.Message)" -ForegroundColor Yellow
            }
        }

        if (Test-Path $TSRDataInventory"\sysinfo_CIM_BIOSAttribute.xml" -PathType Leaf){
            $CIM_BIOSAttribute=[xml](Get-Content -Path ($TSRDataInventory + "\sysinfo_CIM_BIOSAttribute.xml") -Raw)
            $CIM_BIOSAttribute_Instances=$CIM_BIOSAttribute.CIM.MESSAGE.SIMPLEREQ."VALUE.NAMEDINSTANCE".INSTANCE
            
        }Else{
            $CIM_BIOSAttribute_Instances="MISSING"
        }
        if (Test-Path $TSRDataInventory"\sysinfo_DCIM_View.xml" -PathType Leaf){
            $DCIM_View=[xml](Get-Content -Path ($TSRDataInventory + "\sysinfo_DCIM_View.xml") -Raw)
            $DCIM_View_Instances=$DCIM_View.CIM.MESSAGE.SIMPLEREQ."VALUE.NAMEDINSTANCE".INSTANCE
            $DCIM_VIEM_Properties=@()
            foreach ($Object in @($DCIM_View_Instances)) {
                $PSObject = New-Object PSObject
                foreach ($Property in @($Object.Property)) {
                    $PSObject | Add-Member NoteProperty $Property.Name $Property.InnerText
                }
            $DCIM_VIEM_Properties+=$PSObject
            }
        }Else{
            $DCIM_View_Instances="MISSING"
            $DCIM_VIEM_Properties=@()
        }
        if (Test-Path $TSRDataInventory"\sysinfo_DCIM_SoftwareIdentity.xml" -PathType Leaf){
            $DCIM_SoftwareIdentity=[xml](Get-Content -Path ($TSRDataInventory + "\sysinfo_DCIM_SoftwareIdentity.xml") -Raw)
            $DCIM_SoftwareIdentity_NAMEDINSTANCE=$DCIM_SoftwareIdentity.CIM.MESSAGE.SIMPLEREQ."VALUE.NAMEDINSTANCE"
            $DCIM_SoftwareIdentity_Properties=@()
            foreach ($Object in @($DCIM_SoftwareIdentity_NAMEDINSTANCE.INSTANCE)) {
                $PSObject = New-Object PSObject
                foreach ($Property in @($Object.Property)) {
                    $PSObject | Add-Member NoteProperty $Property.Name $Property.InnerText
                }
            $DCIM_SoftwareIdentity_Properties+=$PSObject
            }
        }Else{
            $DCIM_SoftwareIdentity_NAMEDINSTANCE="MISSING"
        }

        if (($CIM_BIOSAttribute_Instances -eq "MISSING") -or ($DCIM_View_Instances -eq "MISSING") -or ($DCIM_SoftwareIdentity_NAMEDINSTANCE -eq "MISSING")) {
            $ViewerHtmlPath = Get-ChildItem -Path $E -Filter 'viewer.html' -File -Recurse -Force -ErrorAction SilentlyContinue | Select-Object -First 1 | ForEach-Object { $_.FullName }
            $RedfishWalkPath = Get-ChildItem -Path $E -Filter 'redfishidracwalk.tar.gz' -File -Recurse -Force -ErrorAction SilentlyContinue | Select-Object -First 1 | ForEach-Object { $_.FullName }

            if ($ViewerHtmlPath) {
                Write-Host "    Found 17G viewer.html. Parsing embedded normalized Redfish data..."
                try {
                    $ViewerWorkFolder = Split-Path -Path $ViewerHtmlPath -Parent
                    $RedfishWalkData = Import-DriFT17GViewerHtml -ViewerHtmlPath $ViewerHtmlPath -WorkFolder $ViewerWorkFolder
                    $Is17GTSR = $true
                    Write-Host "    17G viewer.html firmware inventory entries found: $(@($RedfishWalkData.Firmware).Count)"
                }
                catch {
                    Write-Host "    WARNING: Failed to parse 17G viewer.html: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }

            if ((-not $RedfishWalkData) -and $RedfishWalkPath) {
                Write-Host "    Found 17G Redfish inventory walk. Parsing redfishidracwalk.tar.gz..."
                try {
                    $RedfishWorkFolder = Split-Path -Path $RedfishWalkPath -Parent
                    $RedfishWalkData = Import-DriFT17GRedfishWalk -TarGzPath $RedfishWalkPath -WorkFolder $RedfishWorkFolder
                    $Is17GTSR = $true
                    Write-Host "    17G firmware inventory entries found: $(@($RedfishWalkData.Firmware).Count)"
                }
                catch {
                    Write-Host "    WARNING: Failed to parse 17G Redfish walk: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
        }

        if ($DCIM_SoftwareIdentity_NAMEDINSTANCE -ne "MISSING") {
            # Installed hardware from legacy SoftwareIdentity
            Write-Host "Discovering Installed Hardware..."
            $DCIM_SoftwareIdentity_NAMEDINSTANCE_INSTANCENAME_KEYBINDING_KEYVALUE_Installed = $DCIM_SoftwareIdentity_NAMEDINSTANCE |`
            Where-Object{$_.INSTANCENAME.KEYBINDING.KEYVALUE."#text" -Match "DCIM:INSTALLED"} 
            #Converting to custom property to make it easier to manage installed hardware
            $TSRSystemInfo=@()
            ForEach($Prop in $DCIM_SoftwareIdentity_NAMEDINSTANCE_INSTANCENAME_KEYBINDING_KEYVALUE_Installed.INSTANCE ){
                $TSRSystemInfo+=[PSCustomObject]@{
                    ComponentType=($Prop.Property | Where-Object{$_.Name -eq 'ComponentType'} | Select-Object Value).value
                    ComponentID=($Prop.Property | Where-Object{$_.Name -eq 'ComponentID'} | Select-Object Value).value
                    VendorID=($Prop.Property | Where-Object{$_.Name -eq 'VendorID'} | Select-Object Value).value
                    DeviceID=($Prop.Property | Where-Object{$_.Name -eq 'DeviceID'} | Select-Object Value).value
                    SubDeviceID=($Prop.Property | Where-Object{$_.Name -eq 'SubDeviceID'} | Select-Object Value).value
                    SubVendorID=($Prop.Property | Where-Object{$_.Name -eq 'SubVendorID'} | Select-Object Value).value
                    Version=($Prop.Property | Where-Object{$_.Name -eq 'VersionString'} | Select-Object Value).value
                    Display=($Prop.Property | Where-Object{$_.Name -eq 'FQDD'} | Select-Object Value).value
                    ElementName=($Prop.PROPERTY | Where-Object{$_.Name -eq 'ElementName'}|Select-Object Value).Value
                }
            }
            #Filtering for only unique hardware
            Write-Host "Removing Duplicate Discovered Hardware..."
            $InstalledHardwareUnique=$TSRSystemInfo|Group-Object 'ComponentID','VendorID','DeviceID','SubDeviceID','SubVendorID'|`
              ForEach-Object{$_.Group|Select-Object 'componentType','componentID','vendorID','deviceID','subDeviceID','subVendorID','version','display','ElementName' -First 1}|`
              Where-Object{($_.DeviceID.length -ge 1)-or($_.ComponentID.length -ge 1)}|`
              sort-object 'componentType','componentID','vendorID','deviceID','subDeviceID','subVendorID','version','display'
        }
        elseif ($RedfishWalkData -and @($RedfishWalkData.Firmware).Count -gt 0) {
            Write-Host "Building installed hardware inventory from 17G Redfish firmware inventory..."
            $DCIM_SoftwareIdentity_Properties = @($RedfishWalkData.Firmware)
            $InstalledHardwareUnique = @($RedfishWalkData.Firmware)
            Write-Host "    17G InstalledHardwareUnique rows: $(@($InstalledHardwareUnique).Count)"
        }
        else {
            Write-host "    WARNING: SoftwareIdentity.xml missing and no 17G Redfish fallback inventory was found. Please upgrade to the latest iDRAC version to see the rest of the hardware." -foregroundcolor Yellow
            $InstalledHardwareUnique=@()
        }
        # Azure Stack Hub
           
            # All possible PM modules 
            [hashtable]$AZHUBElementNames = @{}
            $AZHUBElementNames.Add('3X9R5','R640')
            $AZHUBElementNames.Add('R4GYM','R840')
            $AZHUBElementNames.Add('MGFK5','R640')
            $AZHUBElementNames.Add('V4F4H','R640')
            $AZHUBElementNames.Add('V8NYX','R640')
            $AZHUBElementNames.Add('50CRT','R740XD')
            #$AZHUBElementNames
            
            # PM Module from TSR                            
            $IDModule = $DCIM_SoftwareIdentity_Properties | Where-Object{$_.ElementName -imatch 'Identity Module'} | Select-Object -First 1 ElementName
            $IsAZHub=$False
            if ($IDModule -and $IDModule.ElementName -match '\(([^\)]+)\)') {
                IF($AZHUBElementNames.keys -imatch $Matches[1]){
                    $IsAZHub=$True
                    Write-Host "    Found Azure Stack Hub: $IsAZHub"
                }
            }
            
        # Installed hardware inventory has already been built above. Do not rebuild it here;
        # rebuilding here would overwrite 17G Redfish inventory with an empty legacy DCIM result.
    }
    #Support Assist Enterprise Collection XML
    IF($SupportAssistDataType -lt 3){
    
        If($SAEDataInventory=Get-ChildItem -Path $DriFTFolders.fullname -Include "MaserInfo.xml","Inventory.xml" -File -Recurse -Force | Select-Object -last 1  | ForEach-Object{ $_.Directory } ){
            $InvPath=""
            $MasPath=""
            $SupportAssistDataType="SAEX"
            # Server Type and Service Tag
            $chasinfoxml=Get-ChildItem -Path $DriFTFolders.fullname -Include "chasinfo.xml" -File -Recurse -Force | sort-object Length | Select-Object -last 1 | ForEach-Object{ $_.Directory } 
            $SvrInfo=[xml](Get-Content -Path ($chasinfoxml + "\chasinfo.xml") -Raw)
            #Firmware inventory
            $InvPath=$SAEDataInventory.FullName+"\Inventory.xml"
                IF([System.IO.File]::Exists($InvPath)){$inv=[xml](Get-Content -Path $InvPath -Raw)
                $XMLLIB="SVMInventory"}
            #Firmware Maser
            $MasPath=$SAEDataInventory.FullName+"\MaserInfo.xml"
                IF([System.IO.File]::Exists($MasPath)){
                $inv=[xml](Get-Content -Path $MasPath -Raw)
                $XMLLIB="OMA"}
            $SystemID=$inv.$XMLLIB.System
            $SystemFW=$inv.$XMLLIB.Device | Select-Object `
                @{Label="componentType";Expression={$_.Application.componentType}},@{Label="componentID";Expression={$_.componentID}},`
                @{Label="vendorID";Expression={$_.vendorID}},@{Label="deviceID";Expression={$_.deviceID}},`
                @{Label="subDeviceID";Expression={$_.subDeviceID}},@{Label="subVendorID";Expression={$_.subVendorID}},`
                @{Label="version";Expression={$_.Application.version}},@{Label="display";Expression={$_.display}},`
                @{Label="application";Expression={$_.application.display}}

            #Filter for unique hardware       
              $InstalledHardwareUnique=$SystemFW|Group-Object 'componentID','vendorID','deviceID','subDeviceID','subVendorID','version'|`
                ForEach-Object{$_.Group|Select-Object 'componentType','componentID','vendorID','deviceID','subDeviceID','subVendorID','version','display','ElementName' -First 1}|`
                Where-Object{($_.DeviceID.length -ge 1)-or($_.ComponentID.length -ge 1)}|`
                sort-object 'componentType','componentID','vendorID','deviceID','subDeviceID','subVendorID','version','display' 
        }
    }
    #Support Assist Enterprise Collection JSON
    IF($SupportAssistDataType -lt 3){
        If($SAEDataInventory=Get-ChildItem -Path $DriFTFolders.fullname -Filter 'supportassist_output.json' | ForEach-Object{ $_.fullname } ){
            $SupportAssistDataType="SAEJ"
            #Support Assist Enterprise Collection JSON in the RAW
            IF($SAEJraw=(Get-Content -raw -path $SAEDataInventory | ConvertFrom-Json)){
                Write-Host "SupportAssist Collection JSON Loaded..."
                Write-Host "    ERROR: Invalid input file detected. Exiting." -foregroundcolor Red 
            }
            $OutputType="No"
            EndScript
        }
    }

   IF ($SupportAssistDataType.Length -lt 3){
        Write-Host
        Write-Host "    ERROR: Invalid input file detected. Exiting." -foregroundcolor Red 
        Write-Host
        $OutputType="No"
        EndScript
   }

    #Check for Server Model number
        Write-host "Finding server model in TSR data...."
        $ServiceTag=@()
        $ServerType=@()
        If($SupportAssistDataType -eq "TSR"){
            if (($CIM_BIOSAttribute_Instances -ne "MISSING") -and ($DCIM_View_Instances -ne "MISSING")) {
                $HostName=(($CIM_BIOSAttribute_Instances|`
                Where-Object{($_.CLASSNAME -match "DCIM_SystemString")}).PROPERTY|`
                Where-Object{$_.VALUE -eq "HostName"}).ParentNode.'PROPERTY.ARRAY' |`
                Where-Object{$_.Name -match "CurrentValue"}|`
                          Select-Object @{Label="CurrentValue";Expression={$_.'VALUE.ARRAY'.VALUE}}
                $HostName=$HostName.CurrentValue
                $ServiceTag=($DCIM_View_Instances| Where-object {($_.CLASSNAME -match "DCIM_SystemView")}).PROPERTY | Where-Object {$_.NAME -eq "ServiceTag"} | Select-Object @{Label="ServiceTag";Expression={$_.Value}} | Select-Object -First 1
                $ServerType=($DCIM_View_Instances| Where-Object {($_.CLASSNAME -eq "DCIM_SystemView")}).PROPERTY | Where-Object {$_.NAME -eq "MODEL"} | Select-Object @{Label="Model";Expression={$_.Value}}| Select-Object -First 1
                $SystemID=(($CIM_BIOSAttribute_Instances| Where-Object {($_.CLASSNAME -eq "DCIM_LCString")}|Where-Object{$_.PROPERTY.VALUE -eq 'SYSID'}).'PROPERTY.ARRAY' | Where-Object {$_.NAME -eq "CurrentValue"}).'VALUE.ARRAY'.VALUE
                $ServerType=$ServerType.Model
            }
            else {
                Write-Host "    Using 17G metadata/Redfish data for system identity..."
                $HostName = Get-DriFTFirstNonEmpty $TSRMetadata.HostName $RedfishWalkData.System.HostName $RedfishWalkData.System.Name
                $ServiceTagValue = Get-DriFTFirstNonEmpty $TSRMetadata.ServiceTag $RedfishWalkData.System.SerialNumber $RedfishWalkData.System.SKU
                $ServiceTag = [PSCustomObject]@{ ServiceTag = $ServiceTagValue }
                $ServerType = Get-DriFTFirstNonEmpty $TSRMetadata.Model $RedfishWalkData.System.Model
                $SystemID = Get-DriFTFirstNonEmpty $TSRMetadata.DeviceSystemId $TSRMetadata.SystemID $RedfishWalkData.System.Oem.Dell.DellSystem.SystemID
            }
                
########### Change this to YES to force ASHCI-catalog.xml
            $S2DCatalogNeeded="No"

            switch ($ServerType){
                
                # Added for Precision rack systems
                {$PSItem -match 'Precision'}{
                    IF($ServerType.Length -gt 4){
                        Write-Host "    Found: Server Type Precision Rack."
                        $SpecialCatalogNeeded="Precision"
                        $ServiceTag=$ServiceTag.ServiceTag
                        }
                    }
    
                #Added for XR2 same as R440
                {$PSItem -match 'XR2'}{
                    IF($ServerType.Length -gt 4){
                        Write-Host "    Found: Server Type XR2. Changing to R440."
                        #$ServerType=$ServerType -replace "XR2","R440"
                        $ServerType="R440"
                        $S2DCatalogNeeded="NO"
                        $SpecialCatalogNeeded="NO"
                        $ServiceTag=$ServiceTag.ServiceTag
                        }
                    }
                {$PSItem -match 'AX'}{
                    IF($ServerType.Length -gt 4){
                        Write-Host "    Found: Server Type of AX"
                        $ServerType=$ServerType -replace "AX-","R"
                        $S2DCatalogNeeded="YES"
                        $SpecialCatalogNeeded="HCI"
                        $ServiceTag=$ServiceTag.ServiceTag+"***"
                        $ServiceTagList+=$ServiceTag+"_"
                        }
                    }
                {$PSItem -match 'Storage Spaces Direct'}{
                    IF($ServerType.Length -gt 4){
                        Write-Host "    Found: Server Type of Storage Spaces Direct Ready Node"
                        $ServerType=$ServerType -replace " Storage Spaces Direct RN","" -replace " Storage Spaces Direct R",""
                        
                        $S2DCatalogNeeded="YES"
                        $SpecialCatalogNeeded="HCI"
                        $ServiceTag=$ServiceTag.ServiceTag+"***"
                        $ServiceTagList+=$ServiceTag+"_"
                        }
                    }
                
                #Added for vSAN Ready Nodes
                {$PSItem -match 'vSAN'}{
                    IF($ServerType.Length -gt 4){
                        Write-Host "    Found: vSAN Ready Node"
                        $IsvSAN=$True
                        Write-Host "    ERROR: vSAN is not supported yet. Try again later." -ForegroundColor Red
                        EndScript
                        }
                    }

                #Added for ScaleIO Ready Nodes
                {$PSItem -match 'ScaleIO'}{
                    IF($ServerType.Length -gt 4){
                        Write-Host "    Found: ScaleIO Ready Node"
                        $IsvSAN=$True
                        Write-Host "    ERROR: ScaleIO is not supported yet. Try again later." -ForegroundColor Red
                        EndScript
                        }
                    }

                #everything else
                default{
                IF($ServerType.Length -gt 4){
                    #Added to pull the server model out Ex. PowerEdge R740XD = R740XD
                    $ServerType=($ServerType -split "\W")[1]
                    }
                    #Added for XC6320 servers 
                    If(($ServerType -like "XC*") -and ((([regex]::match($ServerType,"\d+").Groups[0].Value).Trim()).length -eq 4))`
                    {$ServerType=$ServerType -replace "XC","C"}
                    Else{
                    #Added for XC servers 
                    $ServerType=$ServerType -replace "XC","R"}# -replace "xd",""}
                    $ServiceTag=$ServiceTag.ServiceTag
                    $ServiceTagList+=$ServiceTag+"_"
                    #Added for R320 Servers
                    If($ServerType -eq "R320"){$ServerType=$ServerType+'/NX400'}
                 }
            }

            #No Server Type Found

            If(($ServerType.Length -lt 4)-and ($ServiceTag.length -gt 0)){
                # Retrieve server type from support.dell.com with Service Tag
                Write-Host "    WARNING: Failed to find Server Model in TSR data..." -foregroundcolor Yellow
                Write-Host "             Trying to retrieving Server Model from support.dell.com with Service Tag $ServiceTag..." -foregroundcolor Yellow
                $URL="http://www.dell.com/support/home/us/en/19/product-support/servicetag/$ServiceTag"
                $result = Invoke-webrequest -Uri $URL -Method Get
                IF($result.StatusCode -match 200){
                    $resultTable = @{}
                    # Get the title
                    $resultTable.title = $result.ParsedHtml.title
                    If ($resultTable.title -match 'OEMR'){
                        Write-Host "    ERROR: None Supported System Detected: OEMR" -foregroundcolor Red
                        Write-Host "           No Output will be generated..." -foregroundcolor Red
                        $OutputType="No"
                        EndScript
                    }
                    $ServerType=($resultTable.title -replace "Support for ","").split("|")[0]
                    IF($ServerType -match 'Storage Spaces Direct'){
                        IF($ServerType.Length -gt 4){
                            Write-Host "    Found: Server Type of Storage Spaces Direct Ready Node"
                            $ServerType=$ServerType -replace "Storage Spaces Direct ","" -replace " Ready Node",""
                            $ServerType=$ServerType.Trim()
                            $S2DCatalogNeeded="YES"
                            $SpecialCatalogNeeded="HCI"
                            $ServiceTag=$ServiceTag.ServiceTag+"***"
                            $ServiceTagList+=$ServiceTag+"_"
                        }
                    }Else{$ServerType=([regex]::match($ServerType,"\D[A-Z]\d+").Groups[0].Value).Trim()}
                    Write-Host "    Success: Server Model $ServerType found by Service Tag on support.dell.com..." -foregroundcolor Green
                }Else{
                    Write-Host "    ERROR: Service Tag not found on support.dell.com..." -foregroundcolor Red
                }
       
            }
            #Service tag Not found on support.dell.com
            If($ServerType.Length -lt 4){
                #Added to handle missing Server Model information
                Write-Host "    WARNING: Server Model $ServerType not expected. The expected value should be like R740." -foregroundcolor Yellow
                $MOServerType = Read-Host "Would you like to manually enter the Server Model? [y/n]"
                If (($MOServerType -ieq "n")-or ($MOServerType -ieq "")){
                $OutputType="No"
                EndScript}
                Write-Host "Please type the Server Model and press Enter. "
                $ServerTypeOverride=Read-Host "    Example: R740"
                If (($ServerTypeOverride.Length -lt 4) -or ($ServerTypeOverride -ieq "")){
                Write-Host "    ERROR: Server Model you entered was not in the proper format. Please run again." -foregroundcolor Red
                $OutputType="No"
                EndScript}
                    $ServerTypeOverride1 = @{
                    Model=$ServerTypeOverride
                    }
                    $ServerType = New-Object PSObject -Property $ServerTypeOverride1
                    $ServerType = $ServerType.Model
                    
                    # Enable to force R740XD ASHCI Catalog
                    #$ServerType="$ServerType Storage Spaces Direct RN"
                    
                    $ServiceTag=$ServiceTag.ServiceTag
                    $ServiceTagList+=$ServiceTag+"_"
            }
        }


        If($SupportAssistDataType -eq "SAEX"){
            $ServiceTag=$SvrInfo.OMA.ChassisList.Chassis.ChassisInfo.ChassisProps2.ServiceTag
            $SystemId=$inv.SVMInventory.system.systemid
            $S2DCatalogNeeded="NO"
            $SpecialCatalogNeeded="NO"
            $ServerType=$SvrInfo.OMA.ChassisList.Chassis.ChassisInfo.ChassisProps1.ChassModel
            If(($SystemId.length -lt 4)-and($ServerType.length -lt 4)){
                Write-Host "    ERROR: Server type is missing in TSR data. No data to output...." -foregroundcolor red
                $OutputType="NO"
                EndScript
            }
            #Added for XC servers 
            #$ServerType=$ServerType -replace "XC","R" #-replace "xd",""
            $ServiceTagList+=$ServiceTag+"_"
            switch ($ServerType){
                #Added for Storage Spaces Direct servers
                {$PSItem -match 'Storage Spaces Direct'}{
                    IF($ServerType.Length -gt 4){
                        Write-Host "    Found: Server Type of Storage Spaces Direct Ready Node"
                        $ServerType=$ServerType -replace " Storage Spaces Direct RN","" -replace " Storage Spaces Direct R",""
                        
                        $S2DCatalogNeeded="YES"
                        $SpecialCatalogNeeded="HCI"
                        $ServiceTag=$ServiceTag+"***"
                        $ServiceTagList+=$ServiceTag+"_"
                        }
                    }

                #everything else
                default{
                IF($ServerType.Length -gt 4){
                    #Added to pull the server model out Ex. PowerEdge R740XD = R740XD
                    $ServerType=($ServerType -split "\W")[1]
                    }
                    #Added for XC6320 servers 
                    If(($ServerType -like "XC*") -and ((([regex]::match($ServerType,"\d+").Groups[0].Value).Trim()).length -eq 4))`
                    {$ServerType=$ServerType -replace "XC","C"}
                    Else{
                    #Added for XC servers 
                    $ServerType=$ServerType -replace "XC","R"}# -replace "xd",""}
                    $ServiceTag=$ServiceTag
                    $ServiceTagList+=$ServiceTag+"_"
                    #Added for R320 Servers
                    If($ServerType -eq "R320"){$ServerType=$ServerType+'/NX400'}
                 }
            }
            
        }
        If($SupportAssistDataType -eq "SAEJ"){
            $ServiceTag=($SAEJraw.objects | Where-Object{$_.objectId -match 'BIOS_Setup_1_1_SystemServiceTag'}).fields.Value
            $ServerType=([regex]::match(($SAEJraw.objects | Where-Object{$_.objectId -match 'BIOS_Setup_1_1_SystemModelName'}).fields.Value,"\D[A-Z]\d\d\d").Groups[0].Value).Trim()
            
            If($ServerType.Length -lt 4){
                Write-Host "    ERROR: Server type is missing in input data. No data to output...." -foregroundcolor red
                $OutputType="NO"
                EndScript
            }
            #Added for XC servers 
            $ServerType=$ServerType -replace "XC","R" #-replace "xd",""
            $ServiceTagList+=$ServiceTag+"_"  
        }
        Write-host "    Found server model:" $ServerType
        Write-host "Finding Service Tag...."
        Write-host "    Found Service Tag:" $ServiceTag
        
    #Installed OS Check
    Write-host "Finding which OS is installed in TSR data...."
    #LWXP LW64 LLXP
        $OSCheck=@()
        $OSMjrVer=@()
        $OSMinVer=@()
        $OperatingSystemYear=""
        If($SupportAssistDataType -eq "TSR"){
            if ($CIM_BIOSAttribute_Instances -ne "MISSING") {
                $OSName0=$CIM_BIOSAttribute_Instances| Where-Object {($_.CLASSNAME -match "DCIM_SystemString")} | Where-Object {$_.PROPERTY.Value -Match "OSName"}
                $OSName1=$OSName0.ChildNodes | Where-Object{($_.NAME -match "CurrentValue")}
                $OSCheck=$OSName1.InnerText
            }
            else {
                $OSCheck = Get-DriFTFirstNonEmpty $TSRMetadata.OSName $TSRMetadata.OperatingSystem $RedfishWalkData.System.OperatingSystem $RedfishWalkData.System.Oem.Dell.DellOperatingSystem
            }
            $DriverSupport = $False
            $VMWOSVer=""
            Switch($OSCheck){
                {$OSCheck -imatch "Windows"}{
                    #$OSCheck="Windows Server 2016"            
                    IF ($OSCHECK -match "2008"){
                    $OperatingSystemYear = "2008"
                    $OSMjrVer=6
                    $OSMinVer=0}
                    IF (($OSCHECK -match "2008") -and ($OSCHECK -match "R2")){
                    $OperatingSystemYear = "2008 R2"
                    $OSMjrVer=6 
                    $OSMinVer=1}
                    IF ($OSCHECK -match "2012"){
                    $OperatingSystemYear = "2012"
                    $OSMjrVer=6 
                    $OSMinVer=2}
                    IF (($OSCHECK -match "2012") -and ($OSCHECK -match "R2")){
                    $OperatingSystemYear = "2012 R2"
                    $OSMjrVer=6 
                    $OSMinVer=3}
                    IF ($OSCHECK -match "2016"){
                    $OperatingSystemYear = "2016"
                    $OSMjrVer=10 
                    $OSMinVer=0
                    $Build=$NULL}
                    IF ($OSCHECK -match "2019"){
                    $OperatingSystemYear = "2019"
                    $OSMjrVer=10 
                    $OSMinVer=17763
                    $Build='17784'}
                    IF ($OSCHECK -imatch "20H2"){
                    $OperatingSystemYear = "20H2"
                    $OSMjrVer=10 
                    $OSMinVer=0
                    $Build='17784'}
                    IF ($OSCHECK -imatch "21H2"){
                    $OperatingSystemYear = "21H2"
                    $OSMjrVer=10 
                    $OSMinVer=0
                    $Build='20348'}
                    IF ($OSCHECK -imatch "22H2"){
                    $OperatingSystemYear = "22H2"
                    $OSMjrVer=10 
                    $OSMinVer=0
                    $Build='20349'}
                    IF ($OSCHECK -imatch "23H2"){
                    $OperatingSystemYear = "23H2"
                    $OSMjrVer=10 
                    $OSMinVer=0
                    $Build='25398'}
                    IF ($OSCHECK -match "2022"){
                    #$OSCHECK="2022-21H2-22H2"
                    $OperatingSystemYear = "2022"
                    $OSMjrVer=10 
                    $OSMinVer=0
                    $Build='20348'}
                    IF ($OSCHECK -match "Windows 10"){
                    $OSMjrVer=10 
                    $OSMinVer=0}
                    $DriverSupport = $True
                    $OSVersion=$Build
                }
                {$OSCheck -imatch "VMware"}{
                    # Get installed VMware Version from sysinfo_CIM_BIOSAttribute.xml
                    $IsVMware=$True
                    If($OSCheck -inotmatch "build"){
                        $VMWOSVersion0=$CIM_BIOSAttribute_Instances| Where-Object {($_.CLASSNAME -match "DCIM_SystemString")} | Where-Object {$_.PROPERTY.Value -Match "OSVersion"}
                        $VMWOSVersion1=$VMWOSVersion0.ChildNodes | Where-Object{($_.NAME -match "CurrentValue")}
                        $VMWOSVersionCheck=$VMWOSVersion1.InnerText | Sort-Object -Unique
                    }Else{$VMWOSVersionCheck=$OSCheck}
                    #($VMWOSVersionCheck -split " ")[0]
                    ForEach($V in $VMWOSVersionCheck){
                        If($v.length -gt 0){
                            $VMWOSVer=""
                            Switch ($V){
                                {$V -imatch "build"}{
                                    #$VMWOSVer=($v -replace "VMware ","" -replace "ESXi ","" -replace " Update "," U" -replace ".0 "," " -split "Build")[0]
                                    $VMWOSVer=($v -replace "VMware ","" -replace "ESXi ","" -replace " Update "," U" -split "Build")[0]
                                    $VMWOSVer=$VMWOSVer.trim()
                                    $OSVersion=""
                                    $OSVersion=$VMWOSVer
                                    $DriverSupport = $True
                                }
                                {$V -imatch "Patch"}{
                                    $VMWOSVer=($v -replace "Update "," U" -replace ".0 ","" -split " Patch")[0]
                                    $OSVersion=""
                                    $OSVersion=$VMWOSVer
                                    $DriverSupport = $True
                                }
                                {$V -imatch "GA"}{
                                    $VMWOSVer=($v -replace ".0 ","" -split "GA ")[0]
                                    $OSVersion=""
                                    $OSVersion=$VMWOSVer
                                    $DriverSupport = $True
                                }
                                {$V -imatch "7.0.0"}{
                                    $VMWOSVer=($v -split ".0 ")[0]
                                    $OSVersion=""
                                    $OSVersion=$VMWOSVer
                                    $DriverSupport = $True
                                }
                                Default{
                                    $VMWOSVer=($v  -split " ")[0]
                                    $OSVersion=""
                                    $OSVersion=$VMWOSVer
                                    $DriverSupport = $True
                                    
                                }
                        }
                    }
                }
            }
            }
            #Removes any extra spaces
                $VMWOSVer = $VMWOSVer -replace '\s{2}', ' '
            #Added to compinsate for versions like 7.0.3 U3 where the .3 after the .0 does not matter so we remove it
                IF((($VMWOSVer | Select-String -Pattern '\.' -AllMatches).Matches.Count) -gt 1){
                    $VMWOSVer = $VMWOSVer -replace '\.[0-9]\s', ' '
                }

        }
        If($SupportAssistDataType -eq "SAEX"){
            $SAE_OS_info=$inv.SVMInventory.OperatingSystem
            $OSMjrVer=$SAE_OS_info.majorVersion
            $OSMinVer=$SAE_OS_info.minorVersion
            IF($SAE_OS_info.osVendor -match "Microsoft"){
                IF (($OSMjrVer -eq 6)-and($OSMinVer -eq 0)){
                $OSCheck="Windows Server 2008"}
                IF (($OSMjrVer -eq 6)-and($OSMinVer -eq 1)){
                $OSCheck="Windows Server 2008 R2"}
                IF (($OSMjrVer -eq 6)-and($OSMinVer -eq 2)){
                $OSCheck="Windows Server 2012"}
                IF (($OSMjrVer -eq 6)-and($OSMinVer -eq 3)){
                $OSCheck="Windows Server 2012 R2"}
                IF (($OSMjrVer -eq 10)-and($OSMinVer -eq 0)){
                $OSCheck="Windows Server 2016"}
                IF (($OSMjrVer -eq 0)-and($OSMinVer -eq 0)){
                $OSCheck="Windows Server 2019"}
            }
                
        }
        If($SupportAssistDataType -eq "SAEJ"){
            $SAEJ_OSVer=@()
            $SAEJ_OSVerS=@()
            $SAE_OS_info=($SAEJraw.objects | Where-Object{$_.objectId -match 'OperatingSystem'}).fields.OSName
            $SAEJ_OSVer=($SAEJraw.objects | Where-Object{$_.objectId -match 'OperatingSystem'}).fields.Version.split() | sort-object
            $SAEJ_OSVerS=$SAEJ_OSVer[1].Split(".")
            $OSMjrVer=$SAEJ_OSVerS[0]
            $OSMinVer=$SAEJ_OSVerS[1]
            IF($SAE_OS_info -match "Microsoft"){
            $OSCheck="Windows"}
        }
        $NOOSSupport="NO"

        IF($IsVMware -ne $True){
            $OSVer = "LW64"
            IF((($OSCHECK).Length -gt 0) -and ($ServerType -match "20") -and (!($OSCHECK -match "Windows"))-and (!($SAE_OS_info.osArch -match "x64"))){$OSVer = "LWXP"}
            IF ($Null -eq $OSver){
                #Show firmware only
                Clear-Host
                Write-Host "    ERROR: NON-SUPPORTED OS DETECTED: $OSCheck. No output...." -foregroundcolor red
                $OutputType="NO"
                EndScript
            }
        }
        If(($OSCHECK).Length -eq 0){
            $InstalledOS=" NO OS Detected in TSR Data: Assuming Windows 64bit"
            $OSMjrVer=6 
            $OSMinVer=3
        }Else{$InstalledOS=($OSCheck)}

        #Added for driver check
        IF($DriverSupport -eq $False){
            $NOOSSupport="YES"
            Write-host "    ERROR: No Driver Support for"$InstalledOS ". Firmware ONLY."  -foregroundcolor red
        }
    Write-host "    Found OS:" $OSCheck $OSVersion
    $SkipDriversandFirmware = "NO" 
    # Drivers and Firmware for Precision
    If($SpecialCatalogNeeded -eq 'Precision'){
        Write-Host "Downloading Precision catalog..."
        $CatLocNA="NO"
        Try{
            $SpecialCatalogSource="https://downloads.dell.com/catalog/CatalogIndexPC.cab"
            $SpecialCatalogName=($SpecialCatalogSource -split '\/')[-1]
            $SpecialCatalogNameExt=($SpecialCatalogName -split '\.')[-1]
            $SpecialCatalogFile="$env:TEMP\DriFT\$SpecialCatalogName"
            Invoke-WebRequest -Uri $SpecialCatalogSource -OutFile "$env:TEMP\DriFT\$SpecialCatalogName" -UseDefaultCredentials
            IF(-not(Test-Path $SpecialCatalogFile)){
                $CatLocNA="YES"
            }
        }
        Catch{
            $CatLocNA="YES"
            Write-Host "    WARNING: Special Catalog $SpecialCatalogName Source location NOT accessible. Please provide file."-foregroundcolor Yellow
            Write-Host "    Or manually download from:"$SpecialCatalogSource -foregroundcolor Yellow
            }
        Finally{
            #Ask for the catalog.cab
            If($CatLocNA -eq "YES"){
                Function Get-CatFile()
                    {
                        param([string]$Title,[string]$Directory,[string]$Filter="All Files (*.*)|*.*")
                        Add-Type -AssemblyName System.Windows.Forms
                        $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog -Property @{Multiselect = $true}
                        $OpenFileDialog.Title = $Title
                        $OpenFileDialog.initialDirectory = $Directory
                        $OpenFileDialog.filter = $Filter
                        $OpenFileDialog.ShowDialog() | Out-Null
                        $OpenFileDialog.filenames
                    }
                $FExt=$SpecialCatalogNameExt
                $DownloadFile=Get-CatFile -Title "Please Select Special Catalog" -Directory "C:" -Filter "$($FExt) (*.$($FExt))| *.$($FExt)"
                If(!$DownloadFile){
                    $OutputType="No"
                    EndScript}
                }

            IF($SpecialCatalogNameExt -eq 'gz'){
                $infile = $($env:TEMP)+'\DriFT\'+$($SpecialCatalogName)
                DeGZip-File -Path $infile
            }ElseIf($SpecialCatalogNameExt -eq 'cab'-or $SpecialCatalogNameExt -eq 'zip'){
                $SpecialCatalogExpandedLocation="$env:TEMP\DriFT\"
                # Clean old xml catalogs
                Remove-item -Path "$SpecialCatalogExpandedLocation\*.XML"
                Expand-ZIPFile $($SpecialCatalogFile) $($SpecialCatalogExpandedLocation)
            }
        }# Finally
        # Pull
        Try{
            $SpecialCatalogExtractedPath=Get-ChildItem -Path $SpecialCatalogExpandedLocation -Filter "$(($SpecialCatalogName -split '\.')[0]).xml"
            $PrecisionCatalogLink=(((Get-Content $SpecialCatalogExtractedPath.FullName | Select-String -SimpleMatch "$SystemID.cab") -split 'path=')[-1] -split ' size')[0] -replace '"'
            $PrecisionCatalogLink="https://dl.dell.com/$PrecisionCatalogLink"
            $PrecisionCatalogName=($PrecisionCatalogLink -split '\/')[-1]
            $PrecisionCatalogNameExt=($PrecisionCatalogLink -split '\.')[-1]
            $SpecialCatalogFile="$env:TEMP\DriFT\$PrecisionCatalogName"
            Invoke-WebRequest -Uri $PrecisionCatalogLink -OutFile $SpecialCatalogFile -UseDefaultCredentials
            IF(-not(Test-Path $SpecialCatalogFile)){
                $CatLocNA="YES"
            }
        }Catch{}
        Finally{
            IF($PrecisionCatalogNameExt -eq 'gz'){
                DeGZip-File -Path $($env:TEMP)+'\DriFT\'
            }ElseIf($PrecisionCatalogNameExt -eq 'cab'-or $PrecisionCatalogNameExt -eq 'zip'){
                $PrecisionCatalogExpandedLocation="$env:TEMP\DriFT\"
                Expand-ZIPFile $($SpecialCatalogFile) $($PrecisionCatalogExpandedLocation)
            }
            $PrecisionCatalogExtractedPath=Get-ChildItem -Path $PrecisionCatalogExpandedLocation | Where-Object{$_.name -imatch "$($SystemID).xml"}
            IF($PrecisionCatalogExtractedPath){
            [xml]$PrecisionCatalog = Get-Content -Path $PrecisionCatalogExtractedPath.FullName -Raw
            # Filter to systemID
            $PrecisionCatalogXMLDataFiltered = $PrecisionCatalog.Manifest.SoftwareComponent.SupportedSystems.Brand.Model | Where-Object{$_.SupportedSystems.Brand.Model}
            }Else{
                Write-Host "ERROR: Failed to find SystemID ($SystemID) in CatalogIndexPC.XML. No data will be added to report." -ForegroundColor Red
                #$OutputType = "NO"
                # Skipping driver and fw
                $SkipDriversandFirmware = "YES"
                # Removing data from report so it will not show up
                $allarray=@()
                $SpecialCatalogNeeded="NO"
                #EndScript
            }

        }

    }
   
    #Drivers and Firmware for Storage Space Direct
        switch ($SpecialCatalogNeeded){
        {$PSItem -eq "HCI"}{
            If($IsNewS2DCatalog -eq "YES"){
                #Download S2D_Catalog_and_SCP.zip
                Try{
                    $S2DCatalogSource="https://downloads.dell.com/catalog/ASHCI-Catalog.xml.gz"
                    Invoke-WebRequest $S2DCatalogSource -OutFile "$env:TEMP\DriFT\ASHCI-Catalog.xml.gz" -UseDefaultCredentials
                }Catch{
                    $CatLocNA="YES"
                    Write-Host "    WARNING: S2D Catalog Source location NOT accessible. Please provide ASHCI-Catalog.xml.gz file."-foregroundcolor Yellow
                    Write-Host "    Or manually download from:"$S2DCatalogSource -foregroundcolor Yellow
                    }
                Finally{
                    #Ask for the catalog.cab
                    If($CatLocNA -eq "YES"){
                        Function Get-CatFile($initialDirectory)
                        {
                            Add-Type -AssemblyName System.Windows.Forms
                            $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog -Property @{Multiselect = $true}
                            $OpenFileDialog.Title = "Please Select S2D CATALOG.CAB file..."
                            $OpenFileDialog.initialDirectory = $initialDirectory
                            $OpenFileDialog.filter = "CAB (*.gz)| *.gz"
                            $OpenFileDialog.ShowDialog() | Out-Null
                            $OpenFileDialog.filenames
                        }
                        $DownloadFile=Get-CatFile("C:")
                        if(!$DownloadFile){
                            $OutputType="No"
                            EndScript}
                    }
                }
                

                #Extract S2D_Catalog_and_SCP.zip
                    $infile="$($env:TEMP)\DriFT\ASHCI-Catalog.xml.gz"
                    DeGZip-File $infile
                #Parse S2D Catalog
                    #import the XML
               
                    Write-host "Importing ASHCI-Catalog.xml...."
                    $S2DCatalogXMLData = [xml](Get-Content -Path "$env:TEMP\DriFT\ASHCI-Catalog.xml" -Raw)
                    #$Data = $S2DCatalogXMLData
                    #$Catalogs+= $S2DCatalogXMLData
                    Write-host "Filtering ASHCI-Catalog.xml for latest PowerEdge Firmware and Drivers...."
                    $S2DCatalogXMLDataFiltered=@()
                    $S2DCatalogXMLDataFiltered+=$S2DCatalogXMLData.Manifest.SoftwareComponent|`
                    Where-Object{($_.SupportedSystems.Brand.Model.Display."#cdata-section" -eq $ServerType)-or($_.SupportedSystems.Brand.Model.systemID -eq $SystemID)-or($_.SupportedSystems.Brand.Model.systemID -match 'VRTX')}|`
                    Where-Object{($_.packageType -eq "LW64")-or($_.packageType -eq "LWXP")}|`
                    <# added a filter to include with Null or SupportedOperatingSystems#>
                    Where-Object{($_.SupportedOperatingSystems.OperatingSystem.Display.'#cdata-section' -eq $null  -or $_.SupportedOperatingSystems.OperatingSystem.Display.'#cdata-section' -imatch $OperatingSystemYear)}
                    #pause
                    $S2DCatalogInfo=@()
                    $S2DCatalogInfo="ASHCI-Catalog.xml"
                    #$S2DCatalogInfo+="<br> "+$CatalogXMLData.Manifest.dateTime
                    #$S2DCatalogInfo+="<br> "+$CatalogXMLData.Manifest.releaseID
                    $S2DCatalogInfo+="<br> "+$CatalogXMLData.Manifest.version                  
                    $CatVerInfo+="<br> S2D_Catalog.xml Info <br>&nbsp&nbspData/Time: "+$S2DCatalogXMLData.Manifest.dateTime+"<br>&nbsp&nbspReleaseId: "+$S2DCatalogXMLData.Manifest.releaseID+"<br>&nbsp&nbspVersion: "+$S2DCatalogXMLData.Manifest.version
                    $IsNewS2DCatalog="YES"
                    IF(!$S2DCatalogXMLDataFiltered){Write-host "    ERROR: No Driver/Firmware information found in ASHCI-Catalog.xml for System Type:" $ServerType " with OS:" $OSVer -foregroundcolor red}
                    $SpecialCatalogNeeded="NO"
                }
        }
        }#end switch S2DCatalogNeeded
        Switch($SpecialCatalogNeeded){
            {$PSItem -eq "Precision"}{
                #Write-host "Expanding Special Catalog: $SpecialCatalogNeeded...."
                   #$SpecialCatalogFile
                    Write-host "Importing Special Catalog...."
                    $SpecialCatalogXMLDataFiltered=@()
                    $SpecialCatalogXMLDataFiltered = $PrecisionCatalog.Manifest.SoftwareComponent
                    $SpecialCatalogInfo=@()
                    $SpecialCatalogInfo=$PrecisionCatalogExtractedPath.Name
                    $SpecialCatalogInfo+="<br> "+$SpecialCatalogXMLDataFiltered.Manifest.version
                    $CatVerInfo+="<br> $PrecisionCatalogExtractedPath.Name Info <br>&nbsp&nbspData/Time: "+$SpecialCatalogXMLDataFiltered.Manifest.dateTime+"<br>&nbsp&nbspReleaseId: "+$SpecialCatalogXMLDataFiltered.Manifest.releaseID+"<br>&nbsp&nbspVersion: "+$SpecialCatalogXMLDataFiltered.Manifest.version
                    IF(!$SpecialCatalogXMLDataFiltered){Write-host "    ERROR: No Driver/Firmware information found in catalog.xml for System Type:" $ServerType " with OS:" $OSVer -foregroundcolor red}

            }
        }#End of switch SpecialCatalogNeeded
            #{$PSItem -eq "NO"}{
                # Reloads the catalog.xml data
                IF(($SpecialCatalogNeeded -ne 'Precision') -and ($SkipDriversandFirmware -eq "NO")){
                    Write-host "Expanding Catalog.cab...."
                    IF(!(Test-Path "$ExtracLoc\Catalog.xml")){Expand-Cab -SourceFile $DownloadFile -TargetFolder $ExtracLoc -Item "Catalog.xml" -Force}
                    Write-host "Importing Catalog.xml...."
                    #$Data = $CatalogXMLData
                    #$Catalogs+= $CatalogXMLData
                    Write-host "Filtering Catalog.xml for latest PowerEdge Firmware and Drivers...."
                    $CatalogXMLDataFiltered=@()
                    $CatalogXMLDataFiltered+=$CatalogXMLData.Manifest.SoftwareComponent|`
                    Where-Object{($_.SupportedSystems.Brand.Model.Display."#cdata-section" -eq $ServerType)-or($_.SupportedSystems.Brand.Model.systemID -eq $SystemID)-or($_.SupportedSystems.Brand.Model.systemID -match 'VRTX')}|`
                    Where-Object{($_.packageType -eq "LW64")-or($_.packageType -eq "LWXP")}
                    $CatalogInfo=@()
                    $CatalogInfo="Catalog.xml"
                    #$CatalogInfo+="<br> "+$CatalogXMLData.Manifest.dateTime
                    #$CatalogInfo+="<br> "+$CatalogXMLData.Manifest.releaseID
                    $CatalogInfo+="<br> "+$CatalogXMLData.Manifest.version
                    $CatVerInfo+="<br> Catalog.xml Info <br>&nbsp&nbspData/Time: "+$CatalogXMLData.Manifest.dateTime+"<br>&nbsp&nbspReleaseId: "+$CatalogXMLData.Manifest.releaseID+"<br>&nbsp&nbspVersion: "+$CatalogXMLData.Manifest.version
                    IF(!$CatalogXMLDataFiltered){Write-host "    ERROR: No Driver/Firmware information found in catalog.xml for System Type:" $ServerType " with OS:" $OSVer -foregroundcolor red}
                }

         #   }
        #}

    # Combine Catalogs

    # Drivers and Firmware for PowerEdge Servers
  IF($SkipDriversandFirmware -eq "NO"){
    #Match Installed Hardware to Catalog
        Write-host "Comparing Installed Hardware to Dell Catalog...."
        Write-Host "    Match rule: filtered catalog rows only; BIOS matches by componentID; all others require ComponentType plus componentID or full PCI identity."
        $DriFT17GMatchDebug = @()
        If(($SupportAssistDataType -eq "SAEX")-or($SupportAssistDataType -eq "TSR")){
            ForEach ($Device in $InstalledHardwareUnique){
                $Found=@()
                If($S2DCatalogNeeded -eq "YES"){
                    $CatalogInfoOut=""
                    #Added for iDRAC 4.40 weird chars
                            IF($Device.deviceID.length -gt 0){
                                $Found=
                                $S2DCatalogXMLDataFiltered|`
                                Where-Object{Test-DriFTCatalogDeviceMatch -CatalogDevice $_ -Device $Device}|`
                                sort-Object {[DateTime]$_.releaseDate}| Select-Object -Last 1;`
                                $CatalogInfoOut=$S2DCatalogInfo
                                }
                            Else{
                                $Found=
                                $S2DCatalogXMLDataFiltered|`
                                Where-Object{Test-DriFTCatalogDeviceMatch -CatalogDevice $_ -Device $Device}|`
                                sort-Object {[DateTime]$_.releaseDate}| Select-Object -Last 1;`
                                $CatalogInfoOut=$S2DCatalogInfo
                                }
                            
                    #Write-Host "S2D:" $Found.Name.Display."#cdata-section"
                    # Check Catalog.xml if not found in special catalog
                }
                IF($SpecialCatalogNeeded -eq 'Precision'){
                    $CatalogInfoOut=""
                    #Added for iDRAC 4.40 weird chars
                            IF($Device.deviceID.length -gt 0){
                                $Found=$SpecialCatalogXMLDataFiltered|`
                                Where-Object{Test-DriFTCatalogDeviceMatch -CatalogDevice $_ -Device $Device}|`
                                sort-Object {[DateTime]$_.releaseDate}| Select-Object -Last 1;`
                                $CatalogInfoOut=$SpecialCatalogInfo
                                }
                            Else{
                                $Found=
                                $SpecialCatalogXMLDataFiltered|`
                                Where-Object{Test-DriFTCatalogDeviceMatch -CatalogDevice $_ -Device $Device}|`
                                sort-Object {[DateTime]$_.releaseDate}| Select-Object -Last 1;`
                                $CatalogInfoOut=$SpecialCatalogInfo
                                }
                }
                IF(!$Found){
                            $CatalogInfoOut=""
                            #Added for iDRAC 4.40 weird chars
                            IF($Device.deviceID.length -gt 0){
                                $Found=
                                $CatalogXMLDataFiltered|`
                                Where-Object{Test-DriFTCatalogDeviceMatch -CatalogDevice $_ -Device $Device}|`
                                sort-Object {[DateTime]$_.releaseDate}| Select-Object -Last 1;`
                                $CatalogInfoOut=$CatalogInfo
                                }
                            Else{
                                $Found=
                                $CatalogXMLDataFiltered|`
                                Where-Object{Test-DriFTCatalogDeviceMatch -CatalogDevice $_ -Device $Device}|`
                                sort-Object {[DateTime]$_.releaseDate}| Select-Object -Last 1;`
                                $CatalogInfoOut=$CatalogInfo
                                }
                            }
                # 17G matching is handled by Test-DriFTCatalogDeviceMatch against the already-filtered catalog rows above.
                if ($Is17GTSR) {
                    $CatalogComponentMatch = $false
                    $CatalogPciMatch = $false
                    $MatchMethod = ''
                    $UnmatchedReason = ''

                    if ($Found) {
                        $CatalogComponentMatch = Test-DriFTCatalogComponentIdMatch -CatalogObject $Found -ComponentId $Device.componentID
                        $CatalogPciMatch = Test-DriFTCatalogPciIdentityMatch -CatalogObject $Found -Device $Device

                        if ($CatalogComponentMatch -and $CatalogPciMatch) { $MatchMethod = 'ComponentID+PCI' }
                        elseif ($CatalogComponentMatch) { $MatchMethod = 'ComponentID' }
                        elseif ($CatalogPciMatch) { $MatchMethod = 'PCI' }
                        else { $MatchMethod = 'MatchedByPipelineUnknown' }
                    }
                    else {
                        $HasValidComponentId = Test-DriFTInstalledComponentIdIsValid -ComponentId $Device.componentID
                        $HasPciIdentity = Test-DriFTDeviceHasPciIdentity -Device $Device

                        if (-not $HasValidComponentId -and -not $HasPciIdentity) {
                            $UnmatchedReason = 'No valid componentID and no complete PCI identity'
                        }
                        elseif (-not $HasValidComponentId) {
                            $UnmatchedReason = 'No valid componentID; PCI identity did not match filtered catalog'
                        }
                        elseif (-not $HasPciIdentity) {
                            $UnmatchedReason = 'componentID did not match filtered catalog; no complete PCI identity'
                        }
                        else {
                            $UnmatchedReason = 'componentID and PCI identity did not match filtered catalog'
                        }
                    }

                    $DriFT17GMatchDebug += [PSCustomObject]@{
                        DeviceComponentType = $Device.componentType
                        DeviceComponentID   = $Device.componentID
                        DeviceVendorID      = $Device.vendorID
                        DeviceDeviceID      = $Device.deviceID
                        DeviceSubDeviceID   = $Device.subDeviceID
                        DeviceSubVendorID   = $Device.subVendorID
                        DeviceVersion       = $Device.version
                        DeviceDisplay       = $Device.display
                        DeviceElementName   = $Device.ElementName
                        DeviceRelatedItem   = $Device.RelatedItem
                        Matched             = [bool]$Found
                        MatchMethod         = $MatchMethod
                        UnmatchedReason     = $UnmatchedReason
                        CatalogComponentType= if ($Found) { Get-DriFTCatalogComponentTypeValue -CatalogObject $Found } else { '' }
                        CatalogName         = if ($Found) { $Found.Name.Display."#cdata-section" } else { '' }
                        CatalogComponentIDs = if ($Found) { (Get-DriFTCatalogComponentIdValues -CatalogObject $Found) -join '|' } else { '' }
                        CatalogVendorVersion= if ($Found) { $Found.vendorVersion } else { '' }
                        CatalogReleaseDate  = if ($Found) { ($Found.dateTime -split 'T')[0] } else { '' }
                        CatalogPath         = if ($Found) { $Found.path } else { '' }
                    }
                }
                            #pause
                #Write-Host "CAT:" $Found.Name.Display."#cdata-section"         
                $allArray+=$Found|`
                Select-object @{Label="ServiceTag";Expression={"$ServiceTag"}},`
                @{Label="PowerEdge";Expression={"$ServerType"}},`
                @{Label="OS";Expression={"$InstalledOS"+" "+"$OSVersion"}},`
                @{Label="Type";Expression={If($Device.componentType -is [system.array])`
                 {$Device.componentType[0]}Else{$Device.componentType}}},`
                @{Label="Category";Expression={$_.LUCategory.value}},`
                @{Label="Name";Expression={$_.Name.Display."#cdata-section"}},`
                @{Label="InstalledVersion";Expression={`
                    $DeviceVer=$Device.version
                    $VVersion=$_.vendorVersion
                    If($Device.version -is [system.array])`
                        {$DeviceVer=$Device.version[0]}Else{$DeviceVer=$Device.version};`
                    If($Device.version -match "OSC_")`
                        {$DeviceVer=$Device.version -replace "OSC_"};`
                    If($DeviceVer -is [System.Version])`
                        {If([System.Version]$DeviceVer -lt [System.Version]$_.vendorVersion){"***"+$DeviceVer}Else{$DeviceVer}};`
                    Try{
                        If([Version]$DeviceVer -lt [Version]$_.vendorVersion){"***"+$DeviceVer}Else{$DeviceVer}}
                    Catch{
                          If($DeviceVer -lt $VVersion){"***"+$DeviceVer}Else{$DeviceVer}}
                    Finally{}}},`
                @{Label="AvailableVersion";Expression={$_.vendorVersion}},`
                @{Label="CatalogInfo";Expression={$CatalogInfoOut}},`
                @{Label="Criticality";Expression={$_.Criticality.Display."#cdata-section" -replace '\d\-'}},`
                @{Label="ReleaseDate";Expression={($_.dateTime -split "T")[0]}},`
                @{Label="URL";Expression={$DellURL+$_.path}},`
                @{Label="Details";Expression={$_.ImportantInfo.URL}}|`
                sort-object Type,Category,Name
                
            }
            
          }
          if ($Is17GTSR -and $DriFT17GMatchDebug) {
              $DriFT17GMatchedDebug = @($DriFT17GMatchDebug | Where-Object { $_.Matched -eq $true })
              $DriFT17GUnmatchedDebug = @($DriFT17GMatchDebug | Where-Object { $_.Matched -ne $true })

              Export-DriFT17GDebug -Name "DriFT_17G_CatalogMatch_Debug.csv" -InputObject $DriFT17GMatchDebug | Out-Null
              Export-DriFT17GDebug -Name "DriFT_17G_CatalogMatched_Debug.csv" -InputObject $DriFT17GMatchedDebug | Out-Null
              Export-DriFT17GDebug -Name "DriFT_17G_CatalogUnmatched_Debug.csv" -InputObject $DriFT17GUnmatchedDebug | Out-Null

              Write-Host "    17G catalog matched rows: $(@($DriFT17GMatchedDebug).Count)"
              Write-Host "    17G catalog unmatched rows: $(@($DriFT17GUnmatchedDebug).Count)"
              Write-Host "    17G catalog unmatched debug: $env:TEMP\DriFT_17G_Debug\DriFT_17G_CatalogUnmatched_Debug.csv"
          }
          #pause
  IF($InstalledOS -imatch "Windows" -and $S2DCatalogNeeded -eq "YES"){
    #Gathering Certified CPLD version for Ready Nodes
        Write-Host "Checking for Certified CPLD version for AX/Ready Nodes..."
        $CPLDURL='https://www.dell.com/support/kbdoc/en-us/000127931/firmware-and-driver-update-catalog-for-dell-emc-solutions-for-microsoft-azure-stack-hci'
        IF(-not($CPLDRequest)){
        $CPLDRequest = Invoke-WebRequest -Uri $CPLDURL}

    #Find table in the website
        $CPLDtableHeader = $CPLDRequest.AllElements | Where-Object {$_.tagname -eq 'th'}
        $CPLDtableData = $CPLDRequest.AllElements | Where-Object {$_.tagname -eq 'td'}

    #Table header and data
        $CPLDthead = $CPLDtableHeader.innerText
        $CPLDtdata = $CPLDtableData.innerText

    #Break table data into smaller chuck of data.
        $CPLDdataResult = New-Object System.Collections.ArrayList
        for ($i = 0; $i -le $CPLDtdata.count; $i+= ($CPLDthead.count - 1))
        {
            if ($CPLDtdata.count -eq $i)
            {
                break
            }        
            $CPLDgroup = $i + ($CPLDthead.count - 1)
            [void]$CPLDdataResult.Add($CPLDtdata[$i..$CPLDgroup])
            $i++
        }

    #Html data into powershell table format
        $CPLDfinalResult = @()
        foreach ($CPLDdata in $CPLDdataResult)
        {
            $CPLDnewObject = New-Object psobject
            for ($i = 0; $i -le ($CPLDthead.count - 1); $i++) {
                $CPLDnewObject | Add-Member -Name $CPLDthead[$i] -MemberType NoteProperty -value $CPLDdata[$i]
            }
            $CPLDfinalResult += $CPLDnewObject
        }
        #$CPLDfinalResult | ft -AutoSize

        $CPLDFW = $CPLDfinalResult | Select Component,Type,Category,'Software Bundle',@{L='Version';E={($_.'Minimum Supported Version' -split ' -')[0]}},@{L='ServerType';E={(($_.'Minimum Supported Version' -split ' - ')[-1]).Trim() -replace ' Ready Node',''}},@{L='Documentation';E={"https://www.dell.com/support/home/en-us/drivers/driversdetails?driverid="+$_.'Software Bundle'}}
        $AXServerType=$ServerType -replace 'R64','AX-64'-replace 'R740XD','AX-740XD' -replace 'R6515','AX-6515' -replace 'R7525','AX-7525'
        $CPLD4ServerType=($CPLDFW | Where-Object{$_.ServerType -ilike $AXServerType})
    $FoundCPLD=@()
    $FoundCPLD=$InstalledHardwareUnique | Where-Object{$_.Display -imatch "CPLD"}
    IF($FoundCPLD){
        IF($CPLD4ServerType){
            IF(-not($CPDLDetails)){
            $CPDLDetails=Invoke-WebRequest -Uri $CPLD4ServerType.Documentation}
            #Lookup the dowhnload link
            $DLLink=""
            $DLLink=$CPDLDetails.links.href | Where-Object{$_ -imatch "$($CPLD4ServerType.'Software Bundle')"+"_WN64"} | Where-Object {$_ -imatch ".EXE"}
            $DLLinkVersion=""
            $DLLinkVersion=(($DLLink -split '/')[-1] -split '_')[-2]
            $CPLDURI=($CPLDFW | Where-Object{$_.ServerType -imatch $ServerType}).Documentation
            $CPLD= $CPLD4ServerType|Select-Object `
                     @{Label="ServiceTag";Expression={"$ServiceTag"}}`
                    ,@{Label="PowerEdge";Expression={"$ServerType"}}`
                    ,@{Label="OS";Expression={"$InstalledOS"+" "+"$OSVersion"}}`
                    ,@{Label="Type";Expression={"FRMW"}}`
                    ,@{Label="Category";Expression={$_.Category}}`
                    ,@{Label="Name";Expression={($CPDLDetails.ParsedHtml.IHTMLDocument3_documentElement.getElementsBytagName('h1')| Select-Object innertext).innertext}}`
                    ,@{Label="InstalledVersion";Expression={If([System.Version]$FoundCPLD.Version -lt [System.Version]$DLLinkVersion){"***"+$FoundCPLD.Version}Else{$FoundCPLD.Version}}}`
                    ,@{Label="AvailableVersion";Expression={$DLLinkVersion}}`
                    ,@{Label="CatalogInfo";Expression={"support.dell.com"}}`
                    ,@{Label="Criticality";Expression={((($CPDLDetails.ParsedHtml.IHTMLDocument3_documentElement.getElementsByClassName("h4")|Where-Object {$_.getAttributeNode('Id').Value -eq 'driverImportanceForDriver'}).parentElement.innerText) -split "`n")[-1]}}`
                    ,@{Label="ReleaseDate";Expression={(($CPDLDetails.ParsedHtml.IHTMLDocument3_documentElement.getElementsByClassName('h4')|Where-Object {$_.getAttributeNode('Id').Value -eq 'driverRDFordriver'}).parentElement.textContent -replace 'Release date ').trim()}}`
                    ,@{Label="URL";Expression={$DLLink}}`
                    ,@{Label="Details";Expression={$_.Documentation}}`
                    | sort-object Type,Category,Name
                $allArray+= $CPLD|sort-object URl -Unique
                
        }
    }#IF($FoundCPLD
}
IF($InstalledOS -imatch "Windows"){
    #Chipset Driver
        IF(!($allArray|Where-Object{($_.ServiceTag -eq $ServiceTag)-and($_.Type -eq 'DRVR')-and($_.Category -eq 'Chipset')})){
            $ChpsDrv=@()
            $ChpsNoUsb=@()
            $Chps=@()
            Write-host "Checking for Chipset Driver in Catalog...."
            $ChpsDrv=$S2DCatalogXMLDataFiltered|Where-Object{$_.LUCategory.value -eq "Chipset"}|`
                     Where-Object{$_.ComponentType.value -eq "DRVR"}|`
                     Where-Object{$_.SupportedOperatingSystems.OperatingSystem.majorVersion -eq $OSMjrVer}|`
                     Where-Object{$_.SupportedOperatingSystems.OperatingSystem.minorVersion -eq $OSMinVer};`
                     $CatalogInfoOut=$S2DCatalogInfo
            # Check Catalog.xml if not found in special catalog
            If(!$ChpsDrv){
                         $ChpsDrv=$CatalogXMLDataFiltered|Where-Object{$_.LUCategory.value -eq "Chipset"}|`
                         Where-Object{$_.ComponentType.value -eq "DRVR"}|`
                         Where-Object{$_.SupportedOperatingSystems.OperatingSystem.majorVersion -eq $OSMjrVer}|`
                         Where-Object{$_.SupportedOperatingSystems.OperatingSystem.minorVersion -eq $OSMinVer};`
                         $CatalogInfoOut=$CatalogInfo
                         }
            $ChpsNoUsb=$ChpsDrv | Where-Object {$_.Name.Display."#cdata-section" -notlike "*USB*"}
            $Chps=$ChpsNoUsb |sort-Object {[DateTime]$_.releaseDate} | Select-object -last 1 `
                 @{Label="ServiceTag";Expression={"$ServiceTag"}}`
                ,@{Label="PowerEdge";Expression={"$ServerType"}}`
                ,@{Label="OS";Expression={"$InstalledOS"+" "+"$OSVersion"}}`
                ,@{Label="Type";Expression={"DRVR"}}`
                ,@{Label="Category";Expression={$_.LUCategory.value}}`
                ,@{Label="Name";Expression={$_.Name.Display."#cdata-section"}}`
                ,@{Label="InstalledVersion";Expression={"NA"}}`
                ,@{Label="AvailableVersion";Expression={$_.vendorVersion}}`
                ,@{Label="CatalogInfo";Expression={$CatalogInfoOut}}`
                ,@{Label="Criticality";Expression={$_.Criticality.Display."#cdata-section" -replace '\d\-'}}`
                ,@{Label="ReleaseDate";Expression={($_.dateTime -split "T")[0]}}`
                ,@{Label="URL";Expression={$DellURL+$_.path}}`
                ,@{Label="Details";Expression={$_.ImportantInfo.URL}}`
                | sort-object Type,Category,Name
            $allArray+= $Chps|sort-object URl -Unique
        }# end of Chipset Driver
    #Fibre Channel Driver
        IF(!($allArray|Where-Object{($_.ServiceTag -eq $ServiceTag)-and($_.Type -eq 'DRVR')-and($_.Category -eq 'Fibre Channel')})){
            $FCDrv=@()
            $FCDevice=@()
            $FoundFCDrv=@()
            $InstalledNICs=@()
            Write-host "Checking for Fibre Channel Driver in Catalog...."
            $FoundFCDrv=$InstalledHardwareUnique|Where-Object{($_.Display -Like 'FC.*')}
            ForEach ($FCDevice in $FoundFCDrv){
                $FoundFCDrv=$S2DCatalogXMLDataFiltered|Where-Object{$_.LUCategory.value -eq "Fibre Channel"}|`
                            Where-Object{$_.ComponentType.value -eq "DRVR"}|`
                            Where-Object{$_.SupportedOperatingSystems.OperatingSystem.majorVersion -eq $OSMjrVer}|`
                            Where-Object{$_.SupportedOperatingSystems.OperatingSystem.minorVersion -eq $OSMinVer}|`
                            sort-Object {[DateTime]$_.releaseDate}| Select-Object -Last 1;`
                            $CatalogInfoOut=$S2DCatalogInfo
                # Check Catalog.xml if not found in special catalog
                If(!$FoundFCDrv){
                     $FoundFCDrv=$CatalogXMLDataFiltered|Where-Object{$_.LUCategory.value -eq "Fibre Channel"}|`
                     Where-Object{$_.ComponentType.value -eq "DRVR"}|`
                     Where-Object{$_.SupportedOperatingSystems.OperatingSystem.majorVersion -eq $OSMjrVer}|`
                     Where-Object{$_.SupportedOperatingSystems.OperatingSystem.minorVersion -eq $OSMinVer}|`
                     sort-Object {[DateTime]$_.releaseDate}| Select-Object -Last 1;`
                     $CatalogInfoOut=$CatalogInfo 
                     }
                $FCDrv+=$FoundFCDrv|`
                 Select-Object @{Label="ServiceTag";Expression={"$ServiceTag"}}`
                ,@{Label="PowerEdge";Expression={"$ServerType"}}`
                ,@{Label="OS";Expression={"$InstalledOS"+" "+"$OSVersion"}}`
                ,@{Label="Type";Expression={"DRVR"}}`
                ,@{Label="Category";Expression={$_.LUCategory.value}}`
                ,@{Label="Name";Expression={$_.Name.Display."#cdata-section"}}`
                ,@{Label="InstalledVersion";Expression={"NA"}}`
                ,@{Label="AvailableVersion";Expression={$_.vendorVersion}}`
                ,@{Label="CatalogInfo";Expression={$CatalogInfoOut}}`
                ,@{Label="Criticality";Expression={$_.Criticality.Display."#cdata-section" -replace '\d\-'}}`
                ,@{Label="ReleaseDate";Expression={($_.dateTime -split "T")[0]}}`
                ,@{Label="URL";Expression={$DellURL+$_.path}}`
                ,@{Label="Details";Expression={$_.ImportantInfo.URL}}`
                | sort-object Type,Category,Name
            }
            $allArray+=$FCDrv|sort-object URl -Unique
        }
 
    #Network Driver
        IF(!($allArray|Where-Object{($_.ServiceTag -eq $ServiceTag)-and($_.Type -eq 'DRVR')-and($_.Category -eq 'Network')})){
            $NICDrv=@()
            $NDevice=@()
            $InstalledNICs=@()
            $FoundNICDrv=@()
            Write-host "Checking for Network Driver in Catalog...."
            $InstalledNICs=$InstalledHardwareUnique|`
                Where-Object{($_.Display -Match 'NIC')`
                -or($_.Display -Match 'Ethernet')`
                -or($_.Display -Match 'giga')`
                -or($_.Display -Match 'FastLinQ')`
                -or($_.Display -Match 'QLogic')}
            ForEach ($NDevice in $InstalledNICs){
                $FoundNICDrv=$S2DCatalogXMLDataFiltered|`
                            Where-Object{$_.ComponentType.value -eq 'DRVR'}|`
                            Where-Object{$_.LUCategory.value -eq 'Network'}|`
                            Where-Object{
                                (($_.SupportedDevices.Device.PCIInfo.deviceID -eq $NDevice.deviceID)`
                                -and`
                                ($_.SupportedDevices.Device.PCIInfo.vendorID -eq $NDevice.vendorID))`
                                -or` # Added to make sure the Intel X710 gets added to the report as PCI IDs are missing for this driver
                                (($_.SupportedSystems.Brand.Model.SystemID -eq $SystemID)`
                                -and`
                                ($_.Description.Display.'#cdata-section' -Like "*Intel(R) Ethernet 10G X710 rNDC*"))}|`
                            sort-Object {[DateTime]$_.releaseDate}| Select-Object -Last 1;`
                            $CatalogInfoOut=$S2DCatalogInfo 
                # Check Catalog.xml if not found in special catalog
                IF(!$FoundNICDrv){
                                 $FoundNICDrv=$CatalogXMLDataFiltered|`
                                 Where-Object{$_.ComponentType.value -eq 'DRVR'}|`
                                 Where-Object{$_.LUCategory.value -eq 'Network'}|`
                                 Where-Object{
                                    (($_.SupportedDevices.Device.PCIInfo.deviceID -eq $NDevice.deviceID)`
                                    -and`
                                    ($_.SupportedDevices.Device.PCIInfo.vendorID -eq $NDevice.vendorID))`
                                    -or` # Added to make sure the Intel X710 gets added to the report as PCI IDs are missing for this driver
                                    (($_.SupportedSystems.Brand.Model.SystemID -eq $SystemID)`
                                    -and`
                                    ($_.Description.Display.'#cdata-section' -Like "*Intel(R) Ethernet 10G X710 rNDC*"))}|`
                                 sort-Object {[DateTime]$_.releaseDate}| Select-Object -Last 1;`
                                 $CatalogInfoOut=$CatalogInfo
                                 }
                $NICDrv+=$FoundNICDrv|`
                 Select-Object @{Label="ServiceTag";Expression={"$ServiceTag"}}`
                ,@{Label="PowerEdge";Expression={"$ServerType"}}`
                ,@{Label="OS";Expression={"$InstalledOS"+" "+"$OSVersion"}}`
                ,@{Label="Type";Expression={"DRVR"}}`
                ,@{Label="Category";Expression={$_.LUCategory.value}}`
                ,@{Label="Name";Expression={$_.Name.Display."#cdata-section"}}`
                ,@{Label="InstalledVersion";Expression={"NA"}}`
                ,@{Label="AvailableVersion";Expression={$_.vendorVersion}}`
                ,@{Label="CatalogInfo";Expression={$CatalogInfoOut}}`
                ,@{Label="Criticality";Expression={$_.Criticality.Display."#cdata-section" -replace '\d\-'}}`
                ,@{Label="ReleaseDate";Expression={($_.dateTime -split "T")[0]}}`
                ,@{Label="URL";Expression={$DellURL+$_.path}}`
                ,@{Label="Details";Expression={$_.ImportantInfo.URL}}`
                | sort-object Type,Category,Name
            }
            $allArray+=$NICDrv|sort-object URl -Unique
        }
    #RAID Drivers
        IF(!($allArray|Where-Object{($_.ServiceTag -eq $ServiceTag)-and($_.Type -eq 'DRVR')-and($_.Category -eq 'SAS RAID')})){
            $RAIDDrv=@()
            $RAIDDevice=@()
            $InstalledRAID=@()
            $FoundRAIDDrv=@()
            Write-host "Checking for RAID Drivers in Catalog...."
            #AHCI is for BOSS card
            $InstalledRAID=$InstalledHardwareUnique|Where-Object{(($_.Display -Match 'RAID')-or($_.Display -Match 'AHCI'))}
            ForEach ($RAIDDevice in $InstalledRAID){
                $FoundRAIDDrv=$S2DCatalogXMLDataFiltered|`
                              Where-Object{$_.ComponentType.value -eq "DRVR"}|`
                              Where-Object{$_.LUCategory.value -match 'RAID'}|`
                              Where-Object{$_.SupportedOperatingSystems.OperatingSystem.majorVersion -eq $OSMjrVer}|`
                              Where-Object{$_.SupportedOperatingSystems.OperatingSystem.minorVersion -eq $OSMinVer}|`
                              Where-Object{(($_.SupportedDevices.Device.PCIInfo.deviceID -eq $RAIDDevice.deviceID)`
                              -and($_.SupportedDevices.Device.PCIInfo.subDeviceID -eq $RAIDDevice.subdeviceID)`
                              -and($_.SupportedDevices.Device.PCIInfo.subVendorID -eq $RAIDDevice.subVendorID)`
                              -and($_.SupportedDevices.Device.PCIInfo.vendorID -eq $RAIDDevice.vendorID))}|`
                              sort-Object {[DateTime]$_.releaseDate}| Select-Object -Last 1;`
                              $CatalogInfoOut=$S2DCatalogInfo
                IF(!$FoundRAIDDrv){
                                   $FoundRAIDDrv=$CatalogXMLDataFiltered|`
                                   Where-Object{$_.ComponentType.value -eq "DRVR"}|`
                                   Where-Object{$_.LUCategory.value -match 'RAID'}|`
                                   Where-Object{$_.SupportedOperatingSystems.OperatingSystem.majorVersion -eq $OSMjrVer}|`
                                   Where-Object{$_.SupportedOperatingSystems.OperatingSystem.minorVersion -eq $OSMinVer}|`
                                   Where-Object{(($_.SupportedDevices.Device.PCIInfo.deviceID -eq $RAIDDevice.deviceID)`
                                   -and($_.SupportedDevices.Device.PCIInfo.subDeviceID -eq $RAIDDevice.subdeviceID)`
                                   -and($_.SupportedDevices.Device.PCIInfo.subVendorID -eq $RAIDDevice.subVendorID)`
                                   -and($_.SupportedDevices.Device.PCIInfo.vendorID -eq $RAIDDevice.vendorID))}|`
                                   sort-Object {[DateTime]$_.releaseDate}| Select-Object -Last 1;`
                                   $CatalogInfoOut=$CatalogInfo
                                   }

                $RAIDDrv+=$FoundRAIDDrv|`
                 Select-Object @{Label="ServiceTag";Expression={"$ServiceTag"}}`
                ,@{Label="PowerEdge";Expression={"$ServerType"}}`
                ,@{Label="OS";Expression={"$InstalledOS"+" "+"$OSVersion"}}`
                ,@{Label="Type";Expression={"DRVR"}}`
                ,@{Label="Category";Expression={$_.LUCategory.value}}`
                ,@{Label="Name";Expression={$_.Name.Display."#cdata-section"}}`
                ,@{Label="InstalledVersion";Expression={"NA"}}`
                ,@{Label="AvailableVersion";Expression={$_.vendorVersion}}`
                ,@{Label="CatalogInfo";Expression={$CatalogInfoOut}}`
                ,@{Label="Criticality";Expression={$_.Criticality.Display."#cdata-section" -replace '\d\-'}}`
                ,@{Label="ReleaseDate";Expression={($_.dateTime -split "T")[0]}}`
                ,@{Label="URL";Expression={$DellURL+$_.path}}`
                ,@{Label="Details";Expression={$_.ImportantInfo.URL}}`
                | sort-object Type,Category,Name
            }
            $allArray+=$RAIDDrv|sort-object URl -Unique
        }
  }# $InstalledOS -imatch "Windows"

  }# End of IF($SkipDriversandFirmware -eq "YES"
  IF($InstalledOS -imatch "VMware"){
    Write-Host "Gathering VMware supported driver versions from Broadcom Compatibility Guide..."
    
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
function Get-BroadcomIoCompatibility {
    param(
        [Parameter(Mandatory)]
        [string]$VendorID,

        [Parameter(Mandatory)]
        [string]$DeviceID,

        [Parameter(Mandatory)]
        [string]$SubVendorID,

        [Parameter(Mandatory)]
        [string]$SubDeviceID,

        [Parameter(Mandatory)]
        [string]$EsxiVersion
    )

    $release = "ESXi $EsxiVersion"

    $body = @{
        programId = "io"
        filters   = @(
            @{ displayKey = "vid"; filterValues = @($VendorID) },
            @{ displayKey = "did"; filterValues = @($DeviceID) },
            @{ displayKey = "svid"; filterValues = @($SubVendorID) },
            @{ displayKey = "ssid"; filterValues = @($SubDeviceID) },
            @{ displayKey = "productReleaseVersion"; filterValues = @($release) }
        )
        keyword = @()
    } | ConvertTo-Json -Depth 6

    Invoke-RestMethod `
        -Uri "https://compatibilityguide.broadcom.com/compguide/programs/viewResults?limit=50&page=1&sortBy=&sortType=ASC" `
        -Method Post `
        -ContentType "application/json" `
        -Body $body `
        -UseBasicParsing
}

function Get-BcgProductInfo {
    param(
        [Parameter(Mandatory)]
        $BcgResult
    )

    $productId = ($BcgResult.model[0].url -replace '.*productId=', '')
    $model     = $BcgResult.model[0].name

    $deviceType = "No Data Found"

    if ($BcgResult.deviceType -and $BcgResult.deviceType[0].name) {
        $deviceType = $BcgResult.deviceType[0].name
    }
    elseif ($BcgResult.deviceTypes -and $BcgResult.deviceTypes[0].name) {
        $deviceType = $BcgResult.deviceTypes[0].name
    }
    elseif ($BcgResult.hoverData) {
        $hoverDeviceType = $BcgResult.hoverData |
            Where-Object { $_.displayName -eq "Device Type" } |
            Select-Object -First 1

        if ($hoverDeviceType.value) {
            $deviceType = $hoverDeviceType.value
        }
    }

    [PSCustomObject]@{
        ProductId   = $productId
        Model       = $model
        DeviceType  = $deviceType
        DetailsLink = "https://compatibilityguide.broadcom.com/detail?persona=live&productId=$productId&program=io"
    }
}

function Get-BcgCompatibilityMatch {
    param(
        [Parameter(Mandatory)]
        $Device,

        [Parameter(Mandatory)]
        [string]$EsxiVersion
    )

    $found = Get-BroadcomIoCompatibility `
        -VendorID $Device.VendorID `
        -DeviceID $Device.DeviceID `
        -SubVendorID $Device.SubVendorID `
        -SubDeviceID $Device.SubDeviceID `
        -EsxiVersion $EsxiVersion

    if (-not $found -or $found.success -ne $true -or $found.data.count -eq 0) {
        $found = Get-BroadcomIoCompatibility `
            -VendorID $Device.VendorID `
            -DeviceID $Device.DeviceID `
            -SubVendorID $Device.SubDeviceID `
            -SubDeviceID $Device.SubVendorID `
            -EsxiVersion $EsxiVersion
    }

    return $found
}        
    #Match Installed Hardware to VMware Catalog
        Write-host "Comparing Installed Hardware to VMware Compatibility Guide...."
        If(($SupportAssistDataType -eq "SAEX")-or($SupportAssistDataType -eq "TSR")){
            $InstalledHardwareUniqueWithDeviceInfo=@()
            $InstalledHardwareUniqueWithDeviceInfo=$InstalledHardwareUnique | Where-Object{$_.SubDeviceID.length -gt 0}
            #Write-Host "    Devices found"
            ForEach ($Device in $InstalledHardwareUniqueWithDeviceInfo){
                $Found = Get-BcgCompatibilityMatch -Device $Device -EsxiVersion $VMWOSVer
                if ($Found -and $Found.success -eq $true -and $Found.data.count -gt 0) {
                    $BCGResult = $Found.data.fieldValues | Select-Object -First 1
                    $BCGInfo   = Get-BcgProductInfo -BcgResult $BCGResult
                    $allArray += [PSCustomObject]@{
                        ServiceTag       = "$ServiceTag"
                        PowerEdge        = "$ServerType"
                        OS               = "$InstalledOS $OSVersion"
                        Type             = "DRVR"
                        Category         = $BCGInfo.DeviceType
                        Name             = $BCGInfo.Model
                        InstalledVersion = "NA"
                        AvailableVersion = "See Broadcom Compatibility Guide"
                        CatalogInfo      = "Broadcom Compatibility Guide"
                        Criticality      = "No Data Found"
                        ReleaseDate      = "No Data Found"
                        Details          = "<a href='$($BCGInfo.DetailsLink)' target='_blank'>Broadcom Compatibility Guide</a><br>"
                        URL              = "BCG Details: <a href='$($BCGInfo.DetailsLink)' target='_blank'>$($BCGInfo.DetailsLink)</a><br>"
                    }
                }
            }
          }
  }# End VMWare section
                     # Analys sel log CurrentMBSel.txt for Warnings and Errors
                #$e="C:\Users\jim_gandy\OneDrive - Dell Technologies\Documents\SRs\109739084\archive_04_22_2021_21_33_45\RDUAITSD01.APPINSTECH.local - 12 node\TSR20210421081924_HZ064Z2\TSR20210421081924_HZ064Z2.pl\tsr\hardware\CurrentMBSEL\CurrentMBSEL.txt"
                #$MBSelLogWarnERR
                #$MBSelLogWarnERR=""
                Write-Host "Checking for SEL log Warnings and Errors in the last 30 days..."
                $MBSelLogWarnERROut=@()
                $CurrentMBSelFullNames=@()
                # Avoid recursively walking the extracted 17G Redfish tree. Its paths can exceed
                # Windows PowerShell 5.1 MAX_PATH limits and SEL logs are not stored there.
                $CurrentMBSelSearchRoots = @(
                    (Join-Path $E 'tsr\hardware\CurrentMBSEL'),
                    (Join-Path $E 'hardware\CurrentMBSEL')
                ) | Where-Object { Test-Path $_ -PathType Container }

                foreach ($SelRoot in $CurrentMBSelSearchRoots) {
                    $CurrentMBSelFullNames += Get-ChildItem -Path $SelRoot -Filter CurrentMBSel.txt -File -Recurse -Force -ErrorAction SilentlyContinue | ForEach-Object{ $_.fullname }
                }

                # Legacy fallback only. For 17G, do not scan the entire extracted TSR because
                # redfishidracwalk paths are very deep and can throw PathTooLongException.
                if ((-not $CurrentMBSelFullNames) -and (-not $Is17GTSR)) {
                    $CurrentMBSelFullNames = Get-ChildItem -Path $E -Filter CurrentMBSel.txt -File -Recurse -Force -ErrorAction SilentlyContinue | ForEach-Object{ $_.fullname }
                }

                if ($CurrentMBSelFullNames){
                    ForEach($MBSelLog in $CurrentMBSelFullNames){
                        $MBSelLogContent=Get-Content $MBSelLog -Delimiter '\ssdlkfjhaslodfhijasl\s'
                        IF($MBSelLogContent -imatch 'Severity\s\:\s[3-4]'){
                            $MBSelLogContentEntries=$MBSelLogContent -split '\-{60,}[\n]'
                            ForEach($Entry in $MBSelLogContentEntries){
                                IF($Entry -imatch 'Severity\s\:\s[3-4]'){
                                    $MBSelLogContentEntriesLines=$Entry -split '[\n]'
                                    $MBSelLogWarnERR+=[PSCustomObject]@{
                                        Node = $ServiceTag -replace '\*{1,}'
                                        Record = ($MBSelLogContentEntriesLines | Where-Object{$_ -imatch 'Record'}) -replace 'Record\s\:\s'
                                        DateTime = ($MBSelLogContentEntriesLines | Where-Object{$_ -imatch 'Date\/Time'}) -replace 'Date\/Time\s\:\s'
                                        Severity = ($MBSelLogContentEntriesLines | Where-Object{$_ -imatch 'Severity'}) -replace 'Severity\s\:\s'
                                        Description = ($MBSelLogContentEntriesLines | Where-Object{$_ -imatch 'Description'}) -replace 'Description\s\:\s'
                                    }
                                }
                            }
                        }
                    }
                }Else{
                    Write-Host "    No SEL log found in TSR Data. Nothing to do."
                    $ContinueOn="NO"
                }
                IF(!$ContinueOn){
                    #$MBSelLogWarnERR | ft
                    function Convert-DateString ([String]$Date, [String[]]$Format){
                        $result = New-Object DateTime
                    
                        $convertible = [DateTime]::TryParseExact(
                            $Date,
                            $Format,
                            [System.Globalization.CultureInfo]::InvariantCulture,
                            [System.Globalization.DateTimeStyles]::None,
                            [ref]$result)
                    
                        if ($convertible) { $result }
                    }
                    #$MBSelLogWarnERR|ft
                    $MBSelLogWarnERR1=$MBSelLogWarnERR | Select-Object Node,@{L='DateTime';E={Convert-DateString -Date $_.DateTime -replace 'Date\/Time\s\:\s' -Format 'ddd MMM dd yyyy hh:mm:ss'}},Severity,Description
                    # Filter for last 30 days
                    $filterDate = [datetime]::Today.AddDays(-30)
                    $MBSelLogWarnERROut+=$MBSelLogWarnERR1| Where-Object {$_.DateTime -ge $filterDate}
                    IF(-not($MBSelLogWarnERROut)){Write-Host "    None found"}Else{Write-Host "$MBSelLogWarnERROut"}
                }
  #$allArray=@()
  #$allArray
  #Pause
    #Latest OS CU for 2008r2 - 2022
    IF($SkipDriversandFirmware -eq "NO"){
        If(($OSCheck -match '20')){
            $WSLCU=@()
            $WSLCUout=@()
            #$OSCheck='2019'
            Write-host "Found Microsoft $OSCheck Installed."
            Write-host "Checking for Latest Microsoft Windows Server Update...."
            # Get the latest OS Build KBs
            
#region Recommended updates and hotfixes for Windows Server 
        $dstart=Get-Date
#Returns the download link of KB from Microsoft catalog
if (-not ('GetKBDLLink' -as [type])) {
Add-Type @"
using System.Net;
using System.IO;
using System.Text.RegularExpressions;
//using System; //for console writeline

public static class GetKBDLLink
{
    public static string GetDownloadLink(string KBNumber, string Product)
    {
        string kbGUID = "";
        string kbDLUriSource = "";

        // Search for URI to retrieve the latest KB information
        // Extracting the KBGUID from the KBPage
        var webRequest = WebRequest.Create("https://www.catalog.update.microsoft.com/Search.aspx?q=" + KBNumber);
        webRequest.Method = "GET";
        var webResponse = webRequest.GetResponse();
        var responseStream = webResponse.GetResponseStream();
        var streamReader = new StreamReader(responseStream);
        string responseContent = streamReader.ReadToEnd();
        //Console.WriteLine("Content is "+responseContent);
        var kbMatches=Regex.Matches(responseContent, @"id=(?:""|')(.*?)(?=_link)([^\/]*)");
        //Console.WriteLine(kbMatches.Count);
        foreach (Match ItemMatch in kbMatches)
        {
        //Console.WriteLine(Product);
        //Console.WriteLine(ItemMatch.Groups[2].Value);
            if (ItemMatch.Groups[2].Value.Contains(Product))
            {
                 kbGUID = ItemMatch.Groups[1].Value;
            }
        }
        //Console.WriteLine(kbGUID);

        // Use the KBGUID to find the actual download link for the KB
        string post1 = "https://www.catalog.update.microsoft.com/DownloadDialog.aspx?updateIDs=[{%22size%22%3A0%2C%22languages%22%3A%22%22%2C%22uidInfo%22%3A%22";
        string post2 = "%22%2C%22updateID%22%3A%22";
        string post3 = "%22}]&updateIDsBlockedForImport=&wsusApiPresent=&contentImport=&sku=&serverName=&ssl=&portNumber=&version=";
        string postText = post1 + kbGUID + post2 + kbGUID + post3;
        //Console.WriteLine(postText);
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
"@
}



#$KBNumber = "123456"
#$Product = "Windows 10"
#$downloadLink = [GetKBDLLink]::GetDownloadLink($KBNumber, $Product)
#Write-Host "Download link: $downloadLink"


        $KBLatest=''
$KBList=''
$KBItemsToShow = 6

        #Lastest hotfix for Windows Server from the respective KB pages
            $OSType=$OSCheck
            If($OSCheck -imatch '2008r2'-or $OSCheck -imatch '2008 r2'){
                $OSCheck ='2008 r2'
                # Download the HTML content
                $url = "https://support.microsoft.com/en-us/help/4009469"
                $webClient = New-Object System.Net.WebClient
                $htmlpage = $webClient.DownloadString($url)

                # Find all elements with the "supLeftNavLink" class
                $links=[regex]::Matches($htmlpage,'supLeftNavLink.*?(href=\".*?\")[^>]*>(.*?)(KB\d{7})(.*?)<\/a>')

                $KBList  = $Links[0..($KBItemsToShow-1)] | Select-Object -Property `
                    @{L='KBNumber';E={$_.Groups[3].Value}},`
                    @{L='Date';E={($_.Groups[2].Value -split "&#x2014;")[0]}},`
                    @{L="Description";E={($_.Groups[2].Value -replace "&#x2014;"," ")+$_.Groups[4].Value}},
                    @{L="OS Build";E={"6.1.7601"}},
                    @{L='InfoLink';E={"https://support.microsoft.com"+(($_.Groups[1].Value -split 'href="')[-1] -split '"')[0]}},
                    @{L="DownloadLink";E={""}}
            }

            If($OSCheck -imatch '2012r2'-or $OSCheck -imatch '2012 r2'){
                $OSCheck ='2012 r2'
                # Download the HTML content
                $url = "https://support.microsoft.com/en-us/help/4009470"
                $webClient = New-Object System.Net.WebClient
                $htmlpage = $webClient.DownloadString($url)

                # Find all elements with the "supLeftNavLink" class
                $links=[regex]::Matches($htmlpage,'supLeftNavLink.*?(href=\".*?\")[^>]*>(.*?)(KB\d{7})(.*?)<\/a>')

                $KBList  = $Links[0..($KBItemsToShow-1)] | Select-Object -Property `
                    @{L='KBNumber';E={$_.Groups[3].Value}},`
                    @{L='Date';E={($_.Groups[2].Value -split "&#x2014;")[0]}},`
                    @{L="Description";E={($_.Groups[2].Value -replace "&#x2014;"," ")+$_.Groups[4].Value}},
                    @{L="OS Build";E={"6.3.9600"}},
                    @{L='InfoLink';E={"https://support.microsoft.com"+(($_.Groups[1].Value -split 'href="')[-1] -split '"')[0]}},
                    @{L="DownloadLink";E={""}}
            }
    
            If($OSCheck -imatch '2016'){
                # Download the HTML content
                $url = "https://support.microsoft.com/en-us/help/4000825"
                $webClient = New-Object System.Net.WebClient
                $htmlpage = $webClient.DownloadString($url)

                # Find all elements with the "supLeftNavLink" class
                $links=[regex]::Matches($htmlpage,'supLeftNavLink.*?(href=\".*?\")>(.*?)(KB\d{7})\D+((?:(?!Preview).)14393.*?)\)(?:(?!Preview).)*<\/a>')

                $KBList  = $Links[0..($KBItemsToShow-1)] | Select-Object -Property `
                    @{L='KBNumber';E={$_.Groups[3].Value}},`
                    @{L='Date';E={($_.Groups[2].Value -replace "&#x2014;"," ")}},`
                    @{L="Description";E={($_.Groups[2].Value -replace "&#x2014;"," ")+$_.Groups[3].Value+$_.Groups[4].Value}},
                    @{L="OS Build";E={$_.Groups[4].Value.Trim()}},
                    @{L='InfoLink';E={"https://support.microsoft.com"+(($_.Groups[1].Value -split 'href="')[-1] -split '"')[0]}},
                    @{L="DownloadLink";E={""}}
            }

            If($OSCheck -imatch '2019'){
                # Download the HTML content
                $url = "https://support.microsoft.com/en-us/help/4464619/windows-10-update-history"
                $webClient = New-Object System.Net.WebClient
                $htmlpage = $webClient.DownloadString($url)

                # Find all elements with the "supLeftNavLink" class
                $links=[regex]::Matches($htmlpage,'supLeftNavLink.*?(href=\".*?\")>(.*?)(KB\d{7})\D+((?:(?!Preview).)17763.*?)\)(?:(?!Preview).)*<\/a>')

                $KBList  = $Links[0..($KBItemsToShow-1)] | Select-Object -Property `
                    @{L='KBNumber';E={$_.Groups[3].Value}},`
                    @{L='Date';E={($_.Groups[2].Value -replace "&#x2014;"," ")}},`
                    @{L="Description";E={($_.Groups[2].Value -replace "&#x2014;"," ")+$_.Groups[3].Value+$_.Groups[4].Value}},
                    @{L="OS Build";E={$_.Groups[4].Value.Trim()}},
                    @{L='InfoLink';E={"https://support.microsoft.com"+(($_.Groups[1].Value -split 'href="')[-1] -split '"')[0]}},
                    @{L="DownloadLink";E={""}}
            }

            If($OSCheck -imatch '22H2|21H2|20H2'){
                # Download the HTML content
                #$url = "https://support.microsoft.com/en-us/help/5018894"
                $url = "https://support.microsoft.com/en-us/topic/release-notes-for-azure-stack-hci-version-23h2-018b9b10-a75b-4ad7-b9d1-7755f81e5b0b"
                $webClient = New-Object System.Net.WebClient
                $htmlpage = $webClient.DownloadString($url)

                # Find all elements with the "supLeftNavLink" class
                $divs=[regex]::Matches($htmlpage,'(?s)supLeftNavCategory((?:.*?)(<\/div>)){2}')
                Foreach ($match in $divs) {
                    If ($match.Groups[1].Value -match $OSCheck) {
                        $links=[regex]::Matches($match.Groups[1].Value,'supLeftNavLink.*?(href=\".*?\")[^>]*>((?:(?!preview).)*?)(KB\d{7})(.*?)<\/a>')
                    }
                }

                $KBList  = $Links[0..($KBItemsToShow-1)] | Select-Object -Property `
                    @{L='KBNumber';E={$_.Groups[3].Value}},`
                    @{L='Date';E={($_.Groups[2].Value -split "\s")[0..2] -Join " "}},`
                    @{L="Description";E={($_.Groups[2].Value -replace "&#x2014;"," ")+$_.Groups[3].Value+$_.Groups[4].Value}},
                    @{L="OS Build";E={"20349"}},
                    @{L='InfoLink';E={"https://support.microsoft.com"+(($_.Groups[1].Value -split 'href="')[-1] -split '"')[0]}},
                    @{L="DownloadLink";E={""}}

            }
            If($OSCheck -imatch '2022'){
                # Download the HTML content
                $url = "https://support.microsoft.com/en-us/help/5005454"
                $webClient = New-Object System.Net.WebClient
                $htmlpage = $webClient.DownloadString($url)

                # Find all elements with the "supLeftNavLink" class
                $links=[regex]::Matches($htmlpage,'supLeftNavLink.*?(href=\".*?\")>(.*?)(KB\d{7})\D+((?:(?!Preview).)20348.*?)\)(?:(?!Preview).)*<\/a>')

                #Set OS Type to find 22H2 updates
                $OSType="22H2"

                $KBList  = $Links[0..($KBItemsToShow-1)] | Select-Object -Property `
                    @{L='KBNumber';E={$_.Groups[3].Value}},`
                    @{L='Date';E={($_.Groups[2].Value -replace "&#x2014;"," ")}},`
                    @{L="Description";E={($_.Groups[2].Value -replace "&#x2014;"," ")+$_.Groups[3].Value+$_.Groups[4].Value}},
                    @{L="OS Build";E={"2022"}},
                    @{L='InfoLink';E={"https://support.microsoft.com"+(($_.Groups[1].Value -split 'href="')[-1] -split '"')[0]}},
                    @{L="DownloadLink";E={""}}
            }


#$KBList = $KBList | ? DownloadLink -like "https*"
$KBLatest = $KBList[0]
$KBDLUriSource = ""
if ($KBLatest -and $KBLatest.KBNumber) {
    try { $KBDLUriSource = [GetKBDLLink]::GetDownloadLink($KBLatest.KBNumber,$OSType) }
    catch { Write-Host "    WARNING: Failed to resolve Microsoft Update download link for $($KBLatest.KBNumber)." -ForegroundColor Yellow }
}

        $dstop=Get-Date
        #Write-Host "Total time taken is $(($dstop-$dstart).totalmilliseconds)"
#endregion Recommended updates and hotfixes for Windows Server
            
                $WSLCU = New-Object -TypeName PSObject
                #Add-Member -InputObject $WSLCU -MemberType NoteProperty -Name Build -Value $Build
                $KBLastUpdated = "No Data Found"
                if ($KBLatest -and $KBLatest.Date) {
                    try {
                        $KBLastUpdated = ([DateTime]$KBLatest.Date).ToString("yyyy-MM-dd")
                    }
                    catch {
                        $KBLastUpdated = [string]$KBLatest.Date
                    }
                }

                Add-Member -InputObject $WSLCU -MemberType NoteProperty -Name KBNumber -Value $KBLatest.KBNumber
                Add-Member -InputObject $WSLCU -MemberType NoteProperty -Name LastUpdated -Value $KBLastUpdated
                Add-Member -InputObject $WSLCU -MemberType NoteProperty -Name Title -Value $KBLatest.Description
                Add-Member -InputObject $WSLCU -MemberType NoteProperty -Name URL -Value $KBDLUriSource
                Add-Member -InputObject $WSLCU -MemberType NoteProperty -Name Details -Value $KBLatest.InfoLink
                

                
          $WSLCUout=$WSLCU | Select-object @{Label="ServiceTag";Expression={"$ServiceTag"}}`
                ,@{Label="PowerEdge";Expression={"$ServerType"}}`
                ,@{Label="OS";Expression={"$InstalledOS"+" "+"$OSVersion"}}`
                ,@{Label="Type";Expression={"OS"}}`
                ,@{Label="Category";Expression={"Microsoft Update"}}`
                ,@{Label="Name";Expression={$WSLCU.Title}}`
                ,@{Label="InstalledVersion";Expression={"NA"}}`
                ,@{Label="AvailableVersion";Expression={$WSLCU.KBNumber}}`
                ,@{Label="CatalogInfo";Expression={"catalog.update.microsoft.com"}}`
                ,@{Label="Criticality";Expression={"Not Available"}}`
                ,@{Label="ReleaseDate";Expression={$WSLCU.LastUpdated}}`
                ,@{Label="URL";Expression={$WSLCU.URL}}`
                ,@{Label="Details";Expression={$WSLCU.Details}}`
                | sort-object Category,Firmware,Name
            $allArray+= $WSLCUout
        }
    }  
    IF($CluChkMode -eq "YES"){
    # Switch port to Host map
        Write-host "Checking for Switch Port Connection Info...." 
        $SwPort2HostMap=@()
        $SwitchPortsFromiDRAC=$DCIM_VIEM_Properties|Where-Object{$_.SwitchPortConnectionID -ne $NULL -and $_.SwitchConnectionID -match ":"}
        ForEach($Inst in $SwitchPortsFromiDRAC){
            #Filter out iDRAC becuase we only want to see host NICs
            IF($INST.InstanceID -inotmatch 'iDRAC'){
                $SwPort2HostMap+=$inst|Select-Object @{Label="HostName";Expression={(Split-Path -Path $HostName -Leaf).Split(".")[0]}},@{Label='SwitchMacAddress';Expression={($_.SwitchConnectionID -split '(?<=(?i)[0-9a-f]{4})')[0].substring(0,17) }},@{L="SwitchPortConnectionID";E={$SwitchPortConnectionID=($_.SwitchPortConnectionID -split '(?<=/[0-9A]+e)')[0];$SwitchPortConnectionID.substring(0,$SwitchPortConnectionID.length - 1)}},@{Label='HostNicSlotPort';Expression={$HNSP=($_.InstanceID -split '(?<=-[0-9A]N)')[0];$HNSP.substring(0,$HNSP.length - 1)}},@{Label='HostNICMacAddress';Expression={(($DCIM_VIEM_Properties|Where-Object{$_.PermanentMACAddress -ne $Null}|Where-Object{$_.InstanceID.startswith($Inst.InstanceID)}).PermanentMACAddress -split '(?<=(?i)[0-9a-f]{4})')[0].substring(0,17)}}
            }
        }
        IF($SwPort2HostMap){
            Write-Host ""
            #$SwPort2HostMap | FT
            $SwPort2HostMapAll+=$SwPort2HostMap
        }
            
    #$BIOSandNICCFG=@()
    #BIOS and iDRAC configuration recommendations for servers in a Dell EMC Ready Solution for Microsoft WSSD 
        #From <https://www.dell.com/support/article/us/en/19/sln313842/bios-and-idrac-configuration-recommendations-for-servers-in-a-dell-emc-ready-solution-for-microsoft-wssd?lang=en> 
        
        
        If(($IsNewS2DCatalog="YES") -or ($ServerType -Match "XD")){
            $BIOSandiDRACCfg=@()
            $BIOSandiDRACCfgTable=@()
            Write-host "BIOS and iDRAC configuration recommendations for WSSD...."
            Write-host "    Reference link: https://www.dell.com/support/kbdoc/en-us/000135856/bios-and-idrac-configuration-recommendations-for-servers-in-a-dell-emc-solutions-for-microsoft-azure-stack-hci"    
                $BIOSandiDRACCfgTable+="Memory Settings,Node Interleaving,Disabled,BIOS.Setup.1-1,NodeInterleave,R640 R740XD,DCIM_BIOSEnumeration"

                $BIOSandiDRACCfgTable+="Processor Settings,Logical Processor,Enabled,BIOS.Setup.1-1,LogicalProc,R640 R740XD,DCIM_BIOSEnumeration"
                $BIOSandiDRACCfgTable+="Processor Settings,Virtualization Technology,Enabled,BIOS.Setup.1-1,ProcVirtualization,R640 R740XD R750 R650 R6515 R7525,DCIM_BIOSEnumeration"
                $BIOSandiDRACCfgTable+="Processor Settings,DCU Streamer Prefetcher,Enabled,BIOS.Setup.1-1,DcuStreamerPrefetcher,R640 R740XD,DCIM_BIOSEnumeration"
                $BIOSandiDRACCfgTable+="Processor Settings,DCU IP Prefetcher,Enabled,BIOS.Setup.1-1,DcuIpPrefetcher,R640 R740XD,DCIM_BIOSEnumeration"
                $BIOSandiDRACCfgTable+="Processor Settings,UPI Prefetcher,Enabled,BIOS.Setup.1-1,UpiPrefetch,R640 R740XD,DCIM_BIOSEnumeration"
                $BIOSandiDRACCfgTable+="Processor Settings,Sub NUMA Cluster,Disabled,BIOS.Setup.1-1,SubNumaCluster,R640 R740XD,DCIM_BIOSEnumeration"
                $BIOSandiDRACCfgTable+="Processor Settings,X2 APIC Mode,Enabled,BIOS.Setup.1-1,ProcX2Apic,R640 R740XD R750 R650 R6515 R7525,DCIM_BIOSEnumeration"
                $BIOSandiDRACCfgTable+="Processor Settings,Dell Controlled Turbo,Enabled,BIOS.Setup.1-1,ControlledTurbo,R640 R740XD,DCIM_BIOSEnumeration"
                $BIOSandiDRACCfgTable+="Processor Settings,Kernel DMA Protection,Enabled,BIOS.Setup.1-1,KernelDmaProtection,R650 R750 R7525,DCIM_BIOSEnumeration"
                
                $BIOSandiDRACCfgTable+="SATA Settings,Embedded SATA,AHCIMode,BIOS.Setup.1-1,EmbSata,R640 R740XD,DCIM_BIOSEnumeration"
                $BIOSandiDRACCfgTable+="SATA Settings,Security Freeze Lock,Enabled,BIOS.Setup.1-1,SecurityFreezeLock,R640 R740XD,DCIM_BIOSEnumeration"
                $BIOSandiDRACCfgTable+="SATA Settings,Write Cache,Disabled,BIOS.Setup.1-1,WriteCache,R640 R740XD,DCIM_BIOSEnumeration"
                
                $BIOSandiDRACCfgTable+="NVMe Settings,NVMe Mode,NonRAID,BIOS.Setup.1-1,NvmeMode,R640 R740XD,DCIM_BIOSEnumeration"
                
                $BIOSandiDRACCfgTable+="Boot Settings,Boot Mode,UEFI,BIOS.Setup.1-1,BootMode,R640 R740XD,DCIM_BIOSEnumeration"   
                $BIOSandiDRACCfgTable+="Boot Settings,Boot Sequence Retry,Enabled,BIOS.Setup.1-1,BootSeqRetry,R640 R740XD,DCIM_BIOSEnumeration"

                $BIOSandiDRACCfgTable+="Integrated Devices,SR-IOV Global Enable,Enabled,BIOS.Setup.1-1,SriovGlobalEnable,R640 R740XD R750 R650 R6515 R7525,DCIM_BIOSEnumeration"

                #$BIOSandiDRACCfgTable+="System Profile Settings,System Profile,Custom,BIOS.Setup.1-1,SysProfile,R640 R740XD,DCIM_BIOSEnumeration"
                $BIOSandiDRACCfgTable+="System Profile Settings,System Profile,Performance PerfOptimized,BIOS.Setup.1-1,SysProfile,R640 R740XD2 R740XD R750 R650 R6515 R7525,DCIM_BIOSEnumeration"
                #$BIOSandiDRACCfgTable+="System Profile Settings,System Profile,PerfOptimized,BIOS.Setup.1-1,SysProfile,R750 R650 R6515 R7525,DCIM_BIOSEnumeration"
                $BIOSandiDRACCfgTable+="System Profile Settings,CPU Power Management,MaxPerf,BIOS.Setup.1-1,ProcPwrPerf,R740XD R640,DCIM_BIOSEnumeration"
                $BIOSandiDRACCfgTable+="System Profile Settings,Memory Frequency,MaxPerf,BIOS.Setup.1-1,MemFrequency,R740XD R640,DCIM_BIOSEnumeration"
                $BIOSandiDRACCfgTable+="System Profile Settings,Turbo Boost,Enabled,BIOS.Setup.1-1,ProcTurboMode,R740XD2 R740XD R640 YES,DCIM_BIOSEnumeration"
                $BIOSandiDRACCfgTable+="System Profile Settings,C-States,Disabled,BIOS.Setup.1-1,ProcCStates,R740XD R640,DCIM_BIOSEnumeration"
                $BIOSandiDRACCfgTable+="System Profile Settings,C1E,Disabled,BIOS.Setup.1-1,ProcC1E,R740XD R640,DCIM_BIOSEnumeration"
                $BIOSandiDRACCfgTable+="System Profile Settings,Memory Patrol Scrub,Standard,BIOS.Setup.1-1,MemPatrolScrub,R740XD R640,DCIM_BIOSEnumeration"

                $BIOSandiDRACCfgTable+="System Security,TPM Security,On,BIOS.Setup.1-1,TpmSecurity,R640 R740XD R750 R650 R6515 R7525,DCIM_BIOSEnumeration"
                $BIOSandiDRACCfgTable+="System Security,Intel TXT,Off,BIOS.Setup.1-1,IntelTxt,R740XD R640,DCIM_BIOSEnumeration"
                $BIOSandiDRACCfgTable+="System Security,Intel TXT,On,BIOS.Setup.1-1,IntelTxt,R650 R750 R7525,DCIM_BIOSEnumeration"
                $BIOSandiDRACCfgTable+="System Security,AC Power Recovery,On,BIOS.Setup.1-1,AcPwrRcvry,R640 R740XD R750 R650 R6515 R7525,DCIM_BIOSEnumeration"
                $BIOSandiDRACCfgTable+="System Security,AC Power Recovery Delay,Random,BIOS.Setup.1-1,AcPwrRcvryDelay,R640 R740XD R750 R650 R6515 R7525,DCIM_BIOSEnumeration"
                $BIOSandiDRACCfgTable+="System Security,Secure Boot,Enabled,BIOS.Setup.1-1,SecureBoot,R640 R740XD R750 R650 R6515 R7525,DCIM_BIOSEnumeration"
                $BIOSandiDRACCfgTable+="System Security,Secure Boot Policy,Standard,BIOS.Setup.1-1,SecureBootPolicy,R740XD R640,DCIM_BIOSEnumeration"

                $BIOSandiDRACCfgTable+="TPM Advanced Settings,    TPM PPI Bypass Provision,Enabled,BIOS.Setup.1-1,TpmPpiBypassProvision,R750 R650 R7525,DCIM_BIOSEnumeration"
                $BIOSandiDRACCfgTable+="TPM Advanced Settings,    TPM PPI Bypass Clear,Enabled,BIOS.Setup.1-1,TpmPpiBypassClear,R750 R650 R7525,DCIM_BIOSEnumeration"
                $BIOSandiDRACCfgTable+="TPM Advanced Settings,    TPM2 Algorithm Selection,SHA256,BIOS.Setup.1-1,Tpm2Algorithm,R750 R650 R7525,DCIM_BIOSEnumeration"

                $BIOSandiDRACCfgTable+="Power Configuration,Redundancy Policy,A/B Grid Redundant,System.Embedded.1:ServerPwr.1,PSRedPolicy,R640 R740XD R750 R650 R6515 R7525,DCIM_SystemEnumeration"
                $BIOSandiDRACCfgTable+="Power Configuration,Enable Hot Spare,Enabled,System.Embedded.1:ServerPwr.1,PSRapidOn,R640 R740XD R750 R650 R6515 R7525,DCIM_SystemEnumeration"
                $BIOSandiDRACCfgTable+="Power Configuration,Primary Power Supply Unit,PSU1,System.Embedded.1:ServerPwr.1,RapidOnPrimaryPSU,R640 R740XD R750 R650 R6515 R7525,DCIM_SystemEnumeration"

                $BIOSandiDRACCfgTable+="Network Settings,Enable NIC,Enabled,iDRAC.Embedded.1:CurrentNIC.1,Enabled,R640 R740XD R750 R650 R6515 R7525,DCIM_iDRACCardEnumeration"
                $BIOSandiDRACCfgTable+="Network Settings,NIC Selection,Dedicated,iDRAC.Embedded.1:CurrentNIC.1,Selection,R640 R740XD R750 R650 R6515 R7525,DCIM_iDRACCardEnumeration"

                ForEach($Line In $BIOSandiDRACCfgTable){
                    $Item=@()
                    $Item=$Line -split ","
                    IF($Item[5] -split " " -ieq $ServerType){
                        $BIOSandiDRACCfgLookup=(($CIM_BIOSAttribute_Instances |`
                                                Where-Object{($_.CLASSNAME -match $Item[6])}).PROPERTY|`
                                                Where-Object{$_.VALUE -eq $Item[4]}).ParentNode.'PROPERTY.ARRAY' |`
                                                Where-Object{$_.Name -match "CurrentValue"} |`
                                                Select-Object @{Label="HostName";Expression={(Split-Path -Path $HostName -Leaf).Split(".")[0]}}`
                                                ,@{Label="ServiceTag";Expression={"$ServiceTag"}}`
                                                ,@{Label="PowerEdge";Expression={"$ServerType"}}`
                                                ,@{Label="OS";Expression={"$InstalledOS"+" "+"$OSVersion"}}`
                                                ,@{Label="Type";Expression={"BIOS Config"}}`
                                                ,@{Label="Setting Category";Expression={$Item[0]}}`
                                                ,@{Label="Setting Name";Expression={$Item[1]}}`
                                                ,@{Label="CurrentValue";Expression={IF($Item[2] -notmatch $_.'VALUE.ARRAY'.VALUE){"***"+$_.'VALUE.ARRAY'.VALUE}Else{$_.'VALUE.ARRAY'.VALUE}}}`
                                                ,@{Label="DesiredValue";Expression={$Item[2]}}|`
                                                sort-object Type,Category,Name
                            $BIOSandiDRACCfg+=$BIOSandiDRACCfgLookup
                        }
                    }

                    #$BIOSandiDRACCfg | Format-Table 
                    $BIOSandNICCFG+=$BIOSandiDRACCfg  
        }
        IF($IsAZHub){
        # Azure Stack HUB Bios setting
            $BIOSCfg=@()
            Write-host "BIOS and iDRAC configuration for Azure Stack Hub...."
            # Setting JSON
            $AZHURL="https://solutions.one.dell.com/sites/NAEnterprise/SST/Communities/DRiFT/DRiFT%20Docs/AZHubSettings.txt"
            $AZHDownloadFile="$env:TEMP\AZHubSetting.json"
            IF(-not($AZHSettings)){
            $AZHSettings=Invoke-WebRequest -Uri $AZHURL -OutFile $AZHDownloadFile -UseDefaultCredentials}
            $HubSetting=Get-Content -Path $AZHDownloadFile  | ConvertFrom-Json
            Remove-Item -Path $AZHDownloadFile -Force
            $AZBIOSSETTINGS=($HubSetting."PowerEdge $ServerType").BIOS_SETTINGS
            $AZBIOSSETTINGS.PSObject.Properties|ForEach{
                $AZBItem=$_
                $BIOSCfgLookup=(($CIM_BIOSAttribute_Instances |`
                    Where-Object{($_.CLASSNAME -match 'DCIM_BIOSEnumeration')}).PROPERTY|`
                    Where-Object{$_.VALUE -eq $AZBItem.Name}).ParentNode.'PROPERTY.ARRAY' |`
                    Where-Object{$_.Name -match "CurrentValue"} |`
                    Select-Object @{Label="HostName";Expression={(Split-Path -Path $HostName -Leaf).Split(".")[0]}}`
                    ,@{Label="ServiceTag";Expression={"$ServiceTag"}}`
                    ,@{Label="PowerEdge";Expression={"$ServerType"}}`
                    ,@{Label="OS";Expression={"$InstalledOS"+" "+"$OSVersion"}}`
                    ,@{Label="Type";Expression={"BIOS Config"}}`
                    ,@{Label="Setting Category";Expression={($_.ParentNode.PROPERTY.Name -eq 'AttributeDisplayName').value}}`
                    ,@{Label="Setting Name";Expression={$AZBItem.Name}}`
                    ,@{Label="CurrentValue";Expression={IF($_.'VALUE.ARRAY'.VALUE -notmatch $AZBItem.Value){"***"+$_.'VALUE.ARRAY'.VALUE}Else{$_.'VALUE.ARRAY'.VALUE}}}`
                    ,@{Label="DesiredValue";Expression={$AZBItem.Value}}|`
                    sort-object Type,Category,Name
                $BIOSCfg+=$BIOSCfgLookup
            }
            #$BIOSCfg|FT
            $BIOSandNICCFG+=$BIOSCfg
            $iDRACCfg=@()
            $iDRACCfgLookup=@()
            $AZiDRACSETTINGS=($HubSetting."PowerEdge $ServerType").iDRAC_SETTINGS
            $AZiDRACSETTINGS.PSObject.Properties|ForEach{
                $AZiItem=$_
                $iDRACCfgLookup=($CIM_BIOSAttribute.CIM.MESSAGE.SIMPLEREQ."VALUE.NAMEDINSTANCE" |`
                    Where-Object{$_.INSTANCENAME.KEYBINDING.KEYVALUE.'#text' -imatch ($AZiItem.Name)}).INSTANCE.'PROPERTY.ARRAY'|`
                    Where-Object{$_.Name -match "CurrentValue"} |`
                    Select-Object @{Label="HostName";Expression={(Split-Path -Path $HostName -Leaf).Split(".")[0]}}`
                    ,@{Label="ServiceTag";Expression={"$ServiceTag"}}`
                    ,@{Label="PowerEdge";Expression={"$ServerType"}}`
                    ,@{Label="OS";Expression={"$InstalledOS"+" "+"$OSVersion"}}`
                    ,@{Label="Type";Expression={"iDRAC Config"}}`
                    ,@{Label="Setting Category";Expression={(($AZiItem.Name -split '\#')[1])}}`
                    ,@{Label="Setting Name";Expression={(($AZiItem.Name -split '\#')[2])}}`
                    ,@{Label="CurrentValue";Expression={IF($_.'VALUE.ARRAY'.VALUE -notmatch $AZiItem.Value){"***"+$_.'VALUE.ARRAY'.VALUE}Else{$_.'VALUE.ARRAY'.VALUE}}}`
                    ,@{Label="DesiredValue";Expression={$AZiItem.Value}}|`
                    sort-object Type,Category,Name
                $iDRACCfg+=$iDRACCfgLookup
            }
            #$iDRACCfg | FT    
        }
        $BIOSandNICCFG+=$iDRACCfg
        #QLogic NIC configuration 
        If(($IsNewS2DCatalog="YES")-and($NICDrv.name -imatch "QLogic")){
            $QLogicNicCfg=@()
            $QLogicNicCfgTable=@()
            Write-host "Checking BIOS QLogic NIC configuration for WSSD...." 
                $QLogicNicCfgTable+="BIOS QLogic NIC Config,Boot Protocol,None,NIC.Slot.,FWBootProtocol"
                $QLogicNicCfgTable+="BIOS QLogic NIC Config,Link Speed,QLGC_SmartAN,NIC.Slot.,LnkSpeed"
                $QLogicNicCfgTable+="BIOS QLogic NIC Config,DCBX Protocol,QLGC_Disabled,NIC.Slot.,QLGC_DCBXProtocol"
                $QLogicNicCfgTable+="BIOS QLogic NIC Config,RDMA Operational Mode,QLGC_iWARP,NIC.Slot.,QLGC_RDMAOperationalMode"
                $QLogicNicCfgTable+="BIOS QLogic NIC Config,NIC + RDMA Mode,Enabled,NIC.Slot.,RDMANICModeOnPort"
                $QLogicNicCfgTable+="BIOS QLogic NIC Config,Virtualization Mode,NONE,NIC.Slot.,VirtualizationMode"
                $QLogicNicCfgTable+="BIOS QLogic NIC Config,Virtual LAN Mode,Disabled,NIC.Slot.,VLanMode"
                ForEach($Line In $QLogicNicCfgTable){
                    $Item=@()
                    $Item=$Line -split ","
                        # Filter for QLogic RDMA NICs
                        $NICSLOT=((($CIM_BIOSAttribute_Instances |`
                                                Where-Object{($_.CLASSNAME -match "DCIM_NICEnumeration")}).PROPERTY|`
                                                Where-Object{$_.VALUE -imatch "QLGC_RDMAOperationalModePort"}).ParentNode.PROPERTY|`
                                                Where-Object{($_.Name -match "FQDD")}).VALUE
                                                Where-Object{($_.Name -match "CurrentValue")}|Select-Object VALUE
                                                #Where-Object{[regex]::Matches($_.VALUE,"\d-\d-\d:QLGC_RDMAOperationalModePort")}).Value -replace ":QLGC_RDMAOperationalModePort"

                        ForEach($SNIC in $NICSLOT){
                            $QLogicNicCfgLookup=((($CIM_BIOSAttribute_Instances |`
                                                Where-Object{($_.CLASSNAME -match "DCIM_NICEnumeration")}).PROPERTY|`
                                                Where-Object{$_.VALUE -imatch $SNIC}).ParentNode.PROPERTY|`
                                                Where-Object{$_.VALUE -ieq $Item[4]}).ParentNode.'PROPERTY.ARRAY'|`
                                                Where-Object{$_.Name -match "CurrentValue"}|`
                                                Select-Object @{Label="HostName";Expression={(Split-Path -Path $HostName -Leaf).Split(".")[0]}}`
                                                ,@{Label="ServiceTag";Expression={"$ServiceTag"}}`
                                                ,@{Label="PowerEdge";Expression={"$ServerType"}}`
                                                ,@{Label="OS";Expression={"$InstalledOS"+" "+"$OSVersion"}}`
                                                ,@{Label="Type";Expression={"BIOS Config"}}`
                                                ,@{Label="Setting Category";Expression={$Item[0]}}`
                                                ,@{Label="Device";Expression={$SNIC}}`
                                                ,@{Label="Setting Name";Expression={$Item[1]}}`
                                                ,@{Label="CurrentValue";Expression={IF($_.'VALUE.ARRAY'.VALUE -ne $Item[2]){"***"+$_.'VALUE.ARRAY'.VALUE}Else{$_.'VALUE.ARRAY'.VALUE}}}`
                                                ,@{Label="DesiredValue";Expression={$Item[2]}}|`
                                                sort-object Type,Category,Name
                                                $QLogicNicCfg+=$QLogicNicCfgLookup
                                                }
                             #$QLogicNicCfgLookup | Format-Table 
                        }
                    $QLogicNicCfg=$QLogicNicCfg | Sort-Object Device,'Setting Name' -Unique
                    #$QLogicNicCfg | Format-Table  
                    $BIOSandNICCFG+=$QLogicNicCfg
        }
        }
        
        
 #Links to download button
    ForEach($aitem in $allArray){
        If(($aitem.InstalledVersion -match [Regex]::Escape("***"))`
        -or($aitem.InstalledVersion -match "NA")`
        -or($aitem.InstalledVersion -match "Not Available")`
        -or($aitem.InstalledVersion -match "Not Applicable")){
            $Files2Download+="'"+$aitem.URL+"',"}
    }
    $Files2Download=($Files2Download | sort-object | Get-Unique) 
 
 #Add seperator
     IF($SkipDriversandFirmware -eq "NO"){
         $Folder_Count=$DriFTFolders.GetDirectories | Measure-Object | ForEach-Object{$_.Count}
         IF ($Folder_Count -gt 1){
                $ReportSeperator=@()
                $ReportSeperator = [PSCustomObject]@{
                            ServiceTag=$ServiceTag
                            PowerEdge=""
                            OS=""
                            Type=""
                            Category=""
                            Name=""
                            InstalledVersion=""
                            AvailableVersion=""
                            Criticality=""
                            ReleaseDate=""
                            URL=""
                            Details=""}
                $allArray+=$ReportSeperator
         }
     }
     # Filter for unique URLs and removes non http URLs
     $allArrayout+=$allArray|sort-object URl -Unique|Select-Object ServiceTag,PowerEdge,OS,Type,Category,Name,InstalledVersion,AvailableVersion,CatalogInfo,Criticality,ReleaseDate,`
        @{Label="URL";Expression={IF($_.URL -imatch "http"){$_.URL}Else{""}}},Details
     $allArray=@()
}

<# Upload report data to Azure
ForEach($row in $allArrayout){
    $RowData=$row|Select-Object @{Label="ReportID";Expression={"$DReportID"}},PowerEdge,OS,Type,Category,Name,InstalledVersion,AvailableVersion,CatalogInfo,Criticality,ReleaseDate,URL,Details
    
    add-TableData -TableName "DriFTReportData" -PartitionKey "DriFT" -RowKey (new-guid).guid -data $RowData -sasWriteToken '?SECRET REMOVED'
    }
#>

$DateTime=Get-date
$DTString=Get-Date -Format "yyyyMMdd_HHmmss_"
$Title= "Latest Dell PowerEdge Firmware and Drivers"
#Write-host "OS:" $InstalledOS
#Displays the results

#Output location to the same place as the input file
$SourcePath=@()
If($TSRLoc.count -gt 1){
    $SourcePath=([regex]::match($TSRLoc[0],'(.*\\).*')).Groups[1].value
}Else{$SourcePath=([regex]::match($TSRLoc,'(.*\\).*')).Groups[1].value}

#HTML Out
If($OutputType -match "HTML") {
    $HTAOut = $SourcePath + $DriFTVer + "_" + $DTString + $ServiceTagList + ".html"
    $HTAOut = $HTAOut.Replace("*", "")
    IF ($HTAOut.Length -gt 248){
        $HTAOut = $SourcePath + $DriFTVer + "_" + $DTString + ".html"
    }

    Write-Host "Report Output location: " $HTAOut

    # Normalize report rows before HTML generation. 17G support can feed this same shape later.
    $ReportRows = @()
    $ReportRows += $allArrayout | Select-Object ServiceTag,PowerEdge,OS,Type,Category,Name,InstalledVersion,AvailableVersion,CatalogInfo,Criticality,ReleaseDate,URL,Details,SourceType

    Write-DriFTReport `
        -ReportRows $ReportRows `
        -Path $HTAOut `
        -DriFTVersion $DriFTVer `
        -DisplayVersion $DFTV `
        -ReportDate $DateTime `
        -ServerType $ServerType `
        -S2DCatalogNeeded $S2DCatalogNeeded `
        -NoneSupportedDevices $NoneSupportedDevices `
        -CatVerInfo $CatVerInfo | Out-Null

    IF($SkipDriversandFirmware -eq "NO"){
        If($OutputType -ne "NO"){
            if (Test-Path $HTAOut) { Invoke-Item($HTAOut) } else { Write-Host "    WARNING: HTML report was not created: $HTAOut" -ForegroundColor Yellow }

            IF($IsAZHub -eq $True){
                $NewReportView = New-DriFTCompareView -ReportRows $ReportRows
                $DriFTCSVOut = $NewReportView | Where-Object{$_.type -iNotMatch 'System.__ComObject'} | Sort-Object Type,Name | Select-object -Property * -Exclude RowError, RowState, Table, ItemArray, HasErrors | ConvertTo-Csv
                Out-File $SourcePath+$DriFTVer+"_"+$DTString+".csv" -InputObject $DriFTCSVOut
            }

            # Export BIOSandNICCFG to XML
            If($BIOSandNICCFG.length -gt 0){
                $BIOSandNICCFGOutPutPath = $SourcePath + "\" + $FileNameGuid + "_BIOSandNICCFG.xml"
                Write-Host "BIOS and iDRAC configuration output to: $BIOSandNICCFGOutPutPath"
                Do{$BIOSandNICCFG+$SwPort2HostMapAll+$MBSelLogWarnERROut | Export-Clixml -Path "$BIOSandNICCFGOutPutPath"}
                Until(Test-Path "$BIOSandNICCFGOutPutPath" -PathType Leaf)
            }
        }
    }
}
Write-Host " "
IF(!($CluChkMode -imatch "YES")){
    If(!($args)){
        $Run=Read-Host "Would you like to process another? [y/n]"
    }
}Else{$Run="n"}
If($Run -notmatch "y"){EndScript}
$allArrayout=@()
$allArray=@()
$ServiceTagList=@()
}While($Run -eq "y")

#Variable Cleanup
#Remove-Variable * -ErrorAction SilentlyContinue

# Cleanup files
IF(Test-Path $DownloadFile){Remove-Item -Path $DownloadFile}
IF(Test-Path $ExtracLoc){
    try { Remove-Item $ExtracLoc -Recurse -Force -ErrorAction Stop }
    catch { Write-Host "    WARNING: Failed to remove temp workspace $ExtracLoc. It can be deleted manually." -ForegroundColor Yellow }
}
IF(Test-Path (Join-Path (Join-Path $env:TEMP "DriFT") "Redfish")){
    try { Remove-Item (Join-Path (Join-Path $env:TEMP "DriFT") "Redfish") -Recurse -Force -ErrorAction SilentlyContinue } catch {}
}
}