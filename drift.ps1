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
    2025/01/24:v1.76 -  1. Bug Fix: TP - Fixed MS latest updates by copying and converted it from CluChk

    2024/04/03:v1.73 -  1. New Feature: TP - Added 15g S2d BIOS settings
			            2. New Feature: JG - Moved to GitHub
                        3. New Feature: JG - Added Function Invoke-RunDriFT 

    2022/06/27:v1.72 -  1. Bug Fix: JG - Resolved missing VMWare drivers when we have 7.0.X as .X does not matter.
                        2. Bug Fix: JG - Added a check if SEL does not exsist then display message SEL not found.

    2022/06/26:v1.71 -  1. Bug Fix: JG - Fixed issue displaying updated driver information on Azure Stack HCI-less Windows Servers.

    2022/05/01:v1.70 -  1. New Featrue: JG - Added expanded telemetry data

    2022/02/25:v1.69 -  1. Bug Fix: Resolved AZHCI Catalog vs Catalog Supported OS conflict

    2022/02/xx:v1.68 -  1. Bug Fix: Resolved wrong Documentation link for CPLD

    2021/12/07:v1.67 -  1. Update: Updated the links to the Windows update RSS feed
                        2. New Feature: Removed duplacate webpage is scraps

    2021/11/04:v1.66 -  1. Bug Fix: Resolved 403 errors on report upload

    2021/11/02:v1.65 -  1. Bug Fix: Resolved Telemetry data

    2021/10/xx:v1.64 -  1. Bug Fix: Resolved missing CPLD due to source webpage format change
                        2. New Feature: Added supported OS to driver filter
                        3. Bug Fix: Removed dumps from Switch port to Host map

    2020/08/xx:v1.63 -  1. Bug Fix: Removed extraneous output for new table format
                        2. New Feature: Added support for Precision 7910/20
                        3. New Feature: Moved source code to Azure
                        4. New Feature: Moved telemetry to Azure Tables
                        5. New Feature: Added Report Data to Azure Tables 
                        6. New Feature: Removed doanloaded and extracted files
                        7. New Feature: Do not show emplty reports

    2020/02/xx:v1.62 -  1. Bug Fix: Removed all Alias references
                        2. Add Feature: Added multi file commandline process via -input comma delimited list 
                        3. Bug Fix: Fixed missing Microsoft Update due to ATOM Feed Changes
                        4. New Feature: Added Azure Stack Hub support
                        5. New Feature: New multi node reporting view for easy node comparison
                        6. New Feature: Added SEL log Error/Warning for the last 30 days

    2020/01/28:v1.61 -  1. Bug Fix: Resolved failing CPLD details lookup
                        2. Bug Fix: Removed allways use AZCHI catalog.xml
                        3. New Feature: Add support of new iDRAC 4.40

    2020/01/08:v1.60 -  1. Bug Fix: Add XR2 = R440
                        2. New Feature: Added Memory Settings,Node Interleaving,Disabled
                        3. New Feature: Added R740XD2 to System Profile Settings,Turbo Boost,Enabled
                        4. New Feature: Added System Security,TPM Security,On
                        5. New Feature: Added Power Configuration,Redundancy Policy,Redundant
                        6. New Feature: Added Power Configuration,Enable Hot Spare,Enabled
                        7. New Feature: Added Power Configuration,Primary Power Supply Unit,PSU1
                        8. New Feature: Added Network Settings,Enable NIC,Enabled
                        9. New Feature: Added Network Settings,NIC Selection,Dedicated
                        10. New Feature: Added CPLD updates for S2D AX/Ready Nodes
                        11. New Feature: Moved Switch port to Host map to CluChk mode

    
    
    See older version for previous notes
#>
Function Invoke-RunDriFT{
# logging
$DateTime=Get-Date -Format yyyyMMdd_HHmmss;Start-Transcript -NoClobber -Path "C:\programdata\Dell\DriFT\DriFT_$DateTime.log"
Write-host "Starting log: C:\programdata\Dell\DriFT\DriFT_$DateTime.log"
IF(!($args)){
    #Variable Cleanup
    Remove-Variable * -ErrorAction SilentlyContinue
}
[system.gc]::Collect()
$DriFTVer="DriFT_v1.76"
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
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog -Property @{Multiselect = $true}
    $OpenFileDialog.Title = "Please Select One or More SupportAssist File(s)."
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "ZIP (*.zip)| *.zip"
    $OpenFileDialog.ShowDialog((New-Object System.Windows.Forms.Form -Property @{TopMost = $true })) | Out-Null
    $OpenFileDialog.filenames
}
<#If (-not $args) {
        Write-Host "DriFT running in Auto CluChk mode..."
        Write-Host "    $args"
        $FileNameGuid=New-Guid
        #$FileNameGuid=$args -replace '-cluchk ',""
        Write-Host "File Name Guid:" $FileNameGuid
        Write-Host "Processing TRS File(s)"
        #$TSRInputFiles=@()
        #$TSRInputFiles=(($args -split '-input')[1].trim() -split ',').trim()
        $TSRInputFiles=Get-FileName($env:USERPROFILE)
        $TSRLoc=$TSRInputFiles
        $TSRLoc
        $args=""
        $CluChkMode="YES"

}#>
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
                [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
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
$ExtracLoc="$env:TEMP\DriFT"
if (!(Test-Path $ExtracLoc -PathType Container)) {New-Item -ItemType Directory -Force -Path $ExtracLoc}

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
#import the XML
Write-host "Importing Catalog.xml...."
$CatalogXMLData = [Xml] (Get-Content "$ExtracLoc\Catalog.xml")
Write-host "Filtering Catalog.xml for latest PowerEdge Firmware and Drivers...."
$allArray=@()
$Files2Download=@()
$IsNewS2DCatalog="YES" #Do not change this to No Jim. :)
$SwPort2HostMapAll=@()

Do{
# Telemetry Information
    #~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Write-Host "Logging Telemetry Information..."
    function add-TableData {
        [CmdletBinding()] 
            param(
                [Parameter(Mandatory = $true)]
                [string] $tableName,

                [Parameter(Mandatory = $true)]
                [string] $PartitionKey,

                [Parameter(Mandatory = $true)]
                [string] $RowKey,

                [Parameter(Mandatory = $true)]
                [array] $data,
                
                [Parameter(Mandatory = $true)]
                [array] $sasWriteToken
            )
            $storageAccount = "gsetools"
            #$tableName = "DriftTelemetryData"

            # Need write access 
            #$sasWriteToken = "?sv=2017-04-17&si=Update&tn=DriftTelemetryData&sig=wrOnGPi/vI3g62CZyj4CPiJHILxopLuNLaMY/nk2idA%3D"

            $resource = "$tableName(PartitionKey='$PartitionKey',RowKey='$Rowkey')"

            # should use $resource, not $tableNmae
            $tableUri = "https://$storageAccount.table.core.windows.net/$resource$sasWriteToken"

            # should be headers, because you use headers in Invoke-RestMethod
            $headers = @{
                Accept = 'application/json;odata=nometadata'
            }

            $body = $data | ConvertTo-Json
            #This adds and updates the table record
            $item = Invoke-RestMethod -Method PUT -Uri $tableUri -Headers $headers -Body $body -ContentType application/json 
    }#End function add-TableData
    
    # Generating a unique report id to link telemetry data to report data
        $DReportID=""
        $DReportID=(new-guid).guid

    # Get the internet connection IP address by querying a public API
    $internetIp = Invoke-RestMethod -Uri "https://api.ipify.org?format=json" | Select-Object -ExpandProperty ip

    # Define the API endpoint URL
    $geourl = "http://ip-api.com/json/$internetIp"
    
    # Invoke the API to determine Geolocation
    $response = Invoke-RestMethod $geourl

    $data = @{
    Region=$env:UserDomain
    DriftVersion=$DFTV
    ReportID=$DReportID
    country=$response.country
    counrtyCode=$response.countryCode
    georegion=$response.region
    regionName=$response.regionName
    city=$response.city
    zip=$response.zip
    lat=$response.lat
    lon=$response.lon
    timezone=$response.timezone
}

    add-TableData -TableName "DriftTelemetryData" -PartitionKey "DriFT" -RowKey (new-guid).guid -data $data -sasWriteToken '?sv=2017-04-17&si=Update&tn=DriftTelemetryData&sig=wrOnGPi/vI3g62CZyj4CPiJHILxopLuNLaMY/nk2idA%3D'

    #~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
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
$ExtracLoc="$env:TEMP\DriFT"
if (Test-Path $ExtracLoc -PathType Container){Remove-Item $ExtracLoc -Recurse -Force | Out-Null}
if (!(Test-Path $ExtracLoc -PathType Container)) {New-Item -ItemType Directory -Force -Path $ExtracLoc | Out-Null }

#TSR unzip files
Write-Host "Unziping TSR data files...."
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
        if (Test-Path $TSRDataInventory"\sysinfo_CIM_BIOSAttribute.xml" -PathType Leaf){
            $CIM_BIOSAttribute=[Xml] (Get-Content $TSRDataInventory"\sysinfo_CIM_BIOSAttribute.xml")
            $CIM_BIOSAttribute_Instances=$CIM_BIOSAttribute.CIM.MESSAGE.SIMPLEREQ."VALUE.NAMEDINSTANCE".INSTANCE
            
        }Else{
            $CIM_BIOSAttribute_Instances="MISSING"
        }
        if (Test-Path $TSRDataInventory"\sysinfo_DCIM_View.xml" -PathType Leaf){
            $DCIM_View=[Xml] (Get-Content $TSRDataInventory"\sysinfo_DCIM_View.xml")
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
        }
        if (Test-Path $TSRDataInventory"\sysinfo_DCIM_SoftwareIdentity.xml" -PathType Leaf){
            $DCIM_SoftwareIdentity=[Xml] (Get-Content $TSRDataInventory"\sysinfo_DCIM_SoftwareIdentity.xml")
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
            Write-host "    WARNING: SoftwareIdentity.xml missing. Please upgrade to the latest iDRAC version to see the rest of the hardware." -foregroundcolor Yellow
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
            $IDModule = $DCIM_SoftwareIdentity_Properties | Where-Object{$_.ElementName -imatch 'Identity Module'} |select ElementName
            $IsAZHub=""
            IF($AZHUBElementNames.keys -imatch (($IDModule.ElementName -split '\(')[1] -split '\)')[0]){
                $IsAZHub=$True
                Write-Host "    Found Azure Stack Hub: $IsAZHub"
            }Else{$IsAZHub=$False}
            
        #Installed hardware from Software Identity
            Write-Host "Discovering Installed Hardware..."
            $DCIM_SoftwareIdentity_NAMEDINSTANCE_INSTANCENAME_KEYBINDING_KEYVALUE_Installed = $DCIM_SoftwareIdentity_NAMEDINSTANCE|`
            Where-Object{$_.INSTANCENAME.KEYBINDING.KEYVALUE."#text" -Match "DCIM:INSTALLED"} 
        #Converting to customer property to maker easyer to manage install hardware
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
            #$InstalledHardwareUnique|FT
    }
    #Support Assist Enterprise Collection XML
    IF($SupportAssistDataType -lt 3){
    
        If($SAEDataInventory=Get-ChildItem -Path $DriFTFolders.fullname -Include "MaserInfo.xml","Inventory.xml" -File -Recurse -Force | Select-Object -last 1  | ForEach-Object{ $_.Directory } ){
            $InvPath=""
            $MasPath=""
            $SupportAssistDataType="SAEX"
            # Server Type and Service Tag
            $chasinfoxml=Get-ChildItem -Path $DriFTFolders.fullname -Include "chasinfo.xml" -File -Recurse -Force | sort-object Length | Select-Object -last 1 | ForEach-Object{ $_.Directory } 
            $SvrInfo=[Xml](Get-Content $chasinfoxml"\chasinfo.xml")
            #Firmware inventory
            $InvPath=$SAEDataInventory.FullName+"\Inventory.xml"
                IF([System.IO.File]::Exists($InvPath)){$inv=[Xml](Get-Content $InvPath)
                $XMLLIB="SVMInventory"}
            #Firmware Maser
            $MasPath=$SAEDataInventory.FullName+"\MaserInfo.xml"
                IF([System.IO.File]::Exists($MasPath)){
                $inv=[Xml](Get-Content $MasPath)
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
            $HostName=(($CIM_BIOSAttribute_Instances|`
            Where-Object{($_.CLASSNAME -match "DCIM_SystemString")}).PROPERTY|`
            Where-Object{$_.VALUE -eq "HostName"}).ParentNode.'PROPERTY.ARRAY' |`
            Where-Object{$_.Name -match "CurrentValue"}|`
                      Select-Object @{Label="CurrentValue";Expression={$_.'VALUE.ARRAY'.VALUE}}
            $HostName=$HostName.CurrentValue
            $ServiceTag=($DCIM_View_Instances| Where-object {($_.CLASSNAME -match "DCIM_SystemView")}).PROPERTY | Where-Object {$_.NAME -eq "ServiceTag"} | Select-Object @{Label="ServiceTag";Expression={$_.Value}} | Select-Object -First 1
            $ServerType=($DCIM_View_Instances| Where-Object {($_.CLASSNAME -eq "DCIM_SystemView")}).PROPERTY | Where-Object {$_.NAME -eq "MODEL"} | Select-Object @{Label="Model";Expression={$_.Value}}| Select-Object -First 1
            $SystemID=(($CIM_BIOSAttribute_Instances| Where-Object {($_.CLASSNAME -eq "DCIM_LCString")}|Where-Object{$_.PROPERTY.VALUE -eq 'SYSID'}).'PROPERTY.ARRAY' | Where-Object {$_.NAME -eq "CurrentValue"}).'VALUE.ARRAY'.VALUE
                
########### Change this to YES to force ASHCI-catalog.xml
            $S2DCatalogNeeded="No"

            $ServerType=$ServerType.Model
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
            $OSName0=$CIM_BIOSAttribute_Instances| Where-Object {($_.CLASSNAME -match "DCIM_SystemString")} | Where-Object {$_.PROPERTY.Value -Match "OSName"}
            $OSName1=$OSName0.ChildNodes | Where-Object{($_.NAME -match "CurrentValue")}
            $OSCheck=$OSName1.InnerText
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
                        [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
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
            [xml]$PrecisionCatalog=Get-Content $PrecisionCatalogExtractedPath.FullName
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
                            [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
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
                    $S2DCatalogXMLData = [Xml] (Get-Content "$env:TEMP\DriFT\ASHCI-Catalog.xml")
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
        If(($SupportAssistDataType -eq "SAEX")-or($SupportAssistDataType -eq "TSR")){
            ForEach ($Device in $InstalledHardwareUnique){
                $Found=@()
                If($S2DCatalogNeeded -eq "YES"){
                    $CatalogInfoOut=""
                    #Added for iDRAC 4.40 weird chars
                            IF($Device.deviceID.length -gt 0){
                                $Found=
                                $S2DCatalogXMLDataFiltered|`
                                Where-Object{$_.ComponentType.value -eq $Device.componentType}|`
                                Where-Object{$_.SupportedDevices.Device.componentID -match $Device.componentID}|`
                                Where-Object{(($_.SupportedDevices.Device.PCIInfo.deviceID -eq $Device.deviceID)-and`
                                ($_.SupportedDevices.Device.PCIInfo.subDeviceID -eq $Device.subdeviceID)-and`
                                ($_.SupportedDevices.Device.PCIInfo.subVendorID -eq $Device.subVendorID)-and`
                                ($_.SupportedDevices.Device.PCIInfo.vendorID -eq $Device.vendorID))}|`
                                sort-Object {[DateTime]$_.releaseDate}| Select-Object -Last 1;`
                                $CatalogInfoOut=$S2DCatalogInfo
                                }
                            Else{
                                $Found=
                                $S2DCatalogXMLDataFiltered|`
                                Where-Object{$_.ComponentType.value -eq $Device.componentType}|`
                                Where-Object{$_.SupportedDevices.Device.componentID -match $Device.componentID}|`
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
                                Where-Object{$_.ComponentType.value -eq $Device.componentType}|`
                                Where-Object{$_.SupportedDevices.Device.componentID -match $Device.componentID}|`
                                Where-Object{(($_.SupportedDevices.Device.PCIInfo.deviceID -eq $Device.deviceID)-and`
                                ($_.SupportedDevices.Device.PCIInfo.subDeviceID -eq $Device.subdeviceID)-and`
                                ($_.SupportedDevices.Device.PCIInfo.subVendorID -eq $Device.subVendorID)-and`
                                ($_.SupportedDevices.Device.PCIInfo.vendorID -eq $Device.vendorID))}|`
                                sort-Object {[DateTime]$_.releaseDate}| Select-Object -Last 1;`
                                $CatalogInfoOut=$SpecialCatalogInfo
                                }
                            Else{
                                $Found=
                                $SpecialCatalogXMLDataFiltered|`
                                Where-Object{$_.ComponentType.value -eq $Device.componentType}|`
                                Where-Object{$_.SupportedDevices.Device.componentID -match $Device.componentID}|`
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
                                Where-Object{$_.ComponentType.value -eq $Device.componentType}|`
                                Where-Object{$_.SupportedDevices.Device.componentID -match $Device.componentID}|`
                                Where-Object{(($_.SupportedDevices.Device.PCIInfo.deviceID -eq $Device.deviceID)-and`
                                ($_.SupportedDevices.Device.PCIInfo.subDeviceID -eq $Device.subdeviceID)-and`
                                ($_.SupportedDevices.Device.PCIInfo.subVendorID -eq $Device.subVendorID)-and`
                                ($_.SupportedDevices.Device.PCIInfo.vendorID -eq $Device.vendorID))}|`
                                sort-Object {[DateTime]$_.releaseDate}| Select-Object -Last 1;`
                                $CatalogInfoOut=$CatalogInfo
                                }
                            Else{
                                $Found=
                                $CatalogXMLDataFiltered|`
                                Where-Object{$_.ComponentType.value -eq $Device.componentType}|`
                                Where-Object{$_.SupportedDevices.Device.componentID -match $Device.componentID}|`
                                sort-Object {[DateTime]$_.releaseDate}| Select-Object -Last 1;`
                                $CatalogInfoOut=$CatalogInfo
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
    Write-Host "Gathering VMware supported driver versions from VMware.com..."
    
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    # Find the VMware release number from the VMware.com with the Dell keyword
        $VMwareCompatibilityGuideURL=""
        $VMwareCompatibilityGuideURL="https://www.VMware.com/resources/compatibility/search.php?"
        $VMwareCompatibilityGuidePage=""
        IF(-not($VMwareCompatibilityGuidePage)){
        $VMwareCompatibilityGuidePage=Invoke-WebRequest -Uri $VMwareCompatibilityGuideURL -UseDefaultCredentials}
        $session = New-Object Microsoft.PowerShell.Commands.WebRequestSession
        $session.Cookies.Add($VMwareCompatibilityGuidePage.BaseResponse.Cookies)
        IF(-not($VMwareCompatibilityGuidePage)){
        $VMwareCompatibilityGuidePage=Invoke-WebRequest -Uri $VMwareCompatibilityGuideURL -WebSession $session -UseDefaultCredentials}
        $VMwareReleases=""
        $VMwareReleases=((($VMwareCompatibilityGuidePage -split "var releasesShortDesc")[1] -split "var VMwareProducts")[0]).ToString() -split "[`r`n]"
        IF($VMwareReleases){
            ForEach($VMWR in $VMwareReleases){
                IF($VMWR -match $VMWOSVer){
                    $Release=""
                    $Release=((($VMWR -split ":")[0] -split '"').trim())[1]
                }
            }
        }
        IF(!($VMwareReleases)){Write-Host "    ERROR: Unable to pull supported versions from VMware.com Firmware ONLY." -ForegroundColor Red}

    # Get the link to the js/data_io.js which has all the supported devices
        # example: https://www.VMware.com/resources/compatibility/js/data_io.js?param=cbf961fd3ed631534c8b5cc61f8abd33
        $VMwareCompatibilityGuideURL=""
        $VMwareCompatibilityGuideURL="https://www.VMware.com/resources/compatibility/search.php?deviceCategory=io&details=1&keyword=dell"
        $VMwareCompatibilityGuidePage=""
        IF(-not($VMwareCompatibilityGuidePage)){
        $VMwareCompatibilityGuidePage=Invoke-WebRequest -Uri $VMwareCompatibilityGuideURL -UseDefaultCredentials}
        $data_ioLine=""
        $data_ioLine=$VMwareCompatibilityGuidePage.ToString() -split "[`r`n]" | select-string "data_io"
        $data_ioLink=($data_ioLine -split "src=")[1] -replace "></script>","" -replace '"',""

    # Get contents of the js/data_io.js which is the VMware catalog
        $data_ioURL=""
        $data_ioURL="https://www.VMware.com/resources/compatibility/"+$data_ioLink
        $data_ioContents=""
        IF(-not($data_ioContents)){
        $data_ioContents=Invoke-WebRequest -Uri $data_ioURL -UseDefaultCredentials}
   
    # Filter js/data_io.js for Dell devices
        <#
            window.col_product_id  = 0;
            window.col_partner      = 1;
            window.col_model        = 2;
            window.col_device_type  = 3;
            window.col_vid          = 4;
            window.col_did          = 5;
            window.col_svid         = 6;
            window.col_ssid         = 7;
            window.col_device_group_id = 8;
            window.col_releases     = 9;
            window.col_driver_ver   = 10;
            window.col_VMware_async = 11;
            window.col_driver_type  = 12;
            window.col_vio_solution = 13;
            window.col_solution_releases = 14;
            window.col_solution = 15;
            window.col_firmware     = 16;
            window.col_no_of_port   = 17;
            window.col_feature      = 18;
            window.col_daterange    = 19;
            window.col_DeviceChipset= 20;
            window.col_devicedriver_model= 21;
            window.col_releaseversion = 22;
        #>
        $AllDriversFromVMwareCompatibilityGuidePage=@()
        $AllDriversFromVMwareCompatibilityGuidePage=($data_ioContents.RawContent -split '];')[0] -split "[`r`n]"
        $DellDriversFromVMwareCompatibilityGuidePageByInstalledVMWOSVer=@()

        #Filter for all window.col_partner for Dell or 23
        $pattern = ', "23",'
        $DellDriversFromVMwareCompatibilityGuidePageByInstalledVMWOSVerforDellOnly = $AllDriversFromVMwareCompatibilityGuidePage | Select-String -Pattern $pattern

        ForEach($DDriver in $DellDriversFromVMwareCompatibilityGuidePageByInstalledVMWOSVerforDellOnly){
        #ForEach($DDriver in $AllDriversFromVMwareCompatibilityGuidePage){
            # 23=Dell
            If($DDriver -imatch '"23",'){
                IF($DDriver -imatch $VMWOSVer){
                    $DellDriversFromVMwareCompatibilityGuidePageByInstalledVMWOSVer+=$DDriver
                }
            }
        }
        Write-Host "    Found" $DellDriversFromVMwareCompatibilityGuidePageByInstalledVMWOSVer.count "Dell drivers for $VMWOSVer in VMWare catalog"
        
    #Match Installed Hardware to VMware Catalog
        Write-host "Comparing Installed Hardware to VMware Compatibility Guide...."
        If(($SupportAssistDataType -eq "SAEX")-or($SupportAssistDataType -eq "TSR")){
            $InstalledHardwareUniqueWithDeviceInfo=@()
            $InstalledHardwareUniqueWithDeviceInfo=$InstalledHardwareUnique | Where-Object{$_.SubDeviceID.length -gt 0}
            #Write-Host "    Devices found"
            ForEach ($Device in $InstalledHardwareUniqueWithDeviceInfo){
                # Build search query
                # Example "8086", "1523", "1028", "1F9B"
                # We need two matchs because iDRAC data flips the SVID and SSID sometimes
                    $TheMatch1=""
                    $TheMatch1="'"""+$Device.VendorID+""", """+$Device.DeviceID+""", """+$Device.SubDeviceID+""", """+$Device.SubVendorID+"""'"
                    $TheMatch1=$TheMatch1 -replace "'"
                    #Write-Host $TheMatch1
                    $TheMatch2=""
                    $TheMatch2="'"""+$Device.VendorID+""", """+$Device.DeviceID+""", """+$Device.SubVendorID+""", """+$Device.SubDeviceID+"""'"
                    $TheMatch2=$TheMatch2 -replace "'"
                    #Write-Host $TheMatch2
                # Search for SVID and SSID
                $ActualVMWDriverVersion=@()
                $ActualVMWDriverReleaseDate=@()
                $Found=@()
                $Found=$DellDriversFromVMwareCompatibilityGuidePageByInstalledVMWOSVer|Where-Object{($_ -imatch $TheMatch1 -or $_ -imatch $TheMatch2)}
                IF($Found.length -gt 0){
                    # Grab to download page link from the driver details page
                        #Write-Host "    Getting the download page link..."
                        $DriverDetailsLink=""
                        $DriverDetailsLinkPage=""
                        $ActualVMWDriverDetailsLink=""
                        $productid=""
                        $ActualVMWDriverDetailsLinkCheck=""
                        $productid=(($Found -split ", ")[0] -replace "\[","" -replace '"',"").trim()
                        $DriverDetailsLink="https://www.VMware.com/resources/compatibility/detail.php?deviceCategory=io&productid="+$productid+"&releaseid="+$Release+"&deviceCategory=io&details=1"
                        #Write-Host "    Product Link: "$DriverDetailsLink
                        $DriverDetailsLinkPage=Invoke-WebRequest -Uri $DriverDetailsLink -UseBasicParsing -UseDefaultCredentials -WebSession $session
                        #Write-Host $DriverDetailsLink
                        $ActualVMWDriverDetailsLinkCheck = ($DriverDetailsLinkPage.InputFields.value|Out-String).Split("`r`n")|Where-Object{$_ -imatch "ESXi $VMWOSVer"}
                         # Get device type from driver details 
                         #Write-Host "    Getting the device types from driver details..."
                        $DriverDetailsDeviceType=""
                        $DriverDetailsDeviceType=((((($DriverDetailsLinkPage -split "[`r`n]" | Where-Object{$_ -match 'Device Type :'}) -split 'Device Type :')[1] -split '</td>')[0] -split '<td>')[1]).trim()
                        #Inbox Link
                            #Write-Host "    Checking for Inbox link..."
                            $ActualVMWDriverDetailsLinkInbox=""
                            $ActualVMWDriverDetailsLinkCheckInbox1=""
                            $ActualVMWDriverDetailsLinkCheckInbox1=$ActualVMWDriverDetailsLinkCheck|Where-Object{$_ -imatch "inbox"}|Select-Object -First 1
                            #IF($ActualVMWDriverDetailsLinkCheckInbox1 -imatch "Download driver from"){
                                If($ActualVMWDriverDetailsLinkCheckInbox1|Where-Object{$_ -inotmatch "http"}){$ActualVMWDriverDetailsLinkInbox="None"}
                                #If($ActualVMWDriverDetailsLinkCheckInbox1|?{$_ -imatch "http"}){$ActualVMWDriverDetailsLinkInbox=((((($ActualVMWDriverDetailsLinkCheckInbox1|?{$_ -imatch "http"}) -split ",")[6]) -split "  ")[1]).trim() -replace '&amp;','&' -replace '&quot;',""}
                            # Update contained in Patch
                                If($ActualVMWDriverDetailsLinkCheckInbox1|Where-Object{$_ -imatch "http"}){$ActualVMWDriverDetailsLinkInbox=(([REGEX]::Match($ActualVMWDriverDetailsLinkCheckInbox1,'http.*.html')) -split '.html')[0]}
                            # Normal Update
                                If($ActualVMWDriverDetailsLinkCheckInbox1|Where-Object{$_ -imatch "http"}){$ActualVMWDriverDetailsLinkInbox=([REGEX]::Match($ActualVMWDriverDetailsLinkCheckInbox1,'http.*\:\/\/.*productId\=\d*')).value}
                                #(($ActualVMWDriverDetailsLinkCheckInbox1|?{$_ -imatch "http"})|select-string 'http.*')
                                #(([REGEX]::Match($ActualVMWDriverDetailsLinkCheckInbox1,'http.*.html')) -split '.html')[0]
                            #}
                            IF($ActualVMWDriverDetailsLinkInbox -imatch "productId"){
                            $ActualVMWDriverDetailsLinkInbox=$ActualVMWDriverDetailsLinkInbox -replace 'amp\;',""
                            # Grab the actual download link for the VMware driver
                                IF($ActualVMWDriverDetailsLinkInbox -imatch 'productid'){
                                    $ActualVMWDriverDetailsLinkInbox=$ActualVMWDriverDetailsLinkInbox -replace 'amp\;',""
                                    $ActualVMWDriverDownloadPage=""
                                    $ActualVMWDriverDownloadPageInbox = Invoke-WebRequest -Uri $ActualVMWDriverDetailsLinkInbox -UseBasicParsing -UseDefaultCredentials -WebSession $session
                                    $ActualVMWDriverDownloadPagenb=Invoke-WebRequest -Uri $ActualVMWDriverDetailsLinkInbox -UseDefaultCredentials -WebSession $session
                                    $ActualVMWDriverDownloadPageLink=""
                                    $ActualVMWDriverDownloadPageLink=((($ActualVMWDriverDownloadPage.RawContent -split "[`r`n]"| Where-Object{$_ -match "fileId"}|Select-Object -First 1) -split "href=")[1] -split ">Download")[0] -replace '"'
                                    If($ActualVMWDriverDetailsLinkInbox -imatch "http"){
                                        $VMwarePublicSiteUrl1=""
                                        $VMwarePublicSiteUrl2=""
                                        $VMwarePublicSiteUrl1='https://my.vmware.com/channel/public/api/v1.0/dlg/details?locale=en_US&downloadGroup='
                                        $VMwarePublicSiteUrl2=$ActualVMWDriverDetailsLinkInbox.Split('?') -replace "downloadGroup=",""|Select-Object -last 1
                                        $ActualVMWDriverDetailsLinkInboxText="$VMwarePublicSiteUrl1$VMwarePublicSiteUrl2"
                                        #https://my.vmware.com/channel/public/api/v1.0/dlg/details?locale=en_US&downloadGroup=
                                        $ActualVMWDriverDetailsLinkInboxDownloadProductPage=""
                                        $ActualVMWDriverDetailsLinkInboxDownloadProductPage=Invoke-WebRequest -Uri $ActualVMWDriverDetailsLinkInboxText -UseDefaultCredentials -WebSession $session
                                        #Write-Host "    Inbox Link" $ActualVMWDriverDetailsLinkInbox
                                        $ActualVMWDriverDownloadPageRawContent=@()
                                        $ActualVMWDriverDetailsLinkInboxDownloadProductPageData=@()
                                        $ActualVMWDriverDetailsLinkInboxDownloadProductPageData=($ActualVMWDriverDetailsLinkInboxDownloadProductPage.RawContent -split "['`r`n']" | Select-Object -last 1) -replace '\{' -replace '\}' -replace '^\"\D*\"\:\[' -replace '\]' -split ',' 
                                        # Driver version
                                            #$ActualVMWDriverDownloadPageRawContent | ?{$_ -imatch "Version"}
                                            $ActualVMWDriverVersion+="Inbox: "+((($ActualVMWDriverDetailsLinkInboxDownloadProductPageData|Where-Object{$_ -match '\"version\"'})|Select-Object -First 1) -split '\:')[1] -replace '\"'
                                        # Driver Release Date
                                            $ActualVMWDriverReleaseDate+="Inbox: "+((($ActualVMWDriverDetailsLinkInboxDownloadProductPageData|Where-Object{$_ -match '\"releaseDate\"'})|Select-Object -First 1) -split '\:')[1] -replace '\"'
                                    }
                                }
                            }
                        #Native Link
                           # Write-Host "    Checking for Native link..."
                            $ActualVMWDriverDetailsLinkNative=""
                            $ActualVMWDriverDetailsLinkCheckNative1=""
                            $ActualVMWDriverDetailsLinkCheckNative1=$ActualVMWDriverDetailsLinkCheck|Where-Object{$_ -imatch "native"}|Select-Object -First 1
                            If($ActualVMWDriverDetailsLinkCheckNative1|Where-Object{$_ -inotmatch "http"}|Select-Object -First 1){$ActualVMWDriverDetailsLinkNative="None"}
                            #If($ActualVMWDriverDetailsLinkCheckNative1|?{$_ -imatch "http"}){$ActualVMWDriverDetailsLinkNative=((((($ActualVMWDriverDetailsLinkCheckNative1|?{$_ -imatch "http"}) -split ",")[6]) -split "  ")[1]).trim() -replace '&amp;','&' -replace '&quot;',""}
                            # Update contained in Patch
                            If($ActualVMWDriverDetailsLinkCheckNative1|Where-Object{$_ -imatch "http"}){$ActualVMWDriverDetailsLinkNative=(([REGEX]::Match( $ActualVMWDriverDetailsLinkCheckNative1,'http.*.html')) -split '.html')[0]}
                            # Normal Update
                            If($ActualVMWDriverDetailsLinkCheckNative1|Where-Object{$_ -imatch "http"}){$ActualVMWDriverDetailsLinkNative=([REGEX]::Match( $ActualVMWDriverDetailsLinkCheckNative1,'http.*\:\/\/.*productId\=\d*')).value}
                            #$ActualVMWDriverDetailsLinkNative
                            IF($ActualVMWDriverDetailsLinkNative -imatch "productid"){
                                #$ActualVMWDriverDownloadLinkVmLinux=GetVmwDriverDownloadLink $ActualVMWDriverDetailsLinkVmkLinux
                            # Grab the actual download link for the VMware driver
                                IF($ActualVMWDriverDetailsLinkNative -imatch 'productid'){
                                    $ActualVMWDriverDetailsLinkNative=$ActualVMWDriverDetailsLinkNative -replace 'amp\;',""
                                    $ActualVMWDriverDownloadPage=""
                                    $ActualVMWDriverDownloadPageNative = Invoke-WebRequest -Uri $ActualVMWDriverDetailsLinkNative -UseBasicParsing -UseDefaultCredentials -WebSession $session
                                    $ActualVMWDriverDownloadPagenb=Invoke-WebRequest -Uri $ActualVMWDriverDetailsLinkNative -UseDefaultCredentials -WebSession $session
                                    $ActualVMWDriverDownloadPageLink=""
                                    $ActualVMWDriverDownloadPageLink=((($ActualVMWDriverDownloadPage.RawContent -split "[`r`n]"| Where-Object{$_ -match "fileId"}|Select-Object -First 1) -split "href=")[1] -split ">Download")[0] -replace '"'
                                    If($ActualVMWDriverDetailsLinkNative -imatch "http"){
                                        $VMwarePublicSiteUrl1=""
                                        $VMwarePublicSiteUrl2=""
                                        $VMwarePublicSiteUrl1='https://my.vmware.com/channel/public/api/v1.0/dlg/details?locale=en_US&downloadGroup='
                                        $VMwarePublicSiteUrl2=$ActualVMWDriverDetailsLinkNative.Split('?') -replace "downloadGroup=",""|Select-Object -last 1
                                        $ActualVMWDriverDetailsLinkNativeText="$VMwarePublicSiteUrl1$VMwarePublicSiteUrl2"
                                        #https://my.vmware.com/channel/public/api/v1.0/dlg/details?locale=en_US&downloadGroup=
                                        $ActualVMWDriverDetailsLinkNativeDownloadProductPage=""
                                        $ActualVMWDriverDetailsLinkNativeDownloadProductPage=Invoke-WebRequest -Uri $ActualVMWDriverDetailsLinkNativeText -UseDefaultCredentials -WebSession $session
                                        #Write-Host "    Native Link" $ActualVMWDriverDetailsLinkNative
                                        $ActualVMWDriverDownloadPageRawContent=@()
                                        $ActualVMWDriverDetailsLinkNativeDownloadProductPageData=@()
                                        $ActualVMWDriverDetailsLinkNativeDownloadProductPageData=($ActualVMWDriverDetailsLinkNativeDownloadProductPage.RawContent -split "['`r`n']" | Select-Object -last 1) -replace '\{' -replace '\}' -replace '^\"\D*\"\:\[' -replace '\]' -split ',' 
                                        # Driver version
                                            #$ActualVMWDriverDownloadPageRawContent | ?{$_ -imatch "Version"}
                                            $ActualVMWDriverVersion+="Native: "+((($ActualVMWDriverDetailsLinkNativeDownloadProductPageData|Where-Object{$_ -match '\"version\"'})|Select-Object -First 1) -split '\:')[1] -replace '\"'
                                        # Driver Release Date
                                            $ActualVMWDriverReleaseDate+="Native: "+((($ActualVMWDriverDetailsLinkNativeDownloadProductPageData|Where-Object{$_ -match '\"releaseDate\"'})|Select-Object -First 1) -split '\:')[1] -replace '\"'
                                    }
                                }
                            }
                        #Vmk Link
                            #Write-Host "    Checking for Linux link..."
                            $ActualVMWDriverDetailsLinkVmkLinux=""
                            $ActualVMWDriverDetailsLinkCheckVmkLinux1=""
                            $ActualVMWDriverDetailsLinkCheckVmkLinux1=$ActualVMWDriverDetailsLinkCheck|Where-Object{$_ -imatch "VmkLinux"}|Select-Object -First 1
                            If($ActualVMWDriverDetailsLinkCheckVmkLinux1|Where-Object{$_ -inotmatch "http"}|Select-Object -First 1){$ActualVMWDriverDetailsLinkVmkLinux="None"}
                            #If($ActualVMWDriverDetailsLinkCheckVmkLinux1|?{$_ -imatch "http"}){
                            #    $ActualVMWDriverDetailsLinkVmkLinux=(((($ActualVMWDriverDetailsLinkCheckVmkLinux1|?{$_ -imatch "http"}) -split ",")[6]) -split "  ")[1] -replace '&amp;','&' -replace '&quot;',""}
                            # Update contained in Patch
                            If($ActualVMWDriverDetailsLinkCheckVmkLinux1|Where-Object{$_ -imatch "http"}){$ActualVMWDriverDetailsLinkVmkLinux=(([REGEX]::Match($ActualVMWDriverDetailsLinkCheckVmkLinux1,'http.*.html')) -split '.html')[0]}
                            # Normal Update
                            If($ActualVMWDriverDetailsLinkCheckVmkLinux1|Where-Object{$_ -imatch "http"}){$ActualVMWDriverDetailsLinkVmkLinux=([REGEX]::Match($ActualVMWDriverDetailsLinkCheckVmkLinux1,'http.*\:\/\/.*productId\=\d*')).value}
                            IF($ActualVMWDriverDetailsLinkVmkLinux -imatch "productid"){
                                $ActualVMWDriverDetailsLinkVmkLinux=$ActualVMWDriverDetailsLinkVmkLinux -replace 'amp\;',""
                                #$ActualVMWDriverDownloadLinkVmLinux=GetVmwDriverDownloadLink $ActualVMWDriverDetailsLinkVmkLinux
                            # Grab the actual download link for the VMware driver
                                    IF($ActualVMWDriverDetailsLinkVmkLinux -imatch 'productid'){
                                    $ActualVMWDriverDetailsLinkVmkLinux=$ActualVMWDriverDetailsLinkVmkLinux -replace 'amp\;',""
                                    $ActualVMWDriverDownloadPage=""
                                    $ActualVMWDriverDownloadPageVmkLinux = Invoke-WebRequest -Uri $ActualVMWDriverDetailsLinkVmkLinux -UseBasicParsing -UseDefaultCredentials -WebSession $session
                                    $ActualVMWDriverDownloadPagenb=Invoke-WebRequest -Uri $ActualVMWDriverDetailsLinkVmkLinux -UseDefaultCredentials -WebSession $session
                                    $ActualVMWDriverDownloadPageLink=""
                                    $ActualVMWDriverDownloadPageLink=((($ActualVMWDriverDownloadPage.RawContent -split "[`r`n]"| Where-Object{$_ -match "fileId"}|Select-Object -First 1) -split "href=")[1] -split ">Download")[0] -replace '"'
                                    If($ActualVMWDriverDetailsLinkVmkLinux -imatch "http"){
                                        $VMwarePublicSiteUrl1=""
                                        $VMwarePublicSiteUrl2=""
                                        $VMwarePublicSiteUrl1='https://my.vmware.com/channel/public/api/v1.0/dlg/details?locale=en_US&downloadGroup='
                                        $VMwarePublicSiteUrl2=$ActualVMWDriverDetailsLinkVmkLinux.Split('?') -replace "downloadGroup=",""|Select-Object -last 1
                                        $ActualVMWDriverDetailsLinkVmkLinuxText="$VMwarePublicSiteUrl1$VMwarePublicSiteUrl2"
                                        #https://my.vmware.com/channel/public/api/v1.0/dlg/details?locale=en_US&downloadGroup=
                                        $ActualVMWDriverDetailsLinkVmkLinuxDownloadProductPage=""
                                        $ActualVMWDriverDetailsLinkVmkLinuxDownloadProductPage=Invoke-WebRequest -Uri $ActualVMWDriverDetailsLinkVmkLinuxText -UseDefaultCredentials -WebSession $session
                                        #Write-Host "    VmkLinux Link" $ActualVMWDriverDetailsLinkVmkLinux
                                        $ActualVMWDriverDownloadPageRawContent=@()
                                        $ActualVMWDriverDetailsLinkVmkLinuxDownloadProductPageData=@()
                                        $ActualVMWDriverDetailsLinkVmkLinuxDownloadProductPageData=($ActualVMWDriverDetailsLinkVmkLinuxDownloadProductPage.RawContent -split "['`r`n']" | Select-Object -last 1) -replace '\{' -replace '\}' -replace '^\"\D*\"\:\[' -replace '\]' -split ',' 
                                        # Driver version
                                            $ActualVMWDriverVersion+="VmkLinux: "+((($ActualVMWDriverDetailsLinkVmkLinuxDownloadProductPageData|Where-Object{$_ -match '\"version\"'})|Select-Object -First 1) -split '\:')[1] -replace '\"'
                                        # Driver Release Date
                                            $ActualVMWDriverReleaseDate+="VmkLinux: "+((($ActualVMWDriverDetailsLinkVmkLinuxDownloadProductPageData|Where-Object{$_ -match '\"releaseDate\"'})|Select-Object -First 1) -split '\:')[1] -replace '\"'
                                    }
                                }
                            }
                            #Pause
                    # Add devices to report
                        $allArray+=$Found|`
                        Select-object @{Label="ServiceTag";Expression={"$ServiceTag"}},`
                        @{Label="PowerEdge";Expression={"$ServerType"}},`
                        @{Label="OS";Expression={"$InstalledOS"+" "+"$OSVersion"}},`
                        @{Label="Type";Expression={"DRVR"}},`
                        @{Label="Category";Expression={$DriverDetailsDeviceType}},`
                        @{Label="Name";Expression={(($_ -split ",")[2] -replace '"',"").trim()}},`
                        @{Label="InstalledVersion";Expression={"NA"}},`
                        @{Label="AvailableVersion";Expression={IF($ActualVMWDriverVersion){$ActualVMWDriverVersion}Else{"No Data Found"}}},`
                        @{Label="CatalogInfo";Expression={"VMware.com"}},`
                        @{Label="Criticality";Expression={"No Data Found"}},`
                        @{Label="ReleaseDate";Expression={IF($ActualVMWDriverReleaseDate){$ActualVMWDriverReleaseDate}Else{"No Data Found"}}},`
                        @{Label="Details";Expression={
                            $Details=@()
                            IF($DriverDetailsLink -imatch "http"){$Details+="<a href='$($DriverDetailsLink)' target='_blank'>$("Driver List")</a><br>"}
                            $Details
                        }},@{Label="URL";Expression={
                            $URLDetails=@()
                            If($ActualVMWDriverDetailsLinkInbox -imatch "http"){
                                IF($ActualVMWDriverDetailsLinkInboxText.Length -gt 0){$URLDetails+="InBox: <a href='$($ActualVMWDriverDetailsLinkInbox)' target='_blank'>$($ActualVMWDriverDetailsLinkInbox)</a><br>"}
                                IF($ActualVMWDriverDetailsLinkInbox -inotmatch "productId"){$URLDetails+="InBox: <a href='$($ActualVMWDriverDetailsLinkInbox)' target='_blank'>$("Device Update Contained in OS Patch")</a><br>"}
                            }
                            If($ActualVMWDriverDetailsLinkNative -imatch "http"){
                                IF($ActualVMWDriverDetailsLinkNativeText.Length -gt 0){$URLDetails+="Native: <a href='$($ActualVMWDriverDetailsLinkNative)' target='_blank'>$($ActualVMWDriverDetailsLinkNative)</a><br>"}
                                IF($ActualVMWDriverDetailsLinkNative -inotmatch "productId"){$URLDetails+="Native: <a href='$($ActualVMWDriverDetailsLinkNative)' target='_blank'>$("Device Update Contained in OS Patch")</a><br>"}
                            }
                            If($ActualVMWDriverDetailsLinkVmkLinux -imatch "http"){
                                IF($ActualVMWDriverDetailsLinkVmkLinuxText.Length -gt 0){$URLDetails+="VmkLinux: <a href='$($ActualVMWDriverDetailsLinkVmkLinux)' target='_blank'>$("$ActualVMWDriverDetailsLinkVmkLinux")</a><br>"}
                                IF($ActualVMWDriverDetailsLinkVmkLinux -inotmatch "productId"){$URLDetails+="VmkLinux: <a href='$($ActualVMWDriverDetailsLinkVmkLinux)' target='_blank'>$("Device Update Contained in OS Patch")</a><br>"}
                            }
                            IF(-not $URLDetails){$URLDetails+="<a href='$($DriverDetailsLink)' target='_blank'>$("No Download Links Available, See Driver List")</a><br>"}
                            $URLDetails -replace 'amp\;',""
                        }}|sort-object Type,Category,Name
                        #Pause
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
                $CurrentMBSelFullNames=Get-ChildItem -Path $E -Filter CurrentMBSel.txt -File -Recurse -Force | ForEach-Object{ $_.fullname }
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



#$KBNumber = "123456"
#$Product = "Windows 10"
#$downloadLink = [GetKBDLLink]::GetDownloadLink($KBNumber, $Product)
#Write-Host "Download link: $downloadLink"


        $KBLatest=''
$KBList=''
$KBItemsToShow = 6

        #Lastest hotfix for Windows Server from the respective KB pages
            $OSType=$OSCheck
            If($OSCheck -imatch '2008r2'-or $OSCheck -imatch '2008 r2'){\
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
$KBDLUriSource = [GetKBDLLink]::GetDownloadLink($KBLatest.KBNumber,$OSType)

        $dstop=Get-Date
        #Write-Host "Total time taken is $(($dstop-$dstart).totalmilliseconds)"
#endregion Recommended updates and hotfixes for Windows Server
            
                $WSLCU = New-Object -TypeName PSObject
                #Add-Member -InputObject $WSLCU -MemberType NoteProperty -Name Build -Value $Build
                Add-Member -InputObject $WSLCU -MemberType NoteProperty -Name KBNumber -Value $KBLatest.KBNumber
                Add-Member -InputObject $WSLCU -MemberType NoteProperty -Name LastUpdated -Value  ([DateTime]$KBLatest.Date).ToString("yyyy-MM-dd")
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

# Upload report data to Azure
ForEach($row in $allArrayout){
    $RowData=$row|Select-Object @{Label="ReportID";Expression={"$DReportID"}},PowerEdge,OS,Type,Category,Name,InstalledVersion,AvailableVersion,CatalogInfo,Criticality,ReleaseDate,URL,Details
    
    add-TableData -TableName "DriFTReportData" -PartitionKey "DriFT" -RowKey (new-guid).guid -data $RowData -sasWriteToken '?sv=2017-04-17&si=Update&tn=DriftReportData&sig=Td%2BlJIrST3qQQCNAglJ3OmLVqWmpbmSug4cTEBreVSE%3D'
    }


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
$HTAOut=$SourcePath+$DriFTVer+"_"+$DTString+$ServiceTagList+".html"
$HTAOut=$HTAOut.Replace("*","")
IF ($HTAOut.Length -gt 248){
    $HTAOut=$SourcePath+$DriFTVer+"_"+$DTString+".html"
}
Write-Host "Report Output location: "$HTAOut
if (Test-Path "$HTAOut") {Remove-Item $HTAOut}
$OutTitle=@()
#$OutTitle+='<b><font size="4">Results</font></b><br>'
$OutTitle+="DriFT v"+$DFTV
$OutTitle+='<br>Date/Time: '+$DateTime 
$OutTitle+='<br>*A <a style="background-color:Red;color:White;">red</a> InstalledVersion indicates the InstalledVersion is less than the AvailableVersion.'
$OutTitle+='<br>**A <a style="background-color:Yellow;">Not Available</a> InstalledVersion indicates the InstalledVersion was NOT contained in the Support Assist Collection so the latest version is shown.'
$Header = @"
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
    /* Position the tooltip */
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
$Footer=@()
$FooterNote=@()
IF(($allArrayout.PowerEdge | sort-object -Unique) -imatch 'Precision'){
    $PrecisionNodes=$allArrayout.PowerEdge | sort-object -Unique
    $DownloadsDellComUrl=@()
    ForEach($PN in $PrecisionNodes){
        IF($PN.Length -gt 2){
            $PrecisionNumbers=$PN -replace 'Precision' -replace 'Rack' -replace ' ' -replace "r" -replace "7910","precision-r7910-workstation" -replace "7920","precision-7920r-workstation"
            $DownloadsDellComUrl+="<a href='https://www.dell.com/support/home/en-us/product-support/product/$PrecisionNumbers/drivers' target='_blank'>https://www.dell.com/support/home/en-us/product-support/product/$PrecisionNumbers/drivers</a>"
        }
    }
}Else{
    $DownloadsDellComUrl="<a href='http://dl.dell.com/published/pages/poweredge-$ServerType.html' target='_blank'>http://dl.dell.com/published/pages/poweredge-$ServerType.html</a>"
}
If($NoneSupportedDevices){
$NoneSupportedDevices=$NoneSupportedDevices -replace ",","<br>"
$FooterNote='<font color="red">The following device(s) are NOT listed as supported in the CATALOG.XML for this server type: <br>'
$Footer+=$FooterNote
$Footer+=$NoneSupportedDevices+"</font><br>"
$Footer+="More Driver and FW may be found here: <br>"
$Footer+=$DownloadsDellComUrl
If($S2DCatalogNeeded -eq "YES"){$Footer+="***Storage Spaces Direct Ready Node(s) Found. Special S2D catalog used to determine certified drivers and firmware compliance.<br>"}
$Footer+=$CatVerInfo
}Else{
$Footer="NOTES: <br>"
If($S2DCatalogNeeded -eq "YES"){$Footer+="***Storage Spaces Direct Ready Node(s) Found. Special S2D catalog used to determine certified drivers and firmware compliance.<br>"}
$Footer+="More Driver and FW information can be found here: <br>"
$Footer+=$DownloadsDellComUrl
#$Footer+=$CatVerInfo
$Footer+="<br><a href='https://solutions.one.dell.com/sites/NAEnterprise/SST/Communities/DRiFT/_layouts/15/start.aspx#/Lists/DriFT%20Feedback/Default.aspx' target='_blank'>Got Feedback?</a>"
}
    $AddURL=@()
    $AddURL=$allArrayout | sort-object ServiceTag,PowerEdge,Type,Category | Select-Object ServiceTag,PowerEdge,OS,Type,Category,Name,InstalledVersion,AvailableVersion,CatalogInfo,`
        @{Label="Criticality";Expression={
                    IF($_.Criticality -match "-"){
                        $CriticalityNote = $_.Criticality
                        $SplitPos=$CriticalityNote.Indexof("-")
                        $CriticalityNote0=$CriticalityNote.Substring(0,$SplitPos)
                        $CriticalityNote1=$CriticalityNote.Substring($SplitPos+1)
                        "<div class='tooltip'>$CriticalityNote0<span class='tooltiptext'>$CriticalityNote1</span>"
                        }Else{$_.Criticality}
                    }},`
    ReleaseDate,`
    @{Label="Documentation";Expression={IF($_.Details.length -gt 0){
        If($_.Details -notmatch '<br>'){"<a href='$($_.Details)' target='_blank'>$("Link")</a>"}Else{$_.Details}}}},`
    @{Label="Download Link";Expression={
        IF(($_.URL.length -gt 0) -and ($_.URL -inotmatch "href")){"<a href='$($_.URL)'>$($_.URL)</a>"}
        Else{$_.URL}
        }}
IF(($allArrayout.ServiceTag | sort -Unique).count -gt 1){
    # New multi node report view for comparing nodes
    $NewReportView = New-Object System.Data.DataTable "NodeCompare"
    $NewReportView.Columns.add((New-Object System.Data.DataColumn("Type")))
    $NewReportView.Columns.add((New-Object System.Data.DataColumn("Name")))
    $NewReportView.Columns.add((New-Object System.Data.DataColumn("AvailableVersion")))
    $NewReportView.Columns.add((New-Object System.Data.DataColumn("CatalogInfo")))
    $NewReportView.Columns.add((New-Object System.Data.DataColumn("Criticality")))
    $NewReportView.Columns.add((New-Object System.Data.DataColumn("ReleaseDate")))
    ForEach ($a in ($allArrayout.ServiceTag | Sort-Object -Unique) -Replace '\*'){
        $NewReportView.Columns.Add((New-Object System.Data.DataColumn([string]$a)))}
        $a=$null
        ForEach($b in ($allArrayout | Sort-Object name )){
            IF($b.Name.length -gt 10 -and $b.Name.length -notmatch 'System.__ComObject'){
                if ($b.name -ne $a) {
                    $a=$b.name
                    if ($null -ne $a) {
                        IF($row.constructor -inotmatch 'System.__ComObject'){
                        $NewReportView.rows.add($row)}}
                    $row=$NewReportView.NewRow()
                    $row["Type"]=$b.Type
                    $row["name"]="<a href='$($b.Details)' target='_blank'>$($b.Name)</a>"
                    $row["AvailableVersion"]="<a href='$($b.URL)'>$($b.AvailableVersion)</a>"
                    $row["CatalogInfo"]=$b.CatalogInfo
                    $BCriticality = IF($b.Criticality -match "-"){
                        $CriticalityNote = $b.Criticality
                        $SplitPos=$CriticalityNote.Indexof("-")
                        $CriticalityNote0=$CriticalityNote.Substring(0,$SplitPos)
                        $CriticalityNote1=$CriticalityNote.Substring($SplitPos+1)
                        "<div class='tooltip'>$CriticalityNote0<span class='tooltiptext'>$CriticalityNote1</span>"
                        }Else{$b.Criticality}
                    $row["Criticality"]=$BCriticality
                    $row["ReleaseDate"]=$b.ReleaseDate
                }
                $row["$($b.ServiceTag -Replace '\*')"] = $b.installedversion
            }#IF($b.Name.length -gt 10 
        }#ForEach($b
    }
    #$NewReportView|Sort-Object Type,Name| Format-Table -Property @{E="Name";width = 50},???????,AvailableVersion,Criticality,ReleaseDate,CatalogInfo
    IF(($allArrayout.ServiceTag | sort-object -Unique).count -lt 2){
        # Single Node report
        $ResultConvert=$AddURL | ConvertTo-Html -Head $Header -PreContent $OutTitle -PostContent $Footer
    }Else{
        # Multi Node report
        $ResultConvert=$NewReportView | Where-Object{$_.type -NotMatch '@{ServiceTag='} | Sort-Object Type,Name | Select-object -Property Type,Name,???????,AvailableVersion,Criticality,ReleaseDate,CatalogInfo -Exclude RowError, RowState, Table, ItemArray, HasErrors | ConvertTo-Html -Head $Header -PreContent $OutTitle -PostContent $Footer
    }

    $ResultConvertOut=$ResultConvert -replace '&gt;','>' -replace '&lt;','<' -replace '&#39;',"'"`
    -replace '<td>INSTALLED</td>','<td style="background-color: #00ff00">INSTALLED</td>'`
    -replace '<td>MISSING</td>','<td style="color: #ffffff; background-color: #ff0000">MISSING</td>'`
    -replace '<title">hTML TABLE</title>' ,'<title"></title>'`
    -replace '<tr><th>STATUS</th><th>KB Number</th><th>LINK</th></tr>','<tr style="color: #ffffff; background-color: #0000ff"><th>STATUS</th><th>KB Number</th><th>LINK</th></tr>'`
    -replace 'td">hy','td>hy'`
    -replace [Regex]::Escape("<td>***"),'<td style="color: #ffffff; background-color: #ff0000">'`
    -replace '<td>NA</td>','<td style="background-color: #ffff00">Not Available</td>'`
    -replace '<td>Not Applicable</td>','<td style="background-color: #ffff00">Not Available</td>'
    $HTAOut=$HTAOut.Replace("*","")
    Out-File $HTAOut -InputObject $ResultConvertOut
    IF($SkipDriversandFirmware -eq "NO"){
        If($OutputType -ne "NO"){
            Invoke-Item($HTAOut)
            IF($IsAZHub -eq $True){
                $DriFTCSVOut=$NewReportView|Where-Object{$_.type -iNotMatch 'System.__ComObject'} | Sort-Object Type,Name | Select-object -Property * -Exclude RowError, RowState, Table, ItemArray, HasErrors | ConvertTo-Csv
                Out-File $SourcePath+$DriFTVer+"_"+$DTString+".csv" -InputObject $DriFTCSVOut
            }
            # Export BIOSandNICCFG to XML
            If($BIOSandNICCFG.length -gt 0){
                $BIOSandNICCFGOutPutPath=""
                $BIOSandNICCFGOutPutPath=$SourcePath+"\"+$FileNameGuid+"_BIOSandNICCFG.xml"
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
IF(Test-Path $ExtracLoc){Remove-Item $ExtracLoc -Recurse}
}