# Telemetry Tool Manager v4.1 | Windows PowerShell 5.1+ | Windows Forms
# Requires Az.Accounts and Az.Storage for Azure operations.
# Uses account-key authentication for table entities; no Table Data RBAC role needed.
# Requires permission to list storage account keys. Never logs or exports keys.
# Designed for the registry-driven Azure Functions and dashboard.
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()
$script:Endpoint = 'https://gsetools-bufhdqefb8e6ecc6.centralus-01.azurewebsites.net/api/PostTelemetryData'
$script:StorageName='gsetools'
$script:ConfigTable='TelemetryToolsConfig'
$script:TenantId='f8d8cbe0-40fa-48f6-a286-ae49d95f0fb7'
$script:SubscriptionId='12c8e553-7d64-476e-8808-8da263df3e2c'
$script:ResourceGroup='GSETools'
$script:Account=$null
$script:StorageContext=$null
$script:StorageKey=$null
$script:Saved=$false
$script:InstalledByManager = New-Object System.Collections.ArrayList
$script:ImportedByManager = New-Object System.Collections.ArrayList
function Initialize-AzureModules {
    # Install only missing modules in the current user's profile; never require elevation.
    foreach ($name in @('Az.Accounts','Az.Storage')) {
        $existing = Get-Module -ListAvailable -Name $name | Sort-Object Version -Descending | Select-Object -First 1
        if (-not $existing) {
            Log "Installing $name for CurrentUser (no administrator rights required)..."
            if (-not (Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction SilentlyContinue)) {
                Install-PackageProvider -Name NuGet -Scope CurrentUser -Force -ErrorAction Stop | Out-Null
            }
            Install-Module -Name $name -Scope CurrentUser -Repository PSGallery -Force -AllowClobber -ErrorAction Stop
            [void]$script:InstalledByManager.Add($name)
            Log "Installed $name for CurrentUser."
        }
        $wasLoaded = [bool](Get-Module -Name $name)
        Import-Module -Name $name -ErrorAction Stop
        if (-not $wasLoaded) { [void]$script:ImportedByManager.Add($name) }
    }
}
function Cleanup-AzureModules {
    # Remove only modules imported by this manager, leaving previously loaded modules alone.
    foreach ($name in @($script:ImportedByManager | Select-Object -Unique)) {
        try { Remove-Module -Name $name -Force -ErrorAction Stop; Log "Unloaded $name." }
        catch { Log "WARN: Unable to unload ${name}: $($_.Exception.Message)" }
    }
    # Only uninstall modules installed during this run. Never remove pre-existing modules.
    foreach ($name in @($script:InstalledByManager | Select-Object -Unique)) {
        try {
            Uninstall-Module -Name $name -AllVersions -Force -ErrorAction Stop
            Log "Uninstalled $name (installed by this manager)."
        }
        catch { Log "WARN: Unable to uninstall ${name}: $($_.Exception.Message)" }
    }
    # PowerShellGet may leave dependency modules behind; do not delete shared dependencies.
}

$script:Bg=[Drawing.Color]::FromArgb(13,20,30)
$script:Panel=[Drawing.Color]::FromArgb(23,34,49)
$script:InputBg=[Drawing.Color]::FromArgb(33,47,64)
$script:Ink=[Drawing.Color]::FromArgb(233,241,248)
$script:Muted=[Drawing.Color]::FromArgb(161,182,201)
$script:Blue=[Drawing.Color]::FromArgb(0,118,206)
$script:Green=[Drawing.Color]::FromArgb(48,196,142)
function Label([string]$text,[int]$x,[int]$y,[int]$w,[int]$h=25,[int]$size=10,[bool]$bold=$false){
 $c=New-Object Windows.Forms.Label; $c.Text=$text; $c.SetBounds($x,$y,$w,$h); $c.ForeColor=$script:Ink
 $style=if($bold){[Drawing.FontStyle]::Bold}else{[Drawing.FontStyle]::Regular}; $c.Font=New-Object Drawing.Font('Segoe UI',$size,$style); return $c
}
function Input([int]$x,[int]$y,[int]$w,[string]$value=''){
 $c=New-Object Windows.Forms.TextBox; $c.SetBounds($x,$y,$w,32); $c.Text=$value; $c.Font=New-Object Drawing.Font('Segoe UI',11)
 $c.BackColor=$script:InputBg; $c.ForeColor=$script:Ink; $c.BorderStyle='FixedSingle'; return $c
}
function Button([string]$text,[int]$x,[int]$y,[int]$w,[bool]$secondary=$false){
 $c=New-Object Windows.Forms.Button; $c.Text=$text; $c.SetBounds($x,$y,$w,38); $c.FlatStyle='Flat'; $c.FlatAppearance.BorderSize=0
 $c.BackColor=if($secondary){$script:InputBg}else{$script:Blue}; $c.ForeColor=$script:Ink; $c.Cursor='Hand';
 $c.Font=New-Object Drawing.Font('Segoe UI',10,[Drawing.FontStyle]::Bold); return $c
}
function Log([string]$s){ $txtLog.AppendText(('{0:HH:mm:ss}  {1}' -f (Get-Date),$s)+[Environment]::NewLine); $txtLog.SelectionStart=$txtLog.TextLength; $txtLog.ScrollToCaret() }
function Inputs {
 $name=$txtName.Text.Trim(); if($name -cnotmatch '^[A-Za-z][A-Za-z0-9]{2,40}$'){throw 'Tool name must be 3-41 letters/digits and start with a letter.'}
 $minutes=0; if(-not [int]::TryParse($txtMinutes.Text.Trim(),[ref]$minutes) -or $minutes -lt 0 -or $minutes -gt 100000){throw 'Minutes saved must be a whole number between 0 and 100000.'}
 $version=$txtVersion.Text.Trim()
 # Accept simple variables, scoped variables, property paths and hashtable indexing.
 # Examples: $Ver, $script:Version, $Config.Version, $Config['Version'], $Config["Version"], $Config['Build'].Version
 $identifier='[A-Za-z_][A-Za-z0-9_]*'
 $root='\$(?:(?:script|global|local|private):)?'+$identifier
 $segment='(?:\.'+$identifier+'|\[(?:''[^''\r\n]+''|"[^"\r\n]+")\])'
 $pattern='^'+$root+'(?:'+$segment+')*$'
 if($version -cnotmatch $pattern){throw 'Version must be a variable or property/hashtable reference, e.g. $Ver, $script:Version, $Config.Version, or $Config[''Version''].'}
 return @{Name=$name;Table="${name}TelemetryData";Minutes=$minutes;Version=$version;Enabled=[bool]$chkEnabled.Checked;IncludeInTotals=[bool]$chkTotals.Checked}
}
function Invoke-TableApi([string]$method,[string]$path,$body=$null,[hashtable]$extra=@{}) {
 if(-not $script:StorageKey){throw 'Storage authentication is not initialized. Connect to Microsoft GSE first.'}
 $account=$script:StorageName
 $url="https://$account.table.core.windows.net/$path"
 $date=[DateTime]::UtcNow.ToString('R',[Globalization.CultureInfo]::InvariantCulture)
 # Azure Table Storage SharedKeyLite: sign the date and canonicalized resource.
 $canonical="/$account/$path"
 $stringToSign="$date`n$canonical"
 $hmac=New-Object System.Security.Cryptography.HMACSHA256
 try {
  $hmac.Key=[Convert]::FromBase64String($script:StorageKey)
  $signature=[Convert]::ToBase64String($hmac.ComputeHash([Text.Encoding]::UTF8.GetBytes($stringToSign)))
 } finally { $hmac.Dispose() }
 $headers=@{
  'Authorization'="SharedKeyLite ${account}:$signature"
  'x-ms-date'=$date
  'x-ms-version'='2019-02-02'
  'Accept'='application/json;odata=nometadata'
  'DataServiceVersion'='3.0;NetFx'
  'MaxDataServiceVersion'='3.0;NetFx'
 }
 foreach($key in $extra.Keys){$headers[$key]=$extra[$key]}
 $params=@{Uri=$url;Method=$method;Headers=$headers;ErrorAction='Stop';TimeoutSec=25}
 if($null -ne $body){$params.ContentType='application/json';$params.Body=($body | ConvertTo-Json -Depth 12 -Compress)}
 try { return Invoke-RestMethod @params }
 catch {
  $status='unknown';$details=''
  try{$status=[int]$_.Exception.Response.StatusCode}catch{}
  try { if($_.ErrorDetails.Message){$details=[string]$_.ErrorDetails.Message} }catch{}
  if($details.Length -gt 350){$details=$details.Substring(0,350)}
  throw "Table REST $method $path failed (HTTP $status). $details $($_.Exception.Message)"
 }
}
function Get-StatusCode($errorRecord) {
 $message=[string]$errorRecord.Exception.Message
 if($message -match 'HTTP (\d{3})'){return [int]$Matches[1]}
 try{return [int]$errorRecord.Exception.Response.StatusCode}catch{return 0}
}
function Ensure-Table([string]$name) {
 Log "Checking Azure table: $name"
 try {
  $existing=Get-AzStorageTable -Name $name -Context $script:StorageContext -ErrorAction Stop
  if($existing){Log "Table exists: $name";return}
 } catch {
  if((Get-StatusCode $_) -ne 404 -and $_.Exception.Message -notmatch 'not found|does not exist') {throw "Checking table $name failed: $($_.Exception.Message)"}
 }
 try {
  $null=New-AzStorageTable -Name $name -Context $script:StorageContext -ErrorAction Stop
  Log "Created Azure table: $name"
 } catch {
  if((Get-StatusCode $_) -eq 409){Log "Table already exists: $name"}
  else {throw "Creating table $name failed: $($_.Exception.Message)"}
 }
}
function Read-Registration([string]$name) {
 $key=[uri]::EscapeDataString($name.Replace("'","''"))
 Log "Reading registry entry: $name"
 try {return Invoke-TableApi 'GET' "TelemetryToolsConfig(PartitionKey='Tool',RowKey='$key')"}
 catch {if((Get-StatusCode $_) -eq 404){return $null};throw}
}
function Save-Registration($c,$existing) {
 $entity=@{PartitionKey='Tool';RowKey=$c.Name;TableName=$c.Table;MinutesSaved=[int]$c.Minutes;Enabled=[bool]$c.Enabled;IncludeInTotals=[bool]$c.IncludeInTotals}
 $key=[uri]::EscapeDataString($c.Name)
 if($existing){
  Log "Updating registry entry: $($c.Name)"
  $null=Invoke-TableApi 'PUT' "TelemetryToolsConfig(PartitionKey='Tool',RowKey='$key')" $entity @{'If-Match'='*'}
 }else{
  Log "Creating registry entry: $($c.Name)"
  $null=Invoke-TableApi 'POST' 'TelemetryToolsConfig' $entity
 }
 Log "Saved registry entity: $($c.Name)"
}
function Refresh-Preview {
 try{$c=Inputs; $txtTable.Text=$c.Table; $txtCode.Text=Get-TelemetrySnippet $c; $script:Saved=$false; $lblSave.Text='NOT SAVED'; $lblSave.ForeColor=$script:Muted}
 catch{$txtCode.Text="Input validation: $($_.Exception.Message)"}
}
function Get-TelemetrySnippet($config) {
 $template=@'
# =====================================================
#region Telemetry Information
# =====================================================
$script:TelemetryReportID    = [guid]::NewGuid().Guid
$script:TelemetryGeoResolved = $false
$script:TelemetryGeoData     = @{}
$script:TelemetryStartupSent = $false
$script:uploadToAzure        = $true

function Write-Indent {
    param([string]$Message, [int]$Level = 1, [string]$Color = 'Gray')
    try { Write-Host (('  ' * $Level) + $Message) -ForegroundColor $Color } catch {}
}

function Get-TelemetryMachineHash {
    try {
        $raw = "$env:USERDOMAIN\$env:USERNAME@$env:COMPUTERNAME"
        $sha = [System.Security.Cryptography.SHA256]::Create()
        try {
            $bytes = [System.Text.Encoding]::UTF8.GetBytes($raw)
            $hash = $sha.ComputeHash($bytes)
            return ([BitConverter]::ToString($hash)).Replace('-', '').Substring(0,24)
        }
        finally { if ($sha) { $sha.Dispose() } }
    }
    catch { return '' }
}

function Resolve-TelemetryGeo {
    if ($script:TelemetryGeoResolved) { return }
    try {
        try {
            # Prefer local Windows region; no external call needed.
            $r = [System.Globalization.RegionInfo]::CurrentRegion
            if (-not $r.TwoLetterISORegionName -or $r.TwoLetterISORegionName -eq 'IV') {
                throw 'Local region unavailable.'
            }
            $script:TelemetryGeoData = @{
                country     = [string]$r.EnglishName
                countryCode = [string]$r.TwoLetterISORegionName
                timezone    = [string](Get-TimeZone).Id
            }
        }
        catch {
            # Fallback to ipwho.is only if local region resolution fails.
            if (-not $script:GeoCache) {
                $script:GeoCache = Invoke-RestMethod -Uri 'https://ipwho.is/' -TimeoutSec 5 -ErrorAction Stop
            }
            $r = $script:GeoCache
            if ($r.success -ne $true) { throw 'Geolocation service returned no result.' }
            $script:TelemetryGeoData = @{
                country     = [string]$r.country
                countryCode = [string]$r.country_code
                timezone    = [string]$r.timezone.id
            }
        }
    }
    catch {
        $tz = $null
        try { $tz = [string](Get-TimeZone).Id } catch {}
        $script:TelemetryGeoData = @{ country=$null; countryCode=$null; timezone=$tz }
    }
    finally { $script:TelemetryGeoResolved = $true }
}

function Send-ToolTelemetry {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)][string]$TelemetryName,
        [Parameter(Mandatory=$true)][string]$EventName,
        [Parameter(Mandatory=$true)][string]$Version,
        [Parameter(Mandatory=$true)][string]$Endpoint,
        [int]$ServerCount = 0,
        [int]$GroupCount = 0,
        [switch]$NoGeo,
        [switch]$DebugTelemetry
    )
    if (-not $script:uploadToAzure) { return }
    if ($EventName -match '^(Startup|Launch|AppStart|ToolStart|TelemetryStartup)$') {
        if ($script:TelemetryStartupSent) { return }
        $script:TelemetryStartupSent = $true
    }
    try {
        if (-not $NoGeo) { Resolve-TelemetryGeo }
        $data = [ordered]@{
            PartitionKey = ($TelemetryName -replace 'TelemetryData$','')
            RowKey       = [guid]::NewGuid().Guid
            PSVersion    = $PSVersionTable.PSVersion.ToString()
            countryCode  = $script:TelemetryGeoData.countryCode
            MachineHash  = Get-TelemetryMachineHash
            Version      = $Version
            timezone     = $script:TelemetryGeoData.timezone
            ReportID     = $script:TelemetryReportID
            country      = $script:TelemetryGeoData.country
        }
        # The Azure Function writes only Data into the telemetry table.
        $payload = @{ TelemetryName=$TelemetryName; TableName=$TelemetryName; Data=$data }
        $body = $payload | ConvertTo-Json -Depth 10
        if ($DebugTelemetry) { Write-Host "Telemetry Request: $body" -ForegroundColor Cyan }
        $response = Invoke-RestMethod -Method Post -Uri $Endpoint `
            -ContentType 'application/json' -Body $body -TimeoutSec 15 -ErrorAction Stop
        if ($DebugTelemetry) {
            Write-Host 'Telemetry Response:' -ForegroundColor Green
            $response | ConvertTo-Json -Depth 10 | Write-Host
        }
    }
    catch {
        # Never interrupt the primary tool because telemetry failed.
        if ($DebugTelemetry) {
            Write-Warning "Telemetry failed: $($_.Exception.Message)"
            if ($_.ErrorDetails.Message) { Write-Warning $_.ErrorDetails.Message }
        }
    }
}

# Place this invocation after your application's version has been initialized.
$telemetryParams = @{
    TelemetryName  = '__TABLE__'
    EventName      = 'Startup'
    Version        = __VERSION__
    Endpoint       = '__ENDPOINT__'
    DebugTelemetry = $false
}
Send-ToolTelemetry @telemetryParams
#endregion Telemetry Information
'@
 return $template.Replace('__TABLE__',$config.Table).Replace('__VERSION__',$config.Version).Replace('__ENDPOINT__',$script:Endpoint)
}

$form=New-Object Windows.Forms.Form
$form.Text='Telemetry Tool Manager | v4.1'; $form.Size=New-Object Drawing.Size(1180,1030);$form.MinimumSize=New-Object Drawing.Size(980,940)
$form.StartPosition='CenterScreen';$form.BackColor=$script:Bg;$form.ForeColor=$script:Ink;$form.Font=New-Object Drawing.Font('Segoe UI',10)
$header=Label 'TELEMETRY TOOL MANAGER' 26 18 650 42 22 $true;$form.Controls.Add($header)
$sub=Label 'Microsoft GSE  /  Azure Table Storage  /  One-place tool registration' 29 62 780 26 10;$sub.ForeColor=$script:Muted;$form.Controls.Add($sub)
$btnConnect=Button 'Connect to Microsoft GSE' 850 24 285;$form.Controls.Add($btnConnect);$btnConnect.Anchor='Top,Right'
$form.Controls.Add((Label 'GSE Tenant (fixed)' 29 98 145 27 10))
$txtTenant=Input 175 94 390 $script:TenantId
$txtTenant.Font=New-Object Drawing.Font('Consolas',10)
$txtTenant.ReadOnly=$true
$form.Controls.Add($txtTenant)
$lblConnect=Label 'Not connected' 29 128 1070 26 10;$lblConnect.ForeColor=$script:Muted;$form.Controls.Add($lblConnect)
$left=New-Object Windows.Forms.Panel;$left.SetBounds(25,162,365,604);$left.BackColor=$script:Panel;$left.Anchor='Top,Bottom,Left';$form.Controls.Add($left)
$left.Controls.Add((Label 'NEW TOOL REGISTRATION' 18 18 325 32 13 $true))
$left.Controls.Add((Label 'Tool name' 18 69 320));$txtName=Input 18 95 324 'ECHO';$left.Controls.Add($txtName)
$lblNoSpaces=Label 'No spaces - spaces are removed automatically' 18 127 330 18 8
$lblNoSpaces.ForeColor=$script:Muted;$left.Controls.Add($lblNoSpaces)
$left.Controls.Add((Label 'Azure telemetry table (automatic)' 18 144 330));$txtTable=Input 18 170 324 'ECHOTelemetryData';$txtTable.ReadOnly=$true;$left.Controls.Add($txtTable)
$left.Controls.Add((Label 'Minutes saved per run' 18 218 325));$txtMinutes=Input 18 245 324 '60';$left.Controls.Add($txtMinutes)
$left.Controls.Add((Label 'Version expression (variable / hashtable)' 18 293 325));$txtVersion=Input 18 320 324 '$Ver';$left.Controls.Add($txtVersion)
$chkEnabled=New-Object Windows.Forms.CheckBox;$chkEnabled.Text='Show in telemetry dashboard';$chkEnabled.SetBounds(18,373,324,30);$chkEnabled.Checked=$true;$chkEnabled.ForeColor=$script:Ink;$left.Controls.Add($chkEnabled)
$chkTotals=New-Object Windows.Forms.CheckBox;$chkTotals.Text='Include in All Tools totals';$chkTotals.SetBounds(18,407,324,30);$chkTotals.Checked=$true;$chkTotals.ForeColor=$script:Ink;$left.Controls.Add($chkTotals)
$btnPreview=Button 'Preview Code' 18 461 150 $true;$left.Controls.Add($btnPreview)
$btnSave=Button 'Create / Update Tool' 18 510 324;$left.Controls.Add($btnSave)
$btnManage=Button 'Manage Existing Tools' 18 559 324 $true;$left.Controls.Add($btnManage)
$lblSave=Label 'NOT SAVED' 180 467 160 24 9 $true;$lblSave.ForeColor=$script:Muted;$left.Controls.Add($lblSave)
$right=New-Object Windows.Forms.Panel;$right.SetBounds(408,162,727,604);$right.BackColor=$script:Panel;$right.Anchor='Top,Bottom,Left,Right';$form.Controls.Add($right)
$right.Controls.Add((Label 'POWERSHELL TELEMETRY CODE' 18 18 510 30 13 $true))
$btnCopy=Button 'Copy Code' 562 14 145;$btnCopy.Anchor='Top,Right';$right.Controls.Add($btnCopy)
$info=Label 'Paste after version initialization. Supports $Ver, $script:Version, $Config.Version, $Config[''Version''].' 18 54 680 26 9;$info.ForeColor=$script:Muted;$right.Controls.Add($info)
$txtCode=New-Object Windows.Forms.TextBox;$txtCode.SetBounds(18,86,689,475);$txtCode.Multiline=$true;$txtCode.ReadOnly=$true;$txtCode.ScrollBars='Both';$txtCode.WordWrap=$false
$txtCode.BackColor=$script:InputBg;$txtCode.ForeColor=$script:Ink;$txtCode.Font=New-Object Drawing.Font('Consolas',9);$txtCode.Anchor='Top,Bottom,Left,Right';$right.Controls.Add($txtCode)
$logLabel=Label 'ACTIVITY / VALIDATION' 27 800 360 25 11 $true;$logLabel.Anchor='Bottom,Left';$form.Controls.Add($logLabel)
$txtLog=New-Object Windows.Forms.TextBox;$txtLog.SetBounds(25,830,1110,140);$txtLog.Multiline=$true;$txtLog.ReadOnly=$true;$txtLog.ScrollBars='Vertical';$txtLog.BackColor=$script:InputBg;$txtLog.ForeColor=$script:Ink;$txtLog.Font=New-Object Drawing.Font('Consolas',11);$txtLog.WordWrap=$true;$txtLog.Anchor='Bottom,Left,Right';$form.Controls.Add($txtLog)
# Azure connection executes on a worker runspace. UI timer is the ONLY writer to WinForms controls.
$script:ConnectWorker=$null
$script:ConnectAsync=$null
$script:ConnectState=$null
$script:ConnectStarted=$null
$script:ConnectTimedOut=$false
$script:ConnectTimer=New-Object Windows.Forms.Timer
$script:ConnectTimer.Interval=200
$script:ConnectTimer.Add_Tick({
 if(-not $script:ConnectState){return}
 $messages=@()
 [System.Threading.Monitor]::Enter($script:ConnectState.SyncRoot)
 try {
  while($script:ConnectState.Messages.Count -gt 0){$messages+= [string]$script:ConnectState.Messages.Dequeue()}
 } finally {[System.Threading.Monitor]::Exit($script:ConnectState.SyncRoot)}
 foreach($msg in $messages){Log $msg}
 if(-not $script:ConnectTimedOut -and ((Get-Date)-$script:ConnectStarted).TotalSeconds -ge 60){
  $script:ConnectTimedOut=$true
  $lblConnect.Text='Connection timed out; worker may still be finishing'
  $lblConnect.ForeColor=[Drawing.Color]::OrangeRed
  Log 'ERROR: Connection exceeded 60 seconds. GUI remains usable. Do not retry until worker completes.'
 }
 if(-not $script:ConnectAsync -or -not $script:ConnectAsync.IsCompleted){return}
 $script:ConnectTimer.Stop()
 try {
  $null=$script:ConnectWorker.EndInvoke($script:ConnectAsync)
  if($script:ConnectState.Error){throw [Exception]::new([string]$script:ConnectState.Error)}
  if(-not $script:ConnectState.Key){throw 'Connection finished without an account key.'}
  $script:StorageKey=[string]$script:ConnectState.Key
  $script:StorageContext=New-AzStorageContext -StorageAccountName $script:StorageName -StorageAccountKey $script:StorageKey -ErrorAction Stop
  $script:Account=$true
  $lblConnect.Text="CONNECTED | $($script:SubscriptionId) | GSETools / gsetools"
  $lblConnect.ForeColor=$script:Green
  Log 'Connected successfully. Account-key context ready; table access will be checked when saving.'
 } catch {
  $script:Account=$null;$script:StorageContext=$null;$script:StorageKey=$null
  $lblConnect.Text='Connection failed';$lblConnect.ForeColor=[Drawing.Color]::OrangeRed
  Log "ERROR: $($_.Exception.Message)"
 } finally {
  $script:ConnectState.Key=$null
  $script:ConnectWorker.Dispose();$script:ConnectWorker=$null;$script:ConnectAsync=$null
  $script:ConnectState=$null
  $btnConnect.Enabled=$true;$btnConnect.Text='Connect to Microsoft GSE'
 }
})
$btnConnect.Add_Click({
 if(-not $btnConnect.Enabled){return}
 $btnConnect.Enabled=$false;$btnConnect.Text='Connecting...'
 $lblConnect.Text='Connecting in background...';$lblConnect.ForeColor=$script:Muted
 try {
  # Import the modules on the UI runspace only once. Existing installs are reused.
  Log 'Preparing Azure modules...'
  Initialize-AzureModules
  $script:ConnectState=[hashtable]::Synchronized(@{
   Messages=[System.Collections.Queue]::Synchronized((New-Object System.Collections.Queue));
   Key=$null;Error=$null
  })
  $script:ConnectWorker=[powershell]::Create()
  $workerCode={
   param($state,$tenant,$subscription,$rg,$storage)
   function Report([string]$message){$state.Messages.Enqueue($message)}
   try {
    Import-Module Az.Accounts,Az.Storage -ErrorAction Stop
    Report 'Checking current GSE session...'
    $ctx=Get-AzContext -ErrorAction SilentlyContinue
    if(-not $ctx -or [string]$ctx.Tenant.Id -ne $tenant -or [string]$ctx.Subscription.Id -ne $subscription){
     Report 'Signing in to the fixed GSE tenant (Microsoft sign-in may appear)...'
     $null=Connect-AzAccount -Tenant $tenant -Subscription $subscription -ErrorAction Stop
    } else {Report 'Reusing the GSE tenant/subscription session.'}
    Report 'Selecting GSE subscription...'
    $null=Set-AzContext -SubscriptionId $subscription -TenantId $tenant -ErrorAction Stop
    Report 'Locating gsetools storage account...'
    $account=Get-AzStorageAccount -ResourceGroupName $rg -Name $storage -ErrorAction Stop
    if(-not $account){throw 'Storage account not found.'}
    Report 'Retrieving gsetools account key...'
    $keys=@(Get-AzStorageAccountKey -ResourceGroupName $rg -Name $storage -ErrorAction Stop)
    if(-not $keys.Count -or -not $keys[0].Value){throw 'No storage account key returned.'}
    $state.Key=[string]$keys[0].Value
    Report 'Storage key retrieved. Skipping blocking table-list verification.'
   } catch {$state.Error=$_.Exception.Message}
  }
  $null=$script:ConnectWorker.AddScript($workerCode.ToString()).AddArgument($script:ConnectState).AddArgument($script:TenantId).AddArgument($script:SubscriptionId).AddArgument($script:ResourceGroup).AddArgument($script:StorageName)
  $script:ConnectStarted=Get-Date;$script:ConnectTimedOut=$false
  $script:ConnectAsync=$script:ConnectWorker.BeginInvoke()
  $script:ConnectTimer.Start()
  Log 'Azure connection started in background; the window will remain responsive.'
 } catch {
  Log "ERROR starting connection: $($_.Exception.Message)"
  $btnConnect.Enabled=$true;$btnConnect.Text='Connect to Microsoft GSE'
  $lblConnect.Text='Connection failed'
 }
})
$btnPreview.Add_Click({Refresh-Preview})
$script:NormalizingToolName=$false
$txtName.Add_TextChanged({
 if($script:NormalizingToolName){return}
 $original=$txtName.Text
 $clean=[regex]::Replace($original,'\s+','')
 if($clean -cne $original){
  $script:NormalizingToolName=$true
  try {
   $caret=$txtName.SelectionStart
   $removed=[regex]::Replace($original.Substring(0,[Math]::Min($caret,$original.Length)),'\s+','').Length
   $txtName.Text=$clean
   $txtName.SelectionStart=[Math]::Min($removed,$clean.Length)
  } finally {$script:NormalizingToolName=$false}
 }
 Refresh-Preview
})
$txtMinutes.Add_TextChanged({Refresh-Preview});$txtVersion.Add_TextChanged({Refresh-Preview})
$chkEnabled.Add_CheckedChanged({Refresh-Preview});$chkTotals.Add_CheckedChanged({Refresh-Preview})
$btnCopy.Add_Click({
 if($txtCode.Text -like 'Input validation:*'){[Windows.Forms.MessageBox]::Show('Fix input validation first.','Invalid inputs')|Out-Null;return}
 [Windows.Forms.Clipboard]::SetText($txtCode.Text);Log 'Copied PowerShell telemetry code to clipboard.'
})
# Registry and table operations run off the UI thread, including reads and writes.
$script:SaveWorker=$null
$script:SaveAsync=$null
$script:SaveState=$null
$script:SaveStarted=$null
$script:SaveTimedOut=$false
$script:SaveTimer=New-Object Windows.Forms.Timer
$script:SaveTimer.Interval=200
$script:SaveTimer.Add_Tick({
 if(-not $script:SaveState){return}
 $messages=@()
 while($script:SaveState.Messages.Count -gt 0){$messages += [string]$script:SaveState.Messages.Dequeue()}
 foreach($msg in $messages){Log $msg}
 if(-not $script:SaveTimedOut -and ((Get-Date)-$script:SaveStarted).TotalSeconds -gt 90){
  $script:SaveTimedOut=$true
  Log 'WARN: Azure registration has taken over 90 seconds. The background operation may still be running.'
 }
 if(-not $script:SaveAsync -or -not $script:SaveAsync.IsCompleted){return}
 $script:SaveTimer.Stop()
 try {
  $null=$script:SaveWorker.EndInvoke($script:SaveAsync)
  if($script:SaveState.Error){throw [Exception]::new([string]$script:SaveState.Error)}
  if(-not $script:SaveState.Success){throw 'Background registration ended without confirmation.'}
  $script:Saved=$true;$lblSave.Text='SAVED TO AZURE';$lblSave.ForeColor=$script:Green
  Log "Verified: $($script:SaveState.Tool) -> $($script:SaveState.Table)"
  [Windows.Forms.MessageBox]::Show("$($script:SaveState.Tool) is registered in TelemetryToolsConfig and its telemetry table exists.`n`nCopy the PowerShell code into the tool. The dashboard will show it after telemetry arrives and the summary updater runs.",'Registration saved') | Out-Null
 } catch {
  Log "ERROR: $($_.Exception.Message)"
  [Windows.Forms.MessageBox]::Show($_.Exception.Message,'Unable to save') | Out-Null
 } finally {
  $script:SaveWorker.Dispose();$script:SaveWorker=$null;$script:SaveAsync=$null;$script:SaveState=$null
  $btnSave.Enabled=$true;$btnSave.Text='Create / Update Tool';$form.Cursor='Default'
 }
})
$btnSave.Add_Click({
 try {
  $c=Inputs
  if(-not $script:Account -or -not $script:StorageKey){throw 'Connect to Microsoft GSE first.'}
  if($script:SaveWorker){throw 'A registration operation is already running.'}
  $message="Create or update tool '$($c.Name)'?`n`nTelemetry table: $($c.Table)`nMinutes saved: $($c.Minutes)`nEnabled: $($c.Enabled)`nInclude in totals: $($c.IncludeInTotals)`n`nThis will change live Azure Storage."
  if([Windows.Forms.MessageBox]::Show($message,'Confirm Azure changes','YesNo','Question') -ne 'Yes'){return}
  $btnSave.Enabled=$false;$btnSave.Text='Saving in background...'
  $script:Saved=$false;$lblSave.Text='SAVING...';$lblSave.ForeColor=$script:Muted
  $script:SaveState=[hashtable]::Synchronized(@{
   Messages=[System.Collections.Queue]::Synchronized((New-Object System.Collections.Queue));
   Error=$null;Success=$false;Tool=$c.Name;Table=$c.Table
  })
  $script:SaveWorker=[powershell]::Create()
  $saveCode={
   param($state,$account,$key,$config)
   function Report([string]$message){$state.Messages.Enqueue($message)}
   function Request([string]$method,[string]$path,$payload=$null,[hashtable]$extra=@{}) {
    $date=[DateTime]::UtcNow.ToString('R',[Globalization.CultureInfo]::InvariantCulture)
    # Azure Table SharedKeyLite canonical resource excludes query string.
    $resourcePath=$path.Split('?')[0]
    $canonical="/$account/$resourcePath"
    $toSign="$date`n$canonical"
    $hmac=[System.Security.Cryptography.HMACSHA256]::new()
    try {
     $hmac.Key=[Convert]::FromBase64String($key)
     $signature=[Convert]::ToBase64String($hmac.ComputeHash([Text.Encoding]::UTF8.GetBytes($toSign)))
    } finally {$hmac.Dispose()}
    $headers=@{
     Authorization="SharedKeyLite ${account}:$signature"
     'x-ms-date'=$date
     'x-ms-version'='2019-02-02'
     Accept='application/json;odata=nometadata'
     DataServiceVersion='3.0;NetFx'
     MaxDataServiceVersion='3.0;NetFx'
    }
    foreach($h in $extra.Keys){$headers[$h]=$extra[$h]}
    $args=@{Uri="https://$account.table.core.windows.net/$path";Method=$method;Headers=$headers;TimeoutSec=15;ErrorAction='Stop'}
    if($null -ne $payload){$args.ContentType='application/json';$args.Body=($payload | ConvertTo-Json -Depth 8 -Compress)}
    try {return Invoke-RestMethod @args}
    catch {
     $status=0
     try {$status=[int]$_.Exception.Response.StatusCode} catch {}
     $details=[string]$_.ErrorDetails.Message
     if($details.Length -gt 450){$details=$details.Substring(0,450)}
     throw "HTTP $status during $method $path. $details $($_.Exception.Message)"
    }
   }
   function Status($record){if($record.Exception.Message -match '^HTTP (\d{3})'){return [int]$Matches[1]};return 0}
   function EnsureTable([string]$table){
    Report "Checking table $table..."
    try {$null=Request 'GET' "Tables('$table')";Report "Table exists: $table"}
    catch {
     if((Status $_) -ne 404){throw}
     Report "Creating table $table..."
     try {$null=Request 'POST' 'Tables' @{TableName=$table};Report "Created table: $table"}
     catch {if((Status $_) -ne 409){throw};Report "Table already exists: $table"}
    }
   }
   try {
    $name=[string]$config.Name
    $table=[string]$config.Table
    EnsureTable 'TelemetryToolsConfig'
    EnsureTable $table
    $row=[uri]::EscapeDataString($name.Replace("'","''"))
    $path="TelemetryToolsConfig(PartitionKey='Tool',RowKey='$row')"
    Report "Checking existing registry entry: $name..."
    $exists=$false
    try {$null=Request 'GET' $path;$exists=$true}
    catch {if((Status $_) -ne 404){throw}}
    $entity=@{PartitionKey='Tool';RowKey=$name;TableName=$table;MinutesSaved=[int]$config.Minutes;Enabled=[bool]$config.Enabled;IncludeInTotals=[bool]$config.IncludeInTotals}
    if($exists){
     Report "Updating registry entry: $name..."
     $null=Request 'PUT' $path $entity @{'If-Match'='*'}
    }else{
     Report "Creating registry entry: $name..."
     $null=Request 'POST' 'TelemetryToolsConfig' $entity
    }
    Report "Verifying saved registry entry: $name..."
    $saved=Request 'GET' $path
    if(-not $saved -or $saved.TableName -cne $table){throw 'Registry verification failed: TableName mismatch.'}
    $state.Success=$true
    Report "Registration verified for $name."
   } catch {$state.Error=$_.Exception.Message;Report "Registration failed: $($state.Error)"}
  }
  $null=$script:SaveWorker.AddScript($saveCode.ToString()).AddArgument($script:SaveState).AddArgument($script:StorageName).AddArgument($script:StorageKey).AddArgument($c)
  $script:SaveStarted=Get-Date;$script:SaveTimedOut=$false
  $script:SaveAsync=$script:SaveWorker.BeginInvoke()
  $script:SaveTimer.Start()
  Log "Started background registration for $($c.Name); GUI remains responsive."
 }catch{
  Log "ERROR: $($_.Exception.Message)"
  $btnSave.Enabled=$true;$btnSave.Text='Create / Update Tool'
  [Windows.Forms.MessageBox]::Show($_.Exception.Message,'Unable to save')|Out-Null
 }
})

# Manage registry visibility only. This never deletes tables/entities or alters telemetry ingestion.
$btnManage.Add_Click({
 if(-not $script:Account -or -not $script:StorageKey){[Windows.Forms.MessageBox]::Show('Connect to Microsoft GSE first.','Not connected')|Out-Null;return}
 $script:ManageDialog=New-Object Windows.Forms.Form
 $script:ManageDialog.Text='Manage Existing Tools | Dashboard Visibility'
 $script:ManageDialog.Size=New-Object Drawing.Size(760,640)
 $script:ManageDialog.MinimumSize=New-Object Drawing.Size(640,480)
 $script:ManageDialog.StartPosition='CenterParent';$script:ManageDialog.BackColor=$script:Bg;$script:ManageDialog.ForeColor=$script:Ink
 $script:ManageDialog.Font=New-Object Drawing.Font('Segoe UI',10)
 $heading=Label 'EXISTING TELEMETRY TOOLS' 20 16 680 32 15 $true;$script:ManageDialog.Controls.Add($heading)
 $explain=Label 'Enabled controls dashboard visibility and summary inclusion only. Telemetry continues recording.' 20 52 700 42 9
 $explain.ForeColor=$script:Muted;$script:ManageDialog.Controls.Add($explain)
 $script:ManageList=New-Object Windows.Forms.ListView
 $script:ManageList.SetBounds(20,102,700,390);$script:ManageList.View='Details';$script:ManageList.FullRowSelect=$true;$script:ManageList.MultiSelect=$false
 $script:ManageList.GridLines=$true;$script:ManageList.HideSelection=$false;$script:ManageList.BackColor=$script:InputBg;$script:ManageList.ForeColor=$script:Ink
 [void]$script:ManageList.Columns.Add('Tool',215);[void]$script:ManageList.Columns.Add('Dashboard',115)
 [void]$script:ManageList.Columns.Add('Telemetry table',240);[void]$script:ManageList.Columns.Add('Minutes',90)
 $script:ManageList.Anchor='Top,Bottom,Left,Right';$script:ManageDialog.Controls.Add($script:ManageList)
 $script:ManageRefresh=Button 'Refresh Tools' 20 510 160 $true;$script:ManageRefresh.Anchor='Bottom,Left';$script:ManageDialog.Controls.Add($script:ManageRefresh)
 $script:ManageToggle=Button 'Enable / Disable Selected' 195 510 240;$script:ManageToggle.Anchor='Bottom,Left';$script:ManageDialog.Controls.Add($script:ManageToggle)
 $script:ManageToggle.Enabled=$false
 $script:ManageStatus=Label 'Ready to load registry' 20 558 700 35 9;$script:ManageStatus.ForeColor=$script:Muted
 $script:ManageStatus.Anchor='Bottom,Left,Right';$script:ManageDialog.Controls.Add($script:ManageStatus)
 $script:ManageJob=$null;$script:ManageAsync=$null;$script:ManageState=$null
 $script:ManageTimer=New-Object Windows.Forms.Timer;$script:ManageTimer.Interval=200
 $script:ManageStartJob={
  param([string]$mode,[string]$name,[bool]$enabled)
  if($script:ManageJob){return}
  $script:ManageRefresh.Enabled=$false;$script:ManageToggle.Enabled=$false
  $script:ManageStatus.Text=if($mode -eq 'list'){'Loading TelemetryToolsConfig...'}else{"Updating $name..."}
  $script:ManageState=[hashtable]::Synchronized(@{Result=$null;Error=$null})
  $script:ManageJob=[powershell]::Create()
  $script:ManageWorkerCode={
   param($state,$account,$key,$mode,$name,$enabled)
   function CallTable([string]$method,[string]$path,$payload=$null,[hashtable]$extra=@{}){
    $date=[DateTime]::UtcNow.ToString('R',[Globalization.CultureInfo]::InvariantCulture)
    $resource="/$account/$($path.Split('?')[0])"
    $hmac=[System.Security.Cryptography.HMACSHA256]::new()
    try {$hmac.Key=[Convert]::FromBase64String($key);$sig=[Convert]::ToBase64String($hmac.ComputeHash([Text.Encoding]::UTF8.GetBytes("$date`n$resource")))}
    finally {$hmac.Dispose()}
    $headers=@{Authorization="SharedKeyLite ${account}:$sig";'x-ms-date'=$date;'x-ms-version'='2019-02-02';Accept='application/json;odata=nometadata';DataServiceVersion='3.0;NetFx';MaxDataServiceVersion='3.0;NetFx'}
    foreach($k in $extra.Keys){$headers[$k]=$extra[$k]}
    $args=@{Uri="https://$account.table.core.windows.net/$path";Method=$method;Headers=$headers;TimeoutSec=20;ErrorAction='Stop'}
    if($null -ne $payload){$args.ContentType='application/json';$args.Body=($payload|ConvertTo-Json -Depth 6 -Compress)}
    try {return Invoke-WebRequest @args -UseBasicParsing}
    catch {
     $code=0;try {$code=[int]$_.Exception.Response.StatusCode}catch{}
     throw "HTTP $code $method $path : $($_.Exception.Message) $($_.ErrorDetails.Message)"
    }
   }
   try {
    if($mode -eq 'toggle'){
     $escaped=[uri]::EscapeDataString($name.Replace("'","''"))
     $path="TelemetryToolsConfig(PartitionKey='Tool',RowKey='$escaped')"
     # MERGE changes only Enabled; never overwrites other properties or telemetry records.
     $null=CallTable 'MERGE' $path @{Enabled=$enabled} @{'If-Match'='*'}
    }
    $results=New-Object System.Collections.ArrayList
    $nextP=$null;$nextR=$null
    do {
     $path='TelemetryToolsConfig()?$filter='+[uri]::EscapeDataString("PartitionKey eq 'Tool'")
     if($nextP){$path+='&NextPartitionKey='+[uri]::EscapeDataString($nextP)}
     if($nextR){$path+='&NextRowKey='+[uri]::EscapeDataString($nextR)}
     $response=CallTable 'GET' $path
     $data=$response.Content|ConvertFrom-Json
     foreach($item in @($data.value)){
      if($null -ne $item -and [string]$item.PartitionKey -ceq 'Tool'){
       [void]$results.Add([pscustomobject]@{Name=[string]$item.RowKey;Table=[string]$item.TableName;Enabled=([string]$item.Enabled -eq 'True');Minutes=[string]$item.MinutesSaved})
      }
     }
     $nextP=[string]$response.Headers['x-ms-continuation-NextPartitionKey']
     $nextR=[string]$response.Headers['x-ms-continuation-NextRowKey']
    }while($nextP -or $nextR)
    $state.Result=@($results|Sort-Object Name)
   }catch{$state.Error=$_.Exception.Message}
  }
  $null=$script:ManageJob.AddScript($script:ManageWorkerCode.ToString()).AddArgument($script:ManageState).AddArgument($script:StorageName).AddArgument($script:StorageKey).AddArgument($mode).AddArgument($name).AddArgument($enabled)
  $script:ManageAsync=$script:ManageJob.BeginInvoke()
  $script:ManageTimer.Start()
 }
 $script:ManageTimer.Add_Tick({
  if(-not $script:ManageJob -or -not $script:ManageAsync){return}
  if(-not $script:ManageAsync.IsCompleted){return}
  $script:ManageTimer.Stop()
  try {
   $null=$script:ManageJob.EndInvoke($script:ManageAsync)
   if($script:ManageState.Error){throw $script:ManageState.Error}
   $script:ManageList.BeginUpdate();try {
    $script:ManageList.Items.Clear()
    foreach($tool in $script:ManageState.Result){
     $item=New-Object Windows.Forms.ListViewItem($tool.Name)
     [void]$item.SubItems.Add($(if($tool.Enabled){'Enabled'}else{'Disabled'}))
     [void]$item.SubItems.Add($tool.Table)
     [void]$item.SubItems.Add($tool.Minutes)
     $item.Tag=[bool]$tool.Enabled
     if(-not $tool.Enabled){$item.ForeColor=[Drawing.Color]::Silver}
     [void]$script:ManageList.Items.Add($item)
    }
   }finally{$script:ManageList.EndUpdate()}
   $script:ManageStatus.Text="Loaded $($script:ManageList.Items.Count) tools. Disabling hides from dashboard and summary; ingestion is unchanged."
  }catch{$script:ManageStatus.Text="ERROR: $($_.Exception.Message)";Log "Manage tools: $($_.Exception.Message)"}
  finally {
   $script:ManageJob.Dispose();$script:ManageJob=$null;$script:ManageAsync=$null
   $script:ManageRefresh.Enabled=$true;$script:ManageToggle.Enabled=($script:ManageList.SelectedItems.Count -eq 1)
  }
 })
 $script:ManageList.Add_SelectedIndexChanged({$script:ManageToggle.Enabled=(-not $script:ManageJob -and $script:ManageList.SelectedItems.Count -eq 1)})
 $script:ManageRefresh.Add_Click({& $script:ManageStartJob 'list' '' $false})
 $script:ManageToggle.Add_Click({
  if($script:ManageList.SelectedItems.Count -ne 1){return}
  $item=$script:ManageList.SelectedItems[0];$newEnabled=-not [bool]$item.Tag
  $action=if($newEnabled){'SHOW'}else{'HIDE'}
  $message="$action '$($item.Text)' on the telemetry dashboard?`n`nThis changes only the Enabled flag in TelemetryToolsConfig.`nNo telemetry data will be deleted and the tool will continue recording telemetry."
  if([Windows.Forms.MessageBox]::Show($message,'Confirm dashboard visibility','YesNo','Question') -ne 'Yes'){return}
  & $script:ManageStartJob 'toggle' $item.Text $newEnabled
 })
 $script:ManageDialog.Add_FormClosing({if($script:ManageJob){$_.Cancel=$true;$script:ManageStatus.Text='Please wait for the current registry request to finish.'}})
 $script:ManageDialog.Add_FormClosed({$script:ManageTimer.Stop();$script:ManageTimer.Dispose();$script:ManageTimer=$null})
 $script:ManageDialog.Show($form)
 & $script:ManageStartJob 'list' '' $false
})

$form.Add_Resize({
 $left.Height=[Math]::Max(550,$form.ClientSize.Height-350)
 $right.Height=$left.Height
 $txtCode.Height=[Math]::Max(400,$right.ClientSize.Height-115)
 $logLabel.Top=$form.ClientSize.Height-190
 $txtLog.Top=$form.ClientSize.Height-160
 $txtLog.Height=145
 $txtLog.Width=[Math]::Max(400,$form.ClientSize.Width-50)
})
Refresh-Preview
Log 'Ready. Connect to Microsoft GSE, preview, then create/update the registry entity.'
Log 'Registry-driven dashboard detected as already deployed. Run the summary updater after first telemetry is received.'
try { [void]$form.ShowDialog() }
finally {
    $script:ConnectTimer.Stop()
    $script:SaveTimer.Stop()
    if($script:SaveWorker -and $script:SaveAsync -and $script:SaveAsync.IsCompleted){
        try {$null=$script:SaveWorker.EndInvoke($script:SaveAsync)} catch {}
        $script:SaveWorker.Dispose()
    }
    if($script:ConnectWorker -and $script:ConnectAsync -and $script:ConnectAsync.IsCompleted){
        try {$null=$script:ConnectWorker.EndInvoke($script:ConnectAsync)} catch {}
        $script:ConnectWorker.Dispose()
    }
    $script:StorageKey=$null
    $script:StorageContext=$null
    if(-not $script:ConnectWorker -or $script:ConnectAsync.IsCompleted){Cleanup-AzureModules}
    $form.Dispose()
}
