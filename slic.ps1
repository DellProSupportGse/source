<#
.SYNOPSIS
    SLIC - Switch Log Integrity Checker
    Verifies switch configurations for S2D and Azure Local environments.

.DESCRIPTION
    This script analyzes and validates switch configuration data extracted from
    Dell OS10 network switches used in Storage Spaces Direct (S2D) or Azure 
    Local deployments. 

    The script compares parsed data from "show tech-support" files against the
    SDDC baseline configuration to identify deviations, missing settings, and
    compliance gaps of Dell switches.

    Both the switch "Show Tech-Support" output(s) and the SDDC reference data
    must be provided for the verification process to function.

.CREATEDBY
    Jim Gandy
.UPDATES
    2025/11/03:v1.0 - JG - Initial release
#>
Function Invoke-SLIC {
Function EndScript{  
    break
}
$Ver="v1.0"
$ToolName = @"
$Ver
  ___ _    ___ ___ 
 / __| |  |_ _/ __|
 \__ \ |__ | | (__ 
 |___/____|___\___|
 Switch Log Integrity Checker
            By: Jim Gandy
"@
cls
Write-Host $ToolName
Write-Host ""
Write-Host "⚠️ SLIC Compatibility Notice:"
Write-host "       This tool currently supports Azure Local and Windows Server S2D clusters only."
do {
    $run = Read-Host "Ready to run? [Y/N]"
    Write-Host ""

    if ($run -match '^[Yy]$') {
        Write-Host "Running script..."
        $confirmed = $true
    }
    elseif ($run -match '^[Nn]$') {
        Write-Host "Exiting script..."
        EndScript
        $confirmed = $true
    }
    else {
        Write-Host "Please enter Y or N."
        $confirmed = $false
    }

} until ($confirmed)

If($confirmed -eq $true){
    Function Get-FileName([string]$initialDirectory, [string]$infoTxt, [string]$filter) {
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null

    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog -Property @{MultiSelect = $true}
    $OpenFileDialog.Title = $infoTxt
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = $filter
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.filenames
    }

    Write-Host "Please Select Show Tech-Support File(s) to use..."
    $STSLOC = Get-FileName "$env:USERPROFILE\Documents\SRs" "Please Select Show Tech-Support File(s)." "Logs (*.zip,*.txt,*.log)| *.zip;*.TXT;*.log"
    If(!($STSLOC)){
        Write-Host "No logs provided. Exiting script..."
        EndScript
    }Else{
        Write-Host "✅ SwithcLogs:"$STSLOC
    }

    $SDDCPath = Read-Host "Please provide the path to the extracted SDDC"
    
    If(!(Test-Path $SDDCPath -ErrorAction SilentlyContinue)){
        Write-Host "SDDC path not found. Exiting script..." -ForegroundColor Red
        EndScript
    }Else{
        Write-Host "✅ SDDC Path:"$SDDCPath
    }

    #region === HTML Report System ===

    function New-HtmlReport {
        param (
            [string]$Title = "Switch Configuration Report",
            [string]$Version = "",
            [string]$RunDate = (Get-Date),
            [string]$OutputPath = "$env:TEMP\SwitchReport.html"
        )

        $script:HtmlReportSections = @()

$htmlStyle = @"
<style>
body { 
  font-family: Segoe UI, Arial, sans-serif; 
  margin: 10px 20px;
}
table { 
  border-collapse: collapse; 
  width: 100%; 
  margin: 4px 0; 
}
th, td { 
  border: 1px solid #444; 
  padding: 3px 5px; 
}
th { 
  background-color: #6495ED; 
  color: white; 
  cursor: pointer; 
  user-select: none; 
}
tr:nth-child(even) { background-color: #f2f2f2; }
tr:hover td { background-color: #C1D5F8; }
h1 { 
  border-bottom: 4px solid #f4a460; 
  padding-bottom: 4px; 
  margin-bottom: 10px;
}
h2 { 
  margin: 6px 0 2px 0; 
  font-size: 1.1em;
}
h5 { 
  margin: 2px 0; 
  font-weight: normal; 
}
mark { background-color: yellow; color: black; }
.toggle { cursor: pointer; color: #0078d4; font-weight: bold; }
.hidden { display: none; }
div[id^='section'] { margin-bottom: 4px; }
.reset-btn {
  background-color: #f4a460;
  color: white;
  border: none;
  border-radius: 4px;
  padding: 2px 6px;
  margin-bottom: 4px;
  cursor: pointer;
  font-size: 0.8em;
}
.reset-btn:hover { background-color: #d78c3b; }
</style>

<script>
// --- Toggle sections ---
function toggleSection(id, elem) {
  var x = document.getElementById(id);
  if (x.style.display === 'none') { 
    x.style.display = 'block'; 
    elem.innerText = '▼'; 
  } else { 
    x.style.display = 'none'; 
    elem.innerText = '▶'; 
  }
}

// --- Multi-column sortable tables (first row fixed) ---
document.addEventListener('DOMContentLoaded', function () {
  document.querySelectorAll('table').forEach(function (tbl) {
    var headers = tbl.querySelectorAll('th');
    var sortState = [];

    // Insert Reset Sort button above each table
    var resetBtn = document.createElement('button');
    resetBtn.innerText = 'Reset Sort';
    resetBtn.className = 'reset-btn';
    resetBtn.onclick = function() {
      sortState = [];
      var tbody = tbl.querySelector('tbody');
      if (!tbody) return;
      var rows = Array.from(tbody.querySelectorAll('tr'));
      if (rows.length === 0) return;

      // Keep first row as header
      var headerRow = rows.shift();

      // Restore original order
      rows.sort(function(a, b) {
        return a.dataset.originalIndex - b.dataset.originalIndex;
      });

      tbody.innerHTML = '';
      tbody.appendChild(headerRow);
      rows.forEach(r => tbody.appendChild(r));

      // Clear arrows
      headers.forEach(h => h.innerText = h.innerText.replace(/[▲▼]/g, '').trim());
    };
    tbl.parentNode.insertBefore(resetBtn, tbl);

    // Add original index tracking
    var tbody = tbl.querySelector('tbody');
    if (tbody) {
      var allRows = Array.from(tbody.querySelectorAll('tr'));
      allRows.forEach((r, i) => r.dataset.originalIndex = i);
    }

    headers.forEach(function (th, colIndex) {
      th.addEventListener('click', function () {
        var existing = sortState.findIndex(s => s.col === colIndex);
        if (existing > -1) {
          sortState[existing].asc = !sortState[existing].asc;
        } else {
          sortState = sortState.concat({ col: colIndex, asc: true });
        }

        var tbody = tbl.querySelector('tbody');
        if (!tbody) return;

        var rows = Array.from(tbody.querySelectorAll('tr'));
        if (rows.length === 0) return;

        // Exclude first row (header row inside tbody)
        var headerRow = rows.shift();

        // Sort remaining rows
        rows.sort(function (a, b) {
          for (var i = 0; i < sortState.length; i++) {
            var s = sortState[i];
            var A = a.children[s.col]?.innerText.trim() ?? '';
            var B = b.children[s.col]?.innerText.trim() ?? '';
            var cmp = 0;
            if (!isNaN(A - B)) cmp = (A - B);
            else cmp = A.localeCompare(B, undefined, { numeric: true, sensitivity: 'base' });
            if (cmp !== 0) return cmp * (s.asc ? 1 : -1);
          }
          return 0;
        });

        // Rebuild tbody
        tbody.innerHTML = '';
        tbody.appendChild(headerRow);
        rows.forEach(r => tbody.appendChild(r));

        // Update sort arrows
        headers.forEach(h => h.innerText = h.innerText.replace(/[▲▼]/g, '').trim());
        sortState.forEach(s => {
          headers[s.col].innerText += s.asc ? ' ▲' : ' ▼';
        });
      });
    });
  });

  // --- Auto-collapse sections with no highlighted cells ---
  document.querySelectorAll('h2').forEach(function (hdr) {
    const toggleIcon = hdr.querySelector('.toggle');
    const sectionId = toggleIcon ? toggleIcon.getAttribute('onclick').match(/'(.*?)'/)[1] : null;
    if (!sectionId) return;

    const sectionDiv = document.getElementById(sectionId);
    if (!sectionDiv) return;

    const hasHighlight = sectionDiv.querySelector('td[style*="ff0000"], td[style*="ffff00"]');
    if (!hasHighlight) {
      sectionDiv.style.display = 'none';
      if (toggleIcon) toggleIcon.innerText = '▶';
      sectionDiv.style.marginBottom = '2px';
    }
  });
});
</script>
"@



$header = @"
<!DOCTYPE html>
<html>
<head>
<meta charset='UTF-8'>
<title>$Title</title>
$htmlStyle
<style>
.warning-banner {
  background-color: #fff3cd;
  color: #856404;
  border: 1px solid #ffeeba;
  border-radius: 6px;
  padding: 8px 12px;
  margin: 10px 0 20px 0;
  font-size: 14px;
}
</style>
</head>
<body>
<h1>$Title</h1>
<h3>&nbsp;Version: $Version</h3>
<h3>&nbsp;Run Date: $RunDate</h3>

<div class='warning-banner'>
  ⚠️ <b>SLIC Compatibility Notice:</b>
  This tool currently supports <b>Azure Local</b> and <b>Windows Server S2D</b> clusters only.
</div>
"@

$script:HtmlReportHeader = $header
$script:HtmlReportFooter = "</body></html>"
$script:HtmlReportPath = $OutputPath
    }

    function AddTo-HtmlReport {
        [CmdletBinding()]
        param (
    
        [Parameter(Mandatory)]
            [array]$Data,
            [string]$Title = "Report Section",
            [string]$Description = "",
            [string]$Footnotes = "",
            [switch]$IncludeTitle,
            [switch]$IncludeDescription,
            [switch]$IncludeFootnotes
        )

        begin {
            $html = ""
            $sectionId = ($Title -replace '\s','_')
        }

        process {
            if ($IncludeTitle) {
                $html += "<h2><span class='toggle' onclick=`"toggleSection('$sectionId',this)`">▼</span> $Title</h2>`n"
                $html += "<div id='$sectionId' style='display:block;'>"
            }

            if ($IncludeDiscription -and $Discription) {
                $html += "<h5><b>Discription:</b> $ShouldBe</h5>`n"
            }

            # Convert the input objects to HTML table fragment
            $html += ($Data | ConvertTo-Html -Fragment)



            if ($IncludeFootnotes -and $Footnotes) {
                $html += "<p><i>$Footnotes</i></p>`n"
            }

            if ($IncludeTitle) {
                $html += "</div>`n"
            }

            # Color marker replacements
            $html = $html `
                -replace '<td>RREEDD', '<td style="color:#fff;background-color:#ff0000">' `
                -replace '<td>YYEELLLLOOWW', '<td style="background-color:#ffff00">'
        }

        end {
            # Append this section to the global collection
            $script:HtmlReportSections += $html
        }
    }

    function Save-HtmlReport {
        if (-not $script:HtmlReportSections) {
            Write-Warning "No sections added to report."
            return
        }

        $finalHtml = $script:HtmlReportHeader + ($script:HtmlReportSections -join "`n") + $script:HtmlReportFooter
        
        Out-File -FilePath $script:HtmlReportPath -InputObject $finalHtml -Encoding UTF8
        Write-Host "✅ Report saved to: $script:HtmlReportPath"
        Invoke-Item $script:HtmlReportPath
    }

    #endregion === HTML Report System ===

    do {
        $saveChoice = Read-Host "Save report to the SDDC folder ($SDDCPath)? [Y/N]"

        if ($saveChoice -match '^[Yy]$') {
            $OutputPath = $SDDCPath
            if (Test-Path $OutputPath) {
                Write-Host "✅ Report will be saved in: $OutputPath"
                $confirmed = $true
            } else {
                Write-Host "❌ The SDDC folder path does not exist: $OutputPath"
                $confirmed = $false
            }
        }
        elseif ($saveChoice -match '^[Nn]$') {
            $OutputPath = Read-Host "Please type the full folder path where you want to save the report"
            if (Test-Path $OutputPath) {
                Write-Host "✅ Report will be saved in: $OutputPath"
                $confirmed = $true
            } else {
                Write-Host "❌ Invalid path. Please try again or choose Y to use the SDDC folder."
                $confirmed = $false
            }
        }
        else {
            Write-Host "Please enter Y or N."
            $confirmed = $false
        }

    } until ($confirmed)


    # Create new report
    $OutputPath = "$OutputPath\SLIC_Report_{0:yyyyMMdd_HHmm}.html" -f (Get-Date)
    New-HtmlReport -Title "SLIC: Switch Log Insight Checker" -Version "1.0" -RunDate (Get-Date) -OutputPath $OutputPath


    function Get-OS10RunningConfigSections {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory, Position=0)]
            [string]$Path
        )

        if (-not (Test-Path -LiteralPath $Path)) {
            throw "File not found: $Path"
        }

        # Read the whole file
        $text = Get-Content -LiteralPath $Path -Raw

        # Strip ANSI/VT100 escape sequences and CRs
        $text = [regex]::Replace($text, "\x1B\[[0-?]*[ -/]*[@-~]", "")
        $text = $text -replace "`r",""

        # Locate the "show running-configuration" section delimited by dashed headers
        # Matches:
        #   ----------------------------------- show running-configuration -------------------
        # then captures everything up to the next dashed "show ..." header or EOF.
        $pattern = '(?is)^\s*-{3,}\s*show\s+running-configuration\s*-{3,}\s*\n(.*?)(?=^\s*-{3,}\s*show\s+\S.*?-{3,}\s*$|\Z)'
        $m = [regex]::Match($text, $pattern, 'IgnoreCase, Multiline, Singleline')
        #$m | ?{$_ -imatch "hostname"} | select Filename,@{L="HostName";E={($_.lines -imatch "hostname") -replace "hostname "}}
        if (-not $m.Success) {
            throw "Could not locate the 'show running-configuration' section. Check header format in the log."
        }

        $run = $m.Groups[1].Value.Trim()

        # Split into sections where each section starts at a line that is just "!"
        # (Ignore blank chunks)
        $chunks = [regex]::Split($run,'(?m)^\s*!\s*$') | Where-Object { $_ -and $_.Trim() -ne '' }
        $SwitchHostname=""
        $SwitchHostname = (((($chunks | ?{$_ -imatch "hostname"}) -split "hostname ")[-1] -split "`n")[0]).trim()

        $sections = New-Object System.Collections.Generic.List[object]
        $i = 0
        foreach ($chunk in $chunks) {
            $i++
            $lines  = ($chunk -split "`n") | ForEach-Object { $_.TrimEnd() }
            $header = ($lines | Where-Object { $_.Trim() -ne '' } | Select-Object -First 1)
            if (-not $header) { $header = "(empty)" }

            $sections.Add([pscustomobject]@{
                FileName   = $Path.split("/\")[-1]
                SwHostName = $SwitchHostname
                Index      = $i
                Header     = $header.Trim()
                Lines      = $lines
                Text       = ($lines -join "`n")
            })
        }

        return $sections
        }

    Function Get-showversion{
        [CmdletBinding()]
        param(
            [Parameter(Mandatory, Position=0)]
            [string]$Path
        )

        if (-not (Test-Path -LiteralPath $Path)) {
            throw "File not found: $Path"
        }
        
        # Read the whole file
        $text = Get-Content -LiteralPath $Path -Raw

        # Strip ANSI/VT100 escape sequences and CRs
        $text = [regex]::Replace($text, "\x1B\[[0-?]*[ -/]*[@-~]", "")
        $text = $text -replace "`r",""

        # Locate the "show version" section delimited by dashed headers
        # Matches:
        # ----------------------------------- show version -------------------
        # then captures everything up to the next dashed "show ..." header or EOF.
        $pattern = '(?is)^\s*-{3,}\s*show\s+version\s*-{3,}\s*\n(.*?)(?=^\s*-{3,}\s*show\s+\S.*?-{3,}\s*$|\Z)'
        $m = [regex]::Match($text, $pattern, 'IgnoreCase, Multiline, Singleline')
        #$m | ?{$_ -imatch "hostname"} | select Filename,@{L="HostName";E={($_.lines -imatch "hostname") -replace "hostname "}}
        if (-not $m.Success) {
            throw "Could not locate the 'show running-configuration' section. Check header format in the log."
        }

        $run = $m.Groups[1].Value.Trim()
        $lines   = $run -split "`n"
        $ShowVersionOut = ""
        $ShowVersionOut = ([pscustomobject]@{
            FileName   = $Path.split("/\")[-1]
            SwHostName = $SwitchHostname
            OSVersion  = (($lines | ?{$_ -imatch "OS Version:"}) -split ": ")[-1]
            SystemType = (($lines | ?{$_ -imatch "System Type:"}) -split ": ")[-1]
            UpTime     = (($lines | ?{$_ -imatch "Up Time:"}) -split ": ")[-1]
        })
        return $ShowVersionOut
    }
    $ShowVersions = @()
    $ShowVersions += $STSLOC | ForEach-Object{Get-showversion -path $_}
    # Add to HTML report output sections
    if($ShowVersions){
        AddTo-HtmlReport -Title "Show Version" `
            -Data $ShowVersions `
            -Description "" `
            -Footnotes ""`
            -IncludeTitle -IncludeDescription -IncludeFootnotes
    }

        function Get-ShowLldpNeighbors {
            [CmdletBinding()]
            [OutputType([Object[]])]
            param(
                [Parameter(Mandatory, Position=0)]
                [string]$Path
            )

            if (-not (Test-Path -LiteralPath $Path)) {
                throw "File not found: $Path"
            }

            # Read whole file
            $text = Get-Content -LiteralPath $Path -Raw

            # Strip ANSI/VT100 and CRs
            $text = [regex]::Replace($text, "\x1B\[[0-?]*[ -/]*[@-~]", "")
            $text = $text -replace "`r",""
            $SwitchHostname=""
            $SwitchHostname = (((($text | ?{$_ -imatch "hostname"}) -split "hostname ")[-1] -split "`n")[0]).trim()

            # Grab the "show lldp neighbors" section delimited by dashed headers
            $pattern = '(?is)^\s*-{3,}\s*show\s+lldp\s+neighbors\s*-{3,}\s*\n(.*?)(?=^\s*-{3,}\s*show\s+\S.*?-{3,}\s*$|\Z)'
            $m = [regex]::Match($text, $pattern, 'IgnoreCase, Multiline, Singleline')
            if (-not $m.Success) {
                throw "Could not locate the 'show lldp neighbors' section. Check header format in the log."
            }

            $section = $m.Groups[1].Value.Trim()
            $lines   = $section -split "`n"

            # Regex: tolerate spaces inside "Rem Host Name"; require 2+ spaces between columns
            $rowRx = '^(?<LocPort>\S+)\s+(?<RemHost>.+?)\s{2,}(?<RemPort>.+?)\s{2,}(?<RemChassis>\S+)\s*$'

            $objects = foreach ($ln in $lines) {
                $t = $ln.TrimEnd()
                if (-not $t) { continue }
                if ($t -match '^-{3,}$') { continue }                                # underline row
                if ($t -match '^\s*Loc\s+PortID\s+Rem\s+Host\s+Name') { continue }    # header row

                $mx = [regex]::Match($t, $rowRx)
                if ($mx.Success) {
                    $locPort    = $mx.Groups['LocPort'].Value.Trim()
                    $remHost    = $mx.Groups['RemHost'].Value.Trim()
                    $remPort    = $mx.Groups['RemPort'].Value.Trim()
                    $remChassis = $mx.Groups['RemChassis'].Value.Trim().ToLower()

                    if ($remHost -match '^\s*Not\s+Advertised\s*$') { $remHost = $null }

                    [pscustomobject]@{
                        FileName         = $Path.split("/\")[-1]
                        SwHostName       = $SwitchHostname
                        LocPortId        = $locPort
                        RemoteHostName   = $remHost
                        RemotePortId     = $remPort
                        RemoteChassisId  = $remChassis
                    }
                }
            }

            return $objects
        }

    #region Show LLDP Neighbors

        $ShowLldpNeighbors=@()
         #$ShowLldpNeighbors += Get-ShowLldpNeighbors "C:\Users\Jim_Gandy\Downloads\tk5tor01-19c-show-tech-20251021-135024.txt"
         $ShowLldpNeighbors += $STSLOC | ForEach-Object { Get-ShowLldpNeighbors -Path $_ }

         Function Get-GetNetAdapterInfo {
            [CmdletBinding()]
            [OutputType([Object[]])]
            param(
                [Parameter(Mandatory, Position=0)]
                [string]$Path
            )

            if (-not (Test-Path -LiteralPath $Path)) {
                throw "File not found: $Path"
            }

            $NetAdaInfo = Get-ChildItem -Path $SDDCPath -Recurse -ErrorAction SilentlyContinue -Depth 2 -Filter getnetadapter.xml | Import-Clixml | select *,@{L="MADDR";E={$_.MacAddress -replace "-",":"}}
            return  $NetAdaInfo 

         }

         Function Get-GetNetIntents {
            [CmdletBinding()]
            [OutputType([Object[]])]
            param(
                [Parameter(Mandatory, Position=0)]
                [string]$Path
            )

            if (-not (Test-Path -LiteralPath $Path)) {
                throw "File not found: $Path"
            }

            $NetIntentsXml = Get-ChildItem -Path $SDDCPath -Recurse -ErrorAction SilentlyContinue -Depth 2 -Filter GetNetIntent.XML | Import-Clixml | select *,@{L="MADDR";E={$_.MacAddress -replace "-",":"}}
            return  $NetIntentsXml

         }

         $GetNetIntents = Get-GetNetIntents -path $SDDCPath
         $NetIntentStorageNicsInfo = $GetNetIntents| ?{$_.IsStorageIntentSet -eq $True} | select NetAdapterNamesAsList,StorageVLANs
            # Split and expand into separate objects
            $StorageNics = for ($i = 0; $i -lt $NetIntentStorageNicsInfo.NetAdapterNamesAsList.Count; $i++) {
                [pscustomobject]@{
                    NetAdapterName = $NetIntentStorageNicsInfo.NetAdapterNamesAsList[$i]
                    VLAN    = $NetIntentStorageNicsInfo.StorageVLANs[$i]
                }
            }

            # Display
            #$StorageNics | Format-Table

      
              $NetIntentMgmtNicsInfo = $GetNetIntents| ?{$_.IsManagementIntentSet -eq $True} | select NetAdapterNamesAsList,ManagementVLAN
            # Split and expand into separate objects
            $MgmtNics = for ($i = 0; $i -lt $NetIntentMgmtNicsInfo.NetAdapterNamesAsList.Count; $i++) {
                [pscustomobject]@{
                    NetAdapterName = $NetIntentMgmtNicsInfo.NetAdapterNamesAsList[$i]
                    VLAN    = $NetIntentMgmtNicsInfo.ManagementVLAN
                }
            }

            # Display
            #$MgmtNics | Format-Table


         $GetNetAdapterInfos = Get-GetNetAdapterInfo -path $SDDCPath
         $ShowRunningConfigs = $STSLOC | ForEach-Object { Get-OS10RunningConfigSections -Path $_ }
         $ShowRunningConfigs = $ShowRunningConfigs | ?{$_.hostname -ne "False"}

         #Matchup NetAdapters with lldp from the show tech   
            $SwPortToHostMap = @()

            foreach ($NetAdapter in $GetNetAdapterInfos) {

                # Ensure properties exist
                $NetAdapter | Add-Member -NotePropertyName IntentType -NotePropertyValue "" -Force
                $NetAdapter | Add-Member -NotePropertyName vLAN -NotePropertyValue "" -Force

                # Match Storage
                foreach ($StorageNic in $StorageNics) {
                    if ($NetAdapter.Name -eq $StorageNic.NetAdapterName) {
                        $NetAdapter.IntentType = "Storage"
                        $NetAdapter.vLAN       = $StorageNic.vLAN
                    }
                }

                # Match Management
                foreach ($MgmtNic in $MgmtNics) {
                    if ($NetAdapter.Name -eq $MgmtNic.NetAdapterName) {
                        $NetAdapter.IntentType = "Mgmt"
                        $NetAdapter.vLAN       = $MgmtNic.vLAN
                    }
                }

                # Match LLDP neighbor
                foreach ($lldpneighbor in $ShowLldpNeighbors) {
                    if ($lldpneighbor.RemotePortId -eq $NetAdapter.MADDR) {
                        $SwPortToHostMap += $lldpneighbor | Select-Object `
                            @{L="SwHostName";E={$_.SwHostName}},
                            @{L="SwLocPortId";E={$_.LocPortId}},
                            @{L="ComputerName";E={$NetAdapter.PSComputerName}},
                            @{L="ifAlias";E={$NetAdapter.ifAlias}},
                            @{L="ifDesc";E={$NetAdapter.ifDesc}},
                            @{L="MacAddress";E={$NetAdapter.MacAddress}},
                            @{L="IntentType";E={$NetAdapter.IntentType}},
                            @{L="vLAN";E={$NetAdapter.vLAN}}
                    }
                }
            }
        #$SwPortToHostMap | sort SwHostName,SwLocPortId,ComputerName,ifAlias | ft
        # Add to HTML report output sections
        if($SwPortToHostMap){
            AddTo-HtmlReport -Title "Interface-to-Node Map" `
                -Data $SwPortToHostMap `
                -Description "" `
                -Footnotes 'Highlighted in red or yellow if out of spec. <p><a href="https://infohub.delltechnologies.com/en-us/t/switch-configurations-roce-iwarp-mellanox-and-intel-e810-cards-reference-guide/" target="_blank">Ref: Switch Configurations – RoCE/iWarp Reference Guide</a></p>' `
                -IncludeTitle -IncludeDescription -IncludeFootnotes
        }

    #endregion

    #region Show Running Configuration



        $ShowRunningConfigs = $STSLOC | ForEach-Object { Get-OS10RunningConfigSections -Path $_ }
        #$ShowRunningConfigs | ft

        #Check for OS version as we support OS9,10 and Sonic
        #IF vLAN 711-714 these are the storage interfaces see show vLAN status

        #class-map type network-qos(?:\s+\S+)?

        $SwitchHostnames = $ShowRunningConfigs | select Filename, SwHostName | sort Filename -Unique

        #dcbx enable
        $dcbxenable = @()
        $dcbxenableOut = ""
        $dcbxenable = $ShowRunningConfigs | ?{$_.lines -imatch "dcbx enable"}
        IF ($dcbxenable){
            $dcbxenableOut = $dcbxenable | select Filename, SwHostName,@{L="dcbx enable";E={"Found"}}
        }Else{
            $dcbxenableOut = $dcbxenable | select Filename, SwHostName,@{L="dcbx enable";E={"RREEDDMissing"}}
        }
        #$dcbxenableout | ft
        # Add to HTML report output sections
        if($dcbxenableout){
            AddTo-HtmlReport -Title "dcbx enable" `
                -Data $dcbxenableout `
                -Description "" `
                -Footnotes 'Highlighted in red or yellow if out of spec. <p><a href="https://infohub.delltechnologies.com/en-us/t/switch-configurations-roce-iwarp-mellanox-and-intel-e810-cards-reference-guide/" target="_blank">Ref: Switch Configurations – RoCE/iWarp Reference Guide</a></p>' `
                -IncludeTitle -IncludeDescription -IncludeFootnotes
        }

        #class-map type queuing Q0
        $classmaptypequeuingQ = @()
        $classmaptypequeuingQ0 = @()
        $classmaptypequeuingQ57 = @()
        $classmaptypequeuingQOut = @()
        $classmaptypequeuingQ = $ShowRunningConfigs | ?{$_.lines -imatch "class-map type queuing Q"}
        IF($classmaptypequeuingQ){
    
            #Check for #class-map type queuing Q0
            IF($classmaptypequeuingQ | ?{$_.lines -imatch 'class-map type queuing Q0'}){
                $classmaptypequeuingQ0 += $classmaptypequeuingQ | ?{$_.lines -imatch 'class-map type queuing Q0'} | select Filename, SwHostName,
                    @{L="class-map type queuing Q0";E={IF($_.Lines -imatch 'class-map type queuing Q0'){"Found"}Else{"RREEDDMissing"}}},
                    @{L="match queue 0";            E={IF($_.lines -imatch 'match queue 0'){"Found"}Else{"RREEDDMissing"}}}
            }Else{ 
                $classmaptypequeuingQ0 += $classmaptypequeuingQ | sort FileName -Unique | select Filename, SwHostName,@{L="class-map type queuing Q0";E={"RREEDDMissing"}},@{L="match queue 0";E={"RREEDDMissing"}}
            }
            #$classmaptypequeuingQ0 | ft
            # Add to HTML report output sections
            if($classmaptypequeuingQ0){
                AddTo-HtmlReport -Title "class-map type queuing Q0" `
                    -Data $classmaptypequeuingQ0 `
                    -Description "" `
                    -Footnotes 'Highlighted in red or yellow if out of spec. <p><a href="https://infohub.delltechnologies.com/en-us/t/switch-configurations-roce-iwarp-mellanox-and-intel-e810-cards-reference-guide/" target="_blank">Ref: Switch Configurations – RoCE/iWarp Reference Guide</a></p>' `
                    -IncludeTitle -IncludeDescription -IncludeFootnotes
            }

            #Check for #class-map type queuing Q5/7
            ####Add 5 or 7
            IF($classmaptypequeuingQ | ?{$_.lines -imatch 'class-map type queuing Q(5|7)'}){
                $classmaptypequeuingQ57 += $classmaptypequeuingQ | ?{$_.lines -imatch 'class-map type queuing Q5'} | select Filename, SwHostName,
                    @{L="class-map type queuing Q7";E={IF($_.Lines -imatch 'class-map type queuing Q5'){"Found"}Else{"RREEDDMissing"}}},
                    @{L="match queue 5";            E={IF($_.lines -imatch 'match queue 5'){"Found"}Else{"RREEDDMissing"}}}
                
                $classmaptypequeuingQ57 += $classmaptypequeuingQ | ?{$_.lines -imatch 'class-map type queuing Q7'} | select Filename, SwHostName,
                    @{L="class-map type queuing Q7";E={IF($_.Lines -imatch 'class-map type queuing Q7'){"Found"}Else{"RREEDDMissing"}}},
                    @{L="match queue 7";            E={IF($_.lines -imatch 'match queue 7'){"Found"}Else{"RREEDDMissing"}}}
            }Else{
                #Write-Host "     FAIL: Missing both class-map type queuing Q5 and Q7. Assume Q7" -ForegroundColor red
                $classmaptypequeuingQ57 += $classmaptypequeuingQ | sort FileName -Unique | select Filename, SwHostName,@{L="class-map type queuing Q5/7";E={"RREEDDMissing"}},@{L="match queue 5/7";E={"RREEDDMissing"}}
            }
            #$classmaptypequeuingQ7 | ft
            # Add to HTML report output sections
            if($classmaptypequeuingQ57){
                AddTo-HtmlReport -Title "class-map type queuing Q5/7" `
                    -Data $classmaptypequeuingQ57 `
                    -Description "" `
                    -Footnotes 'Highlighted in red or yellow if out of spec. <p><a href="https://infohub.delltechnologies.com/en-us/t/switch-configurations-roce-iwarp-mellanox-and-intel-e810-cards-reference-guide/" target="_blank">Ref: Switch Configurations – RoCE/iWarp Reference Guide</a></p>' `
                    -IncludeTitle -IncludeDescription -IncludeFootnotes
            }

        }Else{
            $classmaptypequeuingQOut += $ShowRunningConfigs | sort FileName -Unique | select Filename, SwHostName,
                @{L="class-map type queuing Q0";E={"RREEDDMissing"}},
                @{L="match queue 0";            E={"RREEDDMissing"}},
                @{L="class-map type queuing Q7";E={"RREEDDMissing"}},
                @{L="match queue 7";            E={"RREEDDMissing"}} 
        }

        #$classmaptypequeuingQOut | ft
        # Add to HTML report output sections
        if($classmaptypequeuingQOut){
            AddTo-HtmlReport -Title "class-map type queuing Q" `
                -Data $classmaptypequeuingQOut `
                -Description "" `
                -Footnotes 'Highlighted in red or yellow if out of spec. <p><a href="https://infohub.delltechnologies.com/en-us/t/switch-configurations-roce-iwarp-mellanox-and-intel-e810-cards-reference-guide/" target="_blank">Ref: Switch Configurations – RoCE/iWarp Reference Guide</a></p>' `
                -IncludeTitle -IncludeDescription -IncludeFootnotes
        }

        #class-map type network-qos Management
        $matchqosgroup0out = @()
        $matchqosgroup3out = @()
        $matchqosgroup7out = @()
        $classmaptypenetworkqosManagement = @()
        $classmaptypenetworkqosManagement = $ShowRunningConfigs | ?{$_.lines -imatch "class-map type network-qos"}
        IF($classmaptypenetworkqosManagement){
            #match qos-group 0
                $matchqosgroup0 = $classmaptypenetworkqosManagement | ?{$_.lines -imatch "match qos-group 0"}
                IF($matchqosgroup0){
                    $matchqosgroup0out = $matchqosgroup0 | sort FileName -Unique | select Filename, SwHostName,@{L=$matchqosgroup0.lines[1];E={IF(($_.lines -imatch "match qos-group 0")){"Found"}Else{"RREEDDMissing"}}},@{L="match qos-group 0";E={IF($_.lines -imatch "match qos-group 0"){"Found"}Else{"RREEDDMissing"}}}
                }Else{
                    $matchqosgroup0out = $classmaptypenetworkqosManagement | sort FileName -Unique | select Filename, SwHostName,@{L="class-map type network-qos";E={"RREEDDMissing"}},@{L="match queue 0";E={"RREEDDMissing"}}
                }
                #$matchqosgroup0out | ft
                # Add to HTML report output sections
                if($matchqosgroup0out){
                    AddTo-HtmlReport -Title "class-map type network-qos group 0" `
                        -Data $matchqosgroup0out `
                        -Description "" `
                        -Footnotes 'Highlighted in red or yellow if out of spec. <p><a href="https://infohub.delltechnologies.com/en-us/t/switch-configurations-roce-iwarp-mellanox-and-intel-e810-cards-reference-guide/" target="_blank">Ref: Switch Configurations – RoCE/iWarp Reference Guide</a></p>' `
                        -IncludeTitle -IncludeDescription -IncludeFootnotes
                }
            #match qos-group 3
                $matchqosgroup3 = $classmaptypenetworkqosManagement | ?{$_.lines -imatch "match qos-group 3"}
                IF($matchqosgroup3){
                    $matchqosgroup3out = $matchqosgroup3 | sort FileName -Unique | select Filename, SwHostName,@{L=$matchqosgroup3.lines[1];E={IF(($_.lines -imatch "class-map type network-qos") -and ($_.lines -imatch "match qos-group 3")){"Found"}Else{"RREEDDMissing"}}},@{L="match qos-group 3";E={IF($_.lines -imatch "match qos-group 3"){"Found"}Else{"RREEDDMissing"}}}
                }Else{
                    $matchqosgroup3out = $classmaptypenetworkqosManagement | sort FileName -Unique | select Filename, SwHostName,@{L="class-map type network-qos";E={"RREEDDMissing"}},@{L="match queue 3";E={"RREEDDMissing"}}
                }          
                #$matchqosgroup3out | ft
                # Add to HTML report output sections
                if($matchqosgroup3out){
                    AddTo-HtmlReport -Title "class-map type network-qos group 3" `
                        -Data $matchqosgroup3out `
                        -Description "" `
                        -Footnotes 'Highlighted in red or yellow if out of spec. <p><a href="https://infohub.delltechnologies.com/en-us/t/switch-configurations-roce-iwarp-mellanox-and-intel-e810-cards-reference-guide/" target="_blank">Ref: Switch Configurations – RoCE/iWarp Reference Guide</a></p>' `
                        -IncludeTitle -IncludeDescription -IncludeFootnotes
                }
            #match qos-group 7
                $matchqosgroup7 = $classmaptypenetworkqosManagement | ?{$_.lines -imatch "match qos-group 7"}
                IF($matchqosgroup7){
                    $matchqosgroup7out = $matchqosgroup7 | sort FileName -Unique | select Filename, SwHostName,@{L=$matchqosgroup7.lines[1];E={IF(($_.lines -imatch "class-map type network-qos") -and ($_.lines -imatch "match qos-group 7")){"Found"}Else{"RREEDDMissing"}}},@{L="match qos-group 7";E={IF($_.lines -imatch "match qos-group 7"){"Found"}Else{"RREEDDMissing"}}}
                }Else{
                    $matchqosgroup7out = $classmaptypenetworkqosManagement | sort FileName -Unique | select Filename, SwHostName,@{L="class-map type network-qos";E={"RREEDDMissing"}},@{L="match queue 7";E={"RREEDDMissing"}}
                }
                #$matchqosgroup7out | ft
                # Add to HTML report output sections
                if($matchqosgroup7out){
                    AddTo-HtmlReport -Title "class-map type network-qos group 7" `
                        -Data $matchqosgroup7out `
                        -Description "" `
                        -Footnotes 'Highlighted in red or yellow if out of spec. <p><a href="https://infohub.delltechnologies.com/en-us/t/switch-configurations-roce-iwarp-mellanox-and-intel-e810-cards-reference-guide/" target="_blank">Ref: Switch Configurations – RoCE/iWarp Reference Guide</a></p>' `
                        -IncludeTitle -IncludeDescription -IncludeFootnotes
                }
        }Else{
            #no class-map type network-qos
            $classmaptypenetworkqos = $ShowRunningConfigs | sort FileName -Unique | select Filename, SwHostName,@{L="class-map type network-qos";E={"RREEDDMissing"}}
            #$classmaptypenetworkqos  | ft
            # Add to HTML report output sections
            if($classmaptypenetworkqos){
                AddTo-HtmlReport -Title "class-map type network-qos" `
                    -Data $classmaptypenetworkqos `
                    -Description "" `
                    -Footnotes 'Highlighted in red or yellow if out of spec. <p><a href="https://infohub.delltechnologies.com/en-us/t/switch-configurations-roce-iwarp-mellanox-and-intel-e810-cards-reference-guide/" target="_blank">Ref: Switch Configurations – RoCE/iWarp Reference Guide</a></p>' `
                    -IncludeTitle -IncludeDescription -IncludeFootnotes
            }
        }

        #trust dot1p-map trust_map
        $trustdot1pmaptrustmap = @()
        $trustdot1pmaptrustmapOut = @()
        $trustdot1pmaptrustmap = $ShowRunningConfigs | ?{$_.Header -imatch "trust dot1p-map trust_map"}
        IF($trustdot1pmaptrustmap){
            $trustdot1pmaptrustmapOut = $trustdot1pmaptrustmap | sort FileName -Unique | select Filename, SwHostName,
                @{L="trust dot1p-map trust_map";E={IF($_.lines -imatch "trust dot1p-map trust_map"){"Found"}Else{"RREEDDMissing"}}},
                @{L="qos-group 0 dot1p 0-2,4-6";E={IF($_.lines -imatch "qos-group 0 dot1p 0-2,4-6"){"Found"}Else{"RREEDDMissing"}}},
                @{L="qos-group 3 dot1p 3";E={IF($_.lines -imatch "qos-group 3 dot1p 3"){"Found"}Else{"RREEDDMissing"}}},
                @{L="qos-group 7 dot1p 7";E={IF($_.lines -imatch "qos-group 7 dot1p 7"){"Found"}Else{"RREEDDMissing"}}}
        }Else{
            #no trust dot1p-map trust_map
            $trustdot1pmaptrustmapOut = $ShowRunningConfigs | sort FileName -Unique | select Filename, SwHostName,@{L="trust dot1p-map trust_map";E={"RREEDDMissing"}}
        }
        #$trustdot1pmaptrustmapOut  | ft
        # Add to HTML report output sections
        if($trustdot1pmaptrustmapOut){
            AddTo-HtmlReport -Title "trust dot1p-map trust_map" `
                -Data $trustdot1pmaptrustmapOut `
                -Description "" `
                -Footnotes 'Highlighted in red or yellow if out of spec. <p><a href="https://infohub.delltechnologies.com/en-us/t/switch-configurations-roce-iwarp-mellanox-and-intel-e810-cards-reference-guide/" target="_blank">Ref: Switch Configurations – RoCE/iWarp Reference Guide</a></p>' `
                -IncludeTitle -IncludeDescription -IncludeFootnotes
        }

        #qos-map traffic-class queue-map
        $qosmaptrafficclassqueuemap = @()
        $qosmaptrafficclassqueuemap = $ShowRunningConfigs | ?{$_.header -imatch "qos-map traffic-class queue-map"}
        IF($qosmaptrafficclassqueuemap){
            $qosmaptrafficclassqueuemapOut = $qosmaptrafficclassqueuemap | sort FileName -Unique | select Filename, SwHostName,
                @{L="qos-map traffic-class queue-map";E={IF($_.lines -imatch "qos-map traffic-class queue-map"){"Found"}Else{"RREEDDMissing"}}},
                @{L="queue 0 qos-group 0-2,4-6";E={IF($_.lines -imatch "queue 0 qos-group 0-2,4-6"){"Found"}Else{"RREEDDMissing"}}},
                @{L="queue 3 qos-group 3";E={IF($_.lines -imatch " queue 3 qos-group 3"){"Found"}Else{"RREEDDMissing"}}},
                @{L="queue 5/7 qos-group 5/7";E={
                    IF($_.lines -imatch "queue 5 qos-group 5"){"Found"}
                    ElseIf($_.lines -imatch "queue 7 qos-group 7"){"Found"}
                    Else{"RREEDDMissing"}}}
        }Else{
            #no qos-map traffic-class queue-map
            $qosmaptrafficclassqueuemapOut = $ShowRunningConfigs | sort FileName -Unique | select Filename, SwHostName,@{L="qos-map traffic-class queue-map";E={"RREEDDMissing"}}
        }
        #$qosmaptrafficclassqueuemapOut  | ft
        # Add to HTML report output sections
        if($qosmaptrafficclassqueuemapOut){
            AddTo-HtmlReport -Title "qos-map traffic-class queue-map" `
                -Data $qosmaptrafficclassqueuemapOut `
                -Description "" `
                -Footnotes 'Highlighted in red or yellow if out of spec. <p><a href="https://infohub.delltechnologies.com/en-us/t/switch-configurations-roce-iwarp-mellanox-and-intel-e810-cards-reference-guide/" target="_blank">Ref: Switch Configurations – RoCE/iWarp Reference Guide</a></p>' `
                -IncludeTitle -IncludeDescription -IncludeFootnotes
        }

        #policy-map type application policy-iscsi
        $policymaptypeapplicationpolicyiscsi = @()
        $policymaptypeapplicationpolicyiscsi = $ShowRunningConfigs | ?{$_.header -imatch "policy-map type application policy-iscsi"}
        IF($policymaptypeapplicationpolicyiscsi){
            $policymaptypeapplicationpolicyiscsiOut = $policymaptypeapplicationpolicyiscsi | sort FileName -Unique | select Filename, SwHostName,
                @{L="policy-map type application policy-iscsi";E={IF($_.lines -imatch "policy-map type application policy-iscsi"){"Found"}Else{"RREEDDMissing"}}}
        }Else{
            #no policy-map type application policy-iscsi
            $policymaptypeapplicationpolicyiscsiOut = $ShowRunningConfigs | sort FileName -Unique | select Filename, SwHostName,@{L="policy-map type application policy-iscsi";E={"RREEDDMissing"}}
        }
        #$policymaptypeapplicationpolicyiscsiOut  | ft
        # Add to HTML report output sections
        if($policymaptypeapplicationpolicyiscsiOut){
            AddTo-HtmlReport -Title "policy-map type application policy-iscsi" `
                -Data $policymaptypeapplicationpolicyiscsiOut `
                -Description "" `
                -Footnotes 'Highlighted in red or yellow if out of spec. <p><a href="https://infohub.delltechnologies.com/en-us/t/switch-configurations-roce-iwarp-mellanox-and-intel-e810-cards-reference-guide/" target="_blank">Ref: Switch Configurations – RoCE/iWarp Reference Guide</a></p>' `
                -IncludeTitle -IncludeDescription -IncludeFootnotes
        }

        #policy-map type queuing ets-policy
        $policymaptypequeuingetspolicy = @()
        $policymaptypequeuingetspolicy = $ShowRunningConfigs | ?{$_.header -imatch "policy-map type queuing ets-policy"}
        IF($policymaptypequeuingetspolicy){
            $policymaptypequeuingetspolicyOut = $policymaptypequeuingetspolicy | sort FileName -Unique | select Filename, SwHostName,
                @{L="policy-map type queuing ets-policy";E={IF($_.lines -imatch "policy-map type queuing ets-policy"){"Found"}Else{"RREEDDMissing"}}}
        }Else{
            #no policy-map type queuing ets-policy
            $policymaptypequeuingetspolicyOut = $ShowRunningConfigs | sort FileName -Unique | select Filename, SwHostName,@{L="policy-map type application policy-iscsi";E={"RREEDDMissing"}}
        }
        #$policymaptypeapplicationpolicyiscsiOut  | ft
        # Add to HTML report output sections
        if($policymaptypequeuingetspolicyOut){
            AddTo-HtmlReport -Title "policy-map type queuing ets-policy" `
                -Data $policymaptypeapplicationpolicyiscsiOut `
                -Description "" `
                -Footnotes 'Highlighted in red or yellow if out of spec. <p><a href="https://infohub.delltechnologies.com/en-us/t/switch-configurations-roce-iwarp-mellanox-and-intel-e810-cards-reference-guide/" target="_blank">Ref: Switch Configurations – RoCE/iWarp Reference Guide</a></p>' `
                -IncludeTitle -IncludeDescription -IncludeFootnotes
        }

        #class Q0,3,5,7
        $classQ0357 = @()
        $classQ0Out = ""
        $classQ3Out = ""
        $classQ5Out = ""
        $classQ7Out = ""
        $classQ57Out = ""
        $classQ0357Out = ""
        $classQ0357 = $ShowRunningConfigs | ?{$_.lines -imatch 'class Q(0|3|5|7)'}
        IF($classQ0357){
            IF($classQ0357 | ?{$_.Header -imatch "Q0"}){
                $classQ0Out = $classQ0357 | ?{$_.Header -imatch "Q0"}| select Filename, SwHostName,
                    @{L="class Q0";E={IF($_.lines -imatch "class Q0"){"Found"}Else{"RREEDDMissing"}}},
                    @{L="bandwidth percent 48 or 49";E={IF($_.lines -imatch "bandwidth percent (48|49)"){"Found"}Else{"RREEDDMissing"}}}
            }Else{
                $classQ0Out = $classQ0357 | select Filename, SwHostName,
                    @{L="class Q0";E={IF($_.lines -imatch "class Q0"){"Found"}Else{"RREEDDMissing"}}},
                    @{L="bandwidth percent 48 or 49";E={IF($_.lines -imatch "bandwidth percent (48|49)"){"Found"}Else{"RREEDDMissing"}}}
            }
            If($classQ0357 | ?{$_.Header -imatch "Q3"}){
                $classQ3Out = $classQ0357 | ?{$_.Header -imatch "Q3"} | select Filename, SwHostName,
                    @{L="class Q3";E={IF($_.lines -imatch "class Q3"){"Found"}Else{"RREEDDMissing"}}},
                    @{L="bandwidth percent 50";E={IF($_.lines -imatch "bandwidth percent 50"){"Found"}Else{"RREEDDMissing"}}}
            }Else{
                $classQ3Out = $classQ0357 | select Filename, SwHostName,
                    @{L="class Q3";E={IF($_.lines -imatch "class Q3"){"Found"}Else{"RREEDDMissing"}}},
                    @{L="bandwidth percent 50";E={IF($_.lines -imatch "bandwidth percent 50"){"Found"}Else{"RREEDDMissing"}}}
            }
            #Case 1 no Q5 and no Q7
            IF($classQ0357 | ?{$_.Header -inotmatch "Q7" -and $_.Header -inotmatch "Q5"}){
                $classQ57Out = $classQ0357 | select Filename, SwHostName,
                    @{L="class Q5/7";E={IF($_.lines -imatch "class Q5/7"){"Found"}Else{"RREEDDMissing Q5 and Q7"}}},
                    @{L="bandwidth percent 1 or 2";E={IF($_.lines -imatch "bandwidth percent (1|2)"){"Found"}Else{"RREEDDMissing"}}}
            }
            #Case 2 we have a Q5 and no Q7
            IF($classQ0357 | ?{$_.Header -imatch "Q5" -and $_.Header -inotmatch "Q7"}){
                If($classQ0357 | ?{$_.Header -imatch "Q5"}){
                    $classQ5Out = $classQ0357 | ?{$_.Header -imatch "Q5"} | select Filename, SwHostName,
                        @{L="class Q7";E={IF($_.lines -imatch "class Q5"){"Found"}Else{"RREEDDMissing"}}},
                        @{L="bandwidth percent 1 or 2";E={IF($_.lines -imatch "bandwidth percent (1|2)"){"Found"}Else{"RREEDDMissing"}}}
                    }
            }
            #Case 3 we have a Q7 and no Q5
            IF($classQ0357 | ?{$_.Header -imatch "Q7" -and $_.Header -inotmatch "Q5"}){
                If($classQ0357 | ?{$_.Header -imatch "Q7"}){
                    $classQ7Out = $classQ0357 | ?{$_.Header -imatch "Q7"} | select Filename, SwHostName,
                        @{L="class Q7";E={IF($_.lines -imatch "class Q7"){"Found"}Else{"RREEDDMissing"}}},
                        @{L="bandwidth percent 1 or 2";E={IF($_.lines -imatch "bandwidth percent (1|2)"){"Found"}Else{"RREEDDMissing"}}}
                }
            }
        }Else{
            #no policy-map type queuing ets-policy
            $classQ0357Out = $ShowRunningConfigs | sort FileName -Unique | select Filename, SwHostName,@{L="class Q0|3|5|7";E={"RREEDDMissing"}}
        }
        #$classQ0Out | ft
        #$classQ3Out | ft
        #$classQ7Out | ft
        #$classQ037Out | ft
        # Add to HTML report output sections
        if($classQ0Out){
            AddTo-HtmlReport -Title "policy-map type queuing ets-policy class Q0" `
                -Data $classQ0Out `
                -Description "" `
                -Footnotes 'Highlighted in red or yellow if out of spec. <p><a href="https://infohub.delltechnologies.com/en-us/t/switch-configurations-roce-iwarp-mellanox-and-intel-e810-cards-reference-guide/" target="_blank">Ref: Switch Configurations – RoCE/iWarp Reference Guide</a></p>' `
                -IncludeTitle -IncludeDescription -IncludeFootnotes
        }
        if($classQ3Out){
            AddTo-HtmlReport -Title "policy-map type queuing ets-policy class Q3" `
                -Data $classQ3Out `
                -Description "" `
                -Footnotes 'Highlighted in red or yellow if out of spec. <p><a href="https://infohub.delltechnologies.com/en-us/t/switch-configurations-roce-iwarp-mellanox-and-intel-e810-cards-reference-guide/" target="_blank">Ref: Switch Configurations – RoCE/iWarp Reference Guide</a></p>' `
                -IncludeTitle -IncludeDescription -IncludeFootnotes
        }
        if($classQ57Out){
            AddTo-HtmlReport -Title "policy-map type queuing ets-policy class Q5/7" `
                -Data $classQ57Out `
                -Description "" `
                -Footnotes 'Highlighted in red or yellow if out of spec. <p><a href="https://infohub.delltechnologies.com/en-us/t/switch-configurations-roce-iwarp-mellanox-and-intel-e810-cards-reference-guide/" target="_blank">Ref: Switch Configurations – RoCE/iWarp Reference Guide</a></p>' `
                -IncludeTitle -IncludeDescription -IncludeFootnotes
        }
        if($classQ5Out){
            AddTo-HtmlReport -Title "policy-map type queuing ets-policy class Q5" `
                -Data $classQ5Out `
                -Description "" `
                -Footnotes 'Highlighted in red or yellow if out of spec. <p><a href="https://infohub.delltechnologies.com/en-us/t/switch-configurations-roce-iwarp-mellanox-and-intel-e810-cards-reference-guide/" target="_blank">Ref: Switch Configurations – RoCE/iWarp Reference Guide</a></p>' `
                -IncludeTitle -IncludeDescription -IncludeFootnotes
        }
        if($classQ7Out){
            AddTo-HtmlReport -Title "policy-map type queuing ets-policy class Q7" `
                -Data $classQ7Out `
                -Description "" `
                -Footnotes 'Highlighted in red or yellow if out of spec. <p><a href="https://infohub.delltechnologies.com/en-us/t/switch-configurations-roce-iwarp-mellanox-and-intel-e810-cards-reference-guide/" target="_blank">Ref: Switch Configurations – RoCE/iWarp Reference Guide</a></p>' `
                -IncludeTitle -IncludeDescription -IncludeFootnotes
        }
        if($classQ0357Out){
            AddTo-HtmlReport -Title "policy-map type queuing ets-policy class Q" `
                -Data $classQ0357Out `
                -Description "" `
                -Footnotes 'Highlighted in red or yellow if out of spec. <p><a href="https://infohub.delltechnologies.com/en-us/t/switch-configurations-roce-iwarp-mellanox-and-intel-e810-cards-reference-guide/" target="_blank">Ref: Switch Configurations – RoCE/iWarp Reference Guide</a></p>' `
                -IncludeTitle -IncludeDescription -IncludeFootnotes
        }
        #policy-map type network-qos pfc-policy
        $policymaptypenetworkqospfcpolicy = @()
        $policymaptypenetworkqospfcpolicy = $ShowRunningConfigs | ?{$_.header -imatch "policy-map type network-qos pfc-policy"}
        IF($policymaptypenetworkqospfcpolicy){
            $policymaptypenetworkqospfcpolicyOut = $policymaptypenetworkqospfcpolicy | sort FileName -Unique | select Filename, SwHostName,
                @{L="policy-map type network-qos pfc-policy";E={IF($_.lines -imatch "policy-map type network-qos pfc-policy"){"Found"}Else{"RREEDDMissing"}}}
        }Else{
            #no policy-map type queuing ets-policy
            $policymaptypenetworkqospfcpolicyOut = $ShowRunningConfigs | sort FileName -Unique | select Filename, SwHostName,@{L="policy-map type network-qos pfc-policy";E={"RREEDDMissing"}}
        }
        #$policymaptypenetworkqospfcpolicyOut  | ft
        # Add to HTML report output sections
        if($policymaptypenetworkqospfcpolicyOut){
            AddTo-HtmlReport -Title "Policy-map type network-qos pfc-policy" `
                -Data $policymaptypenetworkqospfcpolicyOut `
                -Description "" `
                -Footnotes 'Highlighted in red or yellow if out of spec. <p><a href="https://infohub.delltechnologies.com/en-us/t/switch-configurations-roce-iwarp-mellanox-and-intel-e810-cards-reference-guide/" target="_blank">Ref: Switch Configurations – RoCE/iWarp Reference Guide</a></p>' `
                -IncludeTitle -IncludeDescription -IncludeFootnotes
        }


        #pfc-cos 3
        $pfccos3 = @()
        $pfccos3 = $ShowRunningConfigs | ?{$_.Lines -imatch 'pfc-cos 3'}
        IF($pfccos3){
            $pfccos3Out = $pfccos3 | sort FileName -Unique | select Filename, SwHostName,
                @{L=($pfccos3.Lines[1]);E={IF($_.lines -imatch "class "){"Found"}Else{"RREEDDMissing"}}},
                @{L="pause";    E={IF($_.lines -imatch "pause"){"Found"}Else{"RREEDDMissing"}}},
                @{L="pfc-cos 3";E={IF($_.lines -imatch "pfc-cos 3"){"Found"}Else{"RREEDDMissing"}}}
        }Else{
            #no policy-map type queuing ets-policy
            $pfccos3Out = $ShowRunningConfigs | sort FileName -Unique | select Filename, SwHostName,@{L="class Q0|3|7";E={"RREEDDMissing"}}
        }
        #$pfccos3Out | ft
        # Add to HTML report output sections
        if($pfccos3Out){
            AddTo-HtmlReport -Title "policy-map type network-qos pfc-policy pfc-cos 3" `
                -Data $pfccos3Out `
                -Description "" `
                -Footnotes 'Highlighted in red or yellow if out of spec. <p><a href="https://infohub.delltechnologies.com/en-us/t/switch-configurations-roce-iwarp-mellanox-and-intel-e810-cards-reference-guide/" target="_blank">Ref: Switch Configurations – RoCE/iWarp Reference Guide</a></p>' `
                -IncludeTitle -IncludeDescription -IncludeFootnotes
        }

        #system qos
        $systemqos = @()
        $systemqos = $ShowRunningConfigs | ?{$_.Header -imatch 'system qos'}
        IF($systemqos){
            $systemqosOut = $systemqos | sort FileName -Unique | select Filename, SwHostName,
                @{L="system qos";E={IF($_.lines -imatch "system qos"){"Found"}Else{"RREEDDMissing"}}},
                @{L=($systemqos.Lines[2]);    E={IF($_.lines -imatch "trust-map dot1p"){"Found"}Else{"RREEDDMissing"}}}
        }Else{
            #no policy-map type queuing ets-policy
            $systemqosOut = $ShowRunningConfigs | sort FileName -Unique | select Filename, SwHostName,@{L="system qos";E={"RREEDDMissing"}}
        }
        #$systemqosOut | ft
        # Add to HTML report output sections
        if($systemqosOut){
            AddTo-HtmlReport -Title "System QOS" `
                -Data $systemqosOut `
                -Description "" `
                -Footnotes 'Highlighted in red or yellow if out of spec. <p><a href="https://infohub.delltechnologies.com/en-us/t/switch-configurations-roce-iwarp-mellanox-and-intel-e810-cards-reference-guide/" target="_blank">Ref: Switch Configurations – RoCE/iWarp Reference Guide</a></p>' `
                -IncludeTitle -IncludeDescription -IncludeFootnotes
        }

        <#interface vlan711,712,713,714
        $interfacevlan7xx = @()
        $interfacevlan7xxOut = @()
        $interfacevlan711Out = @()
        $interfacevlan712Out = @()
        $interfacevlan713Out = @()
        $interfacevlan714Out = @()
        $interfacevlan7xx = $ShowRunningConfigs | ?{$_.Lines -imatch 'interface vlan(711|712|713|714)'}
        IF($interfacevlan7xx){
            IF($interfacevlan7xx| ?{$_.header -imatch "711"}){
                $interfacevlan711Out += $interfacevlan7xx| ?{$_.header -imatch "711"} | sort FileName -Unique | select Filename, SwHostName,
                @{L="interface vlan711";E={IF($_.lines -imatch "interface vlan711"){"Found"}Else{"RREEDDMissing"}}},
                @{L="MTU9216";    E={IF($_.lines -imatch "9216"){"Found"}Else{"RREEDDMissing"}}},
                @{L="no shutdown";    E={IF($_.lines -imatch "no shutdown"){"Found"}Else{"RREEDDMissing"}}}
            }
            IF($interfacevlan7xx| ?{$_.header -imatch "712"}){
                $interfacevlan712Out += $interfacevlan7xx| ?{$_.header -imatch "712"} | sort FileName -Unique | select Filename, SwHostName,
                @{L="interface vlan712";E={IF($_.lines -imatch "interface vlan712"){"Found"}Else{"RREEDDMissing"}}},
                @{L="MTU9216";    E={IF($_.lines -imatch "9216"){"Found"}Else{"RREEDDMissing"}}},
                @{L="no shutdown";    E={IF($_.lines -imatch "no shutdown"){"Found"}Else{"RREEDDMissing"}}}
            }
            IF($interfacevlan7xx| ?{$_.header -imatch "713"}){
                $interfacevlan713Out +=$interfacevlan7xx| ?{$_.header -imatch "713"}| sort FileName -Unique | select Filename, SwHostName,
                @{L="interface vlan713";E={IF($_.lines -imatch "interface vlan713"){"Found"}Else{"RREEDDMissing"}}},
                @{L="MTU9216";    E={IF($_.lines -imatch "9216"){"Found"}Else{"RREEDDMissing"}}},
                @{L="no shutdown";    E={IF($_.lines -imatch "no shutdown"){"Found"}Else{"RREEDDMissing"}}}
            }
            IF($interfacevlan7xx| ?{$_.header -imatch "714"}){
                $interfacevlan714Out += $interfacevlan7xx| ?{$_.header -imatch "714"} | sort FileName -Unique | select Filename, SwHostName,
                @{L="interface vlan714";E={IF($_.lines -imatch "interface vlan714"){"Found"}Else{"RREEDDMissing"}}},
                @{L="MTU9216";    E={IF($_.lines -imatch "9216"){"Found"}Else{"RREEDDMissing"}}},
                @{L="no shutdown";    E={IF($_.lines -imatch "no shutdown"){"Found"}Else{"RREEDDMissing"}}}
            }

        }Else{
            #no policy-map type queuing ets-policy
            $interfacevlan7xxOut = $ShowRunningConfigs | sort FileName -Unique | select Filename, SwHostName,@{L="interface vlan(711|712|713|714)";E={"RREEDDMissing"}}
        }
        $interfacevlan7xxOut | ft
        $interfacevlan711Out | ft
        $interfacevlan712Out | ft
        $interfacevlan713Out | ft
        $interfacevlan714Out | ft
        #>

    #endregion


    #-------------------------------------------------------------
    # Convert to Comparison Table for easy review
    #-------------------------------------------------------------
    function Convert-ToSwitchComparisonTable {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory)]
            [array]$Interfaces
        )

        # Determine columns (sorted for predictable order)
        $columns = $Interfaces | Sort-Object SwHostName, Header

        # Dynamically detect all properties except metadata
        $exclude = 'FileName','SwHostName','Header'
        $props = $Interfaces |
            ForEach-Object { $_.PSObject.Properties.Name } |
            Where-Object { $_ -notin $exclude } |
            Sort-Object -Unique

        # Build the comparison table
        $table = @{}

        foreach ($p in $props) {
            $row = [ordered]@{ 'ShouldBe' = $p }

            foreach ($iface in $columns) {
                $colName = "$($iface.SwHostName):$($iface.Header)"
                $row[$colName] = $iface.PSObject.Properties[$p].Value
            }

            $table[$p] = [PSCustomObject]$row
        }

        # Output as table object (ready for Format-Table or Export-Csv)
        return $table.Values | Select-Object *
    }

    #-------------------------------------------------------------
    # Used to out-grid is width of output is too large 
    #-------------------------------------------------------------
    function Show-WideTable {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory, ValueFromPipeline)]
            [object]$InputObject,

            [string]$Title = "Data View",

            [int]$MaxWidth = 4000
        )

        begin {
            $data = @()
        }

        process {
            $data += $InputObject
        }

        end {
            if (-not $data) {
                Write-Warning "No data provided."
                return
            }

            # If in PowerShell ISE
            if ($psISE) {
                Write-Host "Detected PowerShell ISE — opening '$Title' in Out-GridView..."
                try {
                    $data | Out-GridView -Title $Title
                } catch {
                    Write-Warning "Unable to open Out-GridView. $_"
                }
                return
            }

            # Otherwise, in Console or Windows Terminal
            try {
                # Measure table width
                $width = ($data | Format-Table -AutoSize | Out-String).Split("`n") |
                    ForEach-Object { $_.Length } |
                    Measure-Object -Maximum |
                    Select-Object -ExpandProperty Maximum

                $width = [Math]::Min($width, $MaxWidth)

                # Adjust console width
                $rawUI = $Host.UI.RawUI
                $rawUI.BufferSize = New-Object Management.Automation.Host.Size($width, $rawUI.BufferSize.Height)
                $rawUI.WindowSize = New-Object Management.Automation.Host.Size([Math]::Min($width, 300), $rawUI.WindowSize.Height)

                Write-Host "`n=== $Title ==="
                Write-Host "Console width set to $width characters (scroll horizontally to view).`n"
            } catch {
                Write-Warning "Unable to resize console (likely a restricted host). Try Out-GridView instead."
            }

            # Display formatted table
            $data | Format-Table -AutoSize
        }
    }


    #-------------------------------------------------------------
    # Find port types from NetworkATC intents
    #-------------------------------------------------------------
    #region Port Configurations

        $MgmtUsedInterfaces=@()
        $StorageUsedInterfaces=@()
        ForEach ($port in $SwPortToHostMap){
            IF($port.IntentType -eq "Mgmt"){
                ForEach($Interface in $ShowRunningConfigs){
                  $MgmtUsedInterfaces+=$Interface | ?{ ($_.SwHostName -eq $port.SwHostName) -and $_.header -eq "interface "+$port.SwLocPortId} | select *,@{L="IntentType";E={$Port.IntentType}},@{L="vLAN";E={$Port.vLAN}}
                }
            }
            IF($port.IntentType -eq "Storage"){
                ForEach($Interface in $ShowRunningConfigs){
                  $StorageUsedInterfaces+=$Interface | ?{ ($_.SwHostName -eq $port.SwHostName) -and $_.header -eq "interface "+$port.SwLocPortId} | select *,@{L="IntentType";E={$Port.IntentType}},@{L="vLAN";E={$Port.vLAN}}
                }
            }
        }

    #-------------------------------------------------------------
    # Find Matches in array
    #-------------------------------------------------------------
    function Get-LineValue {
        param ($lines, $pattern)
        $result = $lines | ForEach-Object { $_.Trim() } |
            Where-Object { $_ -imatch $pattern } |
            Select-Object -First 1
        if ($null -ne $result -and $result -ne '') {
            return $result
        } else {
            return ''
        }
    }

    #-------------------------------------------------------------
    # CHECK MISSING
    #-------------------------------------------------------------
    function Set-MissingNoteProperties {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory, ValueFromPipeline)]
            [pscustomobject[]]$InputObject
        )

        process {
            foreach ($obj in $InputObject) {
                # Only check NoteProperties (not methods or type members)
                $props = $obj.PSObject.Properties |
                         Where-Object { $_.MemberType -eq 'NoteProperty' }

                foreach ($prop in $props) {
                    $value = $prop.Value

                    if ($null -eq $value -or ($value -is [string] -and $value.Trim() -eq '')) {
                        # Set missing or blank value
                        $obj.PSObject.Properties[$prop.Name].Value = 'RREEDDMissing'
                    }
                }

                # Output the updated object
                $obj
            }
        }
    }


    #-------------------------------------------------------------
    # Gather vLANs
    #-------------------------------------------------------------
    $Storagevlans = $StorageUsedInterfaces.vlan
    $Storagevlans = $Storagevlans | sort -Unique
    $MgmtvLans = $MgmtUsedInterfaces.vlan
    $MgmtvLans = $MgmtvLans | sort -Unique


    #-------------------------------------------------------------
    # STORAGE INTERFACES
    #-------------------------------------------------------------
    $StorageUsedInterfacesOut = @()
    $StorageUsedInterfaceInfo = ""
    foreach ($StorageUsedInterface in $StorageUsedInterfaces) {
        $StorageUsedInterfaceInfo = [pscustomobject]@{
            FileName                                           = $StorageUsedInterface.FileName
            SwHostName                                         = $StorageUsedInterface.SwHostName
            Header                                             = $StorageUsedInterface.Header
            PortType                                           = 'Storage'
            vLAN                                               = $StorageUsedInterface.vLAN
            Description                                        = (Get-LineValue $StorageUsedInterface.Lines 'description' | select @{L="Description";E={$_ -replace "description",""}}).description
            'no shutdown'                                      = Get-LineValue $StorageUsedInterface.Lines 'no shutdown'
            'switchport mode trunk'                            = Get-LineValue $StorageUsedInterface.Lines 'switchport mode trunk'
            'switchport trunk allowed vlan'                    = (Get-LineValue $StorageUsedInterface.Lines 'switchport trunk allowed vlan' | select @{L='switchport trunk allowed vlan';E={
                                                                       #Check for missing storage vlan
                                                                        if($_ -inotmatch [regex]::Escape($StorageUsedInterface.vLAN.ToString())){"RREEDD"+$_}
                                                                       #We should NOT have Mgmt vLANs in storage trunk ex: switchport trunk allowed vlan 201,711-712,1701-1702,3939 where 201=Mgmt
                                                                        Elseif($_ -imatch ($MgmtvLans -join '|')){"RREEDD"+$_}
                                                                       #Matches storage vlan 
                                                                        Elseif($_ -imatch [regex]::Escape($StorageUsedInterface.vLAN.ToString())){$_}
                                                                 }}).'switchport trunk allowed vlan'
            'MTU9216'                                          = Get-LineValue $StorageUsedInterface.Lines '9216'
            'flowcontrol receive off'                          = Get-LineValue $StorageUsedInterface.Lines 'flowcontrol receive off'
            'flowcontrol transmit off'                         = Get-LineValue $StorageUsedInterface.Lines 'flowcontrol transmit off'
            'spanning-tree bpduguard enable'                   = Get-LineValue $StorageUsedInterface.Lines 'spanning-tree bpduguard enable'
            'spanning-tree port type edge'                     = Get-LineValue $StorageUsedInterface.Lines 'spanning-tree port type edge'
            'priority-flow-control mode on'                    = Get-LineValue $StorageUsedInterface.Lines 'priority-flow-control mode on'
            'service-policy input type network-qos pfc-policy' = Get-LineValue $StorageUsedInterface.Lines 'service-policy input type network-qos pfc-policy'
            'service-policy output type queuing ets-policy'    = Get-LineValue $StorageUsedInterface.Lines 'service-policy output type queuing ets-policy'
            'ets mode on'                                      = Get-LineValue $StorageUsedInterface.Lines 'ets mode on'
            'qos-map traffic-class queue-map'                  = Get-LineValue $StorageUsedInterface.Lines 'qos-map traffic-class queue-map'
        }
        $StorageUsedInterfacesOut += Set-MissingNoteProperties $StorageUsedInterfaceInfo
    
    }
    #Write-Host "Storage Interfaces"
    #$StorageUsedInterfacesOut | ft * -AutoSize -Wrap
    $StorageUsedInterfacesEasyOut = Convert-ToSwitchComparisonTable -Interfaces $StorageUsedInterfacesOut | sort ShouldBe
    #$StorageUsedInterfacesEasyOut | Show-WideTable -Title "Storage Switch Port Comparison"
    # Add to HTML report output sections
    AddTo-HtmlReport -Title "Storage Interfaces" `
        -Data $StorageUsedInterfacesEasyOut `
        -Description "" `
        -Footnotes 'Highlighted in red or yellow if out of spec. <p><a href="https://infohub.delltechnologies.com/en-us/t/switch-configurations-roce-iwarp-mellanox-and-intel-e810-cards-reference-guide/" target="_blank">Ref: Switch Configurations – RoCE/iWarp Reference Guide</a></p>' `
        -IncludeTitle -IncludeDescription -IncludeFootnotes

    #-------------------------------------------------------------
    # MGMT INTERFACES
    #-------------------------------------------------------------
    $MgmtUsedInterfacesOut = @()

    foreach ($MgmtUsedInterface in $MgmtUsedInterfaces) {
        $MgmtUsedInterfaceInfo = [pscustomobject]@{
            FileName                          = $MgmtUsedInterface.FileName
            SwHostName                        = $MgmtUsedInterface.SwHostName
            Header                            = $MgmtUsedInterface.Header
            PortType                          = 'Mgmt'
            vLAN                              = $MgmtUsedInterface.vLAN
            Description                       = (Get-LineValue $MgmtUsedInterface.Lines 'description' | select @{L="Description";E={$_ -replace "description",""}}).description
            'no shutdown'                     = Get-LineValue $MgmtUsedInterface.Lines 'no shutdown'
            'switchport mode trunk'           = Get-LineValue $MgmtUsedInterface.Lines 'switchport mode trunk'
            'switchport trunk allowed vlan'   = (Get-LineValue $MgmtUsedInterface.Lines 'switchport trunk allowed vlan' | select @{L='switchport trunk allowed vlan';E={
                                                    #Check for missing Mgmt vlan
                                                    if($_ -inotmatch [regex]::Escape($MgmtUsedInterface.vLAN.ToString())){"RREEDD"+$_}
                                                    #We should NOT have storage vLANs in storage trunk ex: switchport trunk allowed vlan 201,711-712,1701-1702,3939 where 201=Mgmt
                                                    Elseif($_ -imatch ($Storagevlans -join '|')){"RREEDD"+$_}
                                                    #Matches Mgmt vlan 
                                                    Elseif($_ -imatch [regex]::Escape($MgmtUsedInterface.vLAN.ToString())){$_}
                                                }}).'switchport trunk allowed vlan'
            'MTU9216'                         = Get-LineValue $MgmtUsedInterface.Lines '9216'
            'flowcontrol receive on'          = (Get-LineValue $MgmtUsedInterface.Lines 'flowcontrol receive' | select @{L="flowcontrol receive on";E={
                                                    If($_ -imatch " on"){$_}Else{"RREEDD"+$_}}}).'flowcontrol receive on'
            'flowcontrol transmit off'        = Get-LineValue $MgmtUsedInterface.Lines 'flowcontrol transmit off'
            'spanning-tree bpduguard enable'  = Get-LineValue $MgmtUsedInterface.Lines 'spanning-tree bpduguard enable'
            'spanning-tree port type edge'    = Get-LineValue $MgmtUsedInterface.Lines 'spanning-tree port type edge'
        }

        $MgmtUsedInterfacesOut += Set-MissingNoteProperties $MgmtUsedInterfaceInfo
    
    }
    #Write-Host "Mgmt Interfaces"
    #$MgmtUsedInterfacesOut | ft
    $MgmtUsedInterfacesEasyOut = Convert-ToSwitchComparisonTable -Interfaces $MgmtUsedInterfacesOut | sort ShouldBe
    #$MgmtUsedInterfacesEasyOut | Show-WideTable -Title "Mgmt Switch Port Comparison"
    # Add to HTML report output sections
    AddTo-HtmlReport -Title "Mgmt Interfaces" `
        -Data $MgmtUsedInterfacesEasyOut `
        -Description "" `
        -Footnotes 'Highlighted in red or yellow if out of spec. <p><a href="https://infohub.delltechnologies.com/en-us/t/switch-configurations-roce-iwarp-mellanox-and-intel-e810-cards-reference-guide/" target="_blank">Ref: Switch Configurations – RoCE/iWarp Reference Guide</a></p>' `
        -IncludeTitle -IncludeDescription -IncludeFootnotes

    #region VLTi
    #-------------------------------------------------------------
    # VLTi INTERFACES
    #-------------------------------------------------------------
        #Find the ports that are assigned to VLTi
        #----------------------------------- show vlt all -------------------
         function Get-showvltall {
                [CmdletBinding()]
                [OutputType([Object[]])]
                param(
                    [Parameter(Mandatory, Position=0)]
                    [string]$Path
                )
                #$path="C:\Users\Jim_Gandy\Downloads\tk5tor17-01a-show-tech-20251021-134546.txt"
                if (-not (Test-Path -LiteralPath $Path)) {
                    throw "File not found: $Path"
                }

                # Read whole file
                $text = Get-Content -LiteralPath $Path -Raw

                # Strip ANSI/VT100 and CRs
                $text = [regex]::Replace($text, "\x1B\[[0-?]*[ -/]*[@-~]", "")
                $text = $text -replace "`r",""
                $SwitchHostname=""
                $SwitchHostname = (((($text | ?{$_ -imatch "hostname"}) -split "hostname ")[-1] -split "`n")[0]).trim()

                # Grab the "show lldp neighbors" section delimited by dashed headers
                $pattern = '(?is)^\s*-{3,}\s*show\s+vlt\s+all\s*-{3,}\s*\n(.*?)(?=^\s*-{3,}\s*show\s+\S.*?-{3,}\s*$|\Z)'
                $m = [regex]::Match($text, $pattern, 'IgnoreCase, Multiline, Singleline')
                if (-not $m.Success) {
                    throw "Could not locate the 'show vlt all' section. Check header format in the log."
                }

                $section = $m.Groups[1].Value.Trim()
                $lines   = $section -split "`n"
            # Split into lines and trim
            $lines = $section -split '\n' | ForEach-Object { $_.Trim() } | Where-Object { $_ }

            # Create output object
            $result = [ordered]@{}
            $peerTable = @()

            # Switch context after reaching peer table header
            $inPeerTable = $false
            foreach ($line in $lines) {
                if ($line -match '^VLT Peer Unit ID') {
                    $inPeerTable = $true
                    continue
                }
                if ($inPeerTable) {
                    if ($line -match '^-{5,}') { continue }  # skip separator
                    if ($line -match '^\d') {
                        $parts = ($line -split '\s{2,}') | ForEach-Object { $_.Trim() }
                        $peerTable += [pscustomobject]@{
                            PeerUnitID        = $parts[0]
                            SystemMacAddress  = $parts[1]
                            Status            = $parts[2]
                            IPAddress         = $parts[3]
                            Version           = $parts[4]
                        }
                    }
                }
                else {
                    if ($line -match '^(?<key>[^:]+):\s*(?<value>.+)$') {
                        $key = ($matches['key'].Trim() -replace '\s+', '_')
                        $result[$key] = $matches['value'].Trim()
                
                    }
                }
            }
            # Build the output dynamically from whatever keys were found
            $obj = [pscustomobject]@{}
            foreach ($kvp in $result.GetEnumerator()) {
                Add-Member -InputObject $obj -NotePropertyName $kvp.Key -NotePropertyValue $kvp.Value
            }

            # Add hostname and peer list
            Add-Member -InputObject $obj -NotePropertyName 'Hostname' -NotePropertyValue $SwitchHostname
            Add-Member -InputObject $obj -NotePropertyName 'Peers' -NotePropertyValue $peerTable

            return $obj

        }


        #----------------------------------- show port-channel summary -------------------
        function Get-showportchannelsummary {
            [CmdletBinding()]
            [OutputType([Object[]])]
            param(
                [Parameter(Mandatory, Position=0)]
                [string]$Path
            )

            if (-not (Test-Path -LiteralPath $Path)) {
                throw "File not found: $Path"
            }

            # Read and clean
            $text = Get-Content -LiteralPath $Path -Raw
            $text = [regex]::Replace($text, "\x1B\[[0-?]*[ -/]*[@-~]", "")
            $text = $text -replace "`r", ""

            # Extract hostname
            $SwitchHostname = (((($text | Select-String -Pattern 'hostname\s+\S+') -split 'hostname ')[-1] -split "`n")[0]).Trim()

            # Extract section
            $pattern = '(?is)^\s*-{3,}\s*show\s+port-channel\s+summary\s*-{3,}\s*\n(.*?)(?=^\s*-{3,}\s*show\s+\S.*?-{3,}\s*$|\Z)'
            $m = [regex]::Match($text, $pattern, 'IgnoreCase, Multiline, Singleline')
            if (-not $m.Success) {
                throw "Could not locate the 'show port-channel summary' section. Check header format in the log."
            }

            $section = $m.Groups[1].Value.Trim()
            $lines = $section -split "`n" | ForEach-Object { $_.TrimEnd() } | Where-Object { $_ }

            # Find header
            $headerLine = $lines | Where-Object { $_ -match '^\s*Group\s+Port-Channel' }
            if (-not $headerLine) { throw "Header not found." }

            # Extract headers cleanly
            $headers = $headerLine -split '\s{2,}' | ForEach-Object { $_.Trim() }
            $lastDashIndex = ($lines | Select-String '^---' | Select-Object -Last 1).LineNumber
            $dataLines = $lines[$lastDashIndex..($lines.Count - 1)] | Where-Object { $_ -notmatch '^---' }

            # Parse data lines based on column spacing
            $objects = foreach ($line in $dataLines) {
                if (-not $line.Trim()) { continue }
                $parts = $line -split '\s{2,}', $headers.Count
                # pad if short
                while ($parts.Count -lt $headers.Count) { $parts += '' }

                $obj = [ordered]@{ Hostname = $SwitchHostname }
                for ($i = 0; $i -lt $headers.Count; $i++) {
                    $obj[$headers[$i]] = $parts[$i].Trim()
                }
                [pscustomobject]$obj
            }

            return $objects
        }

            $showvltall=@()
            $showvltall += $STSLOC | ForEach-Object { Get-showvltall -Path $_ }

    

        #$showportchannelsummary = Get-showportchannelsummary -path "C:\Users\Jim_Gandy\Downloads\tk5tor17-01a-show-tech-20251021-134546.txt"
        $showportchannelsummary = $STSLOC | ForEach-Object { Get-showportchannelsummary -Path $_ }
        $VLTiPorts = ($showportchannelsummary | ?{$_.'Group Port-Channel' -imatch ($showvltall | select port-channel* | GM | ?{$_.MemberType -eq "NoteProperty"}).Name} | select @{L="VLTi Ports";E={$_.'Member Ports' -split " "-replace '\([A-Z]+\)', '' }}).'VLTi Ports'| sort -Unique
        $VLTiUsedInterfaces = @()
        ForEach($VLTiPort in $VLTiPorts){
            ForEach($Interface in $ShowRunningConfigs){
                $VLTiUsedInterfaces += $Interface | ?{$_.Header -imatch "interface ethernet"+$VLTiPort}
            }
        }
                <# mtu 9216
                 flowcontrol receive off
                 flowcontrol transmit off
                 priority-flow-control mode on
                 service-policy input type network-qos pfc-policy
                 service-policy output type queuing ets-policy
                 ets mode on
                 qos-map traffic-class queue-map
                 no shutdown
                 no switchport#>
            
                $VLTiUsedInterfacesOut = @()
                $VLTiUsedInterfaceInfo = ""
                foreach ($VLTiUsedInterface in $VLTiUsedInterfaces) {
                    $VLTiUsedInterfaceInfo = [pscustomobject]@{
                        FileName                                           = $VLTiUsedInterface.FileName
                        SwHostName                                         = $VLTiUsedInterface.SwHostName
                        Header                                             = $VLTiUsedInterface.Header
                        PortType                                           = 'VLTi'
                        Description                                        = (Get-LineValue $VLTiUsedInterface.Lines 'description' | select @{L="Description";E={$_ -replace "description",""}}).description
                        'no shutdown'                                      = Get-LineValue $VLTiUsedInterface.Lines 'no shutdown'
                        'no switchport'                                    = Get-LineValue $VLTiUsedInterface.Lines 'no switchport'
                        'MTU9216'                                          = Get-LineValue $VLTiUsedInterface.Lines '9216'
                        'flowcontrol receive off'                          = Get-LineValue $VLTiUsedInterface.Lines 'flowcontrol receive off'
                        'flowcontrol transmit off'                         = Get-LineValue $VLTiUsedInterface.Lines 'flowcontrol transmit off'
                        'priority-flow-control mode on'                    = Get-LineValue $VLTiUsedInterface.Lines 'priority-flow-control mode on'
                        'service-policy input type network-qos pfc-policy' = Get-LineValue $VLTiUsedInterface.Lines 'service-policy input type network-qos pfc-policy'
                        'service-policy output type queuing ets-policy'    = Get-LineValue $VLTiUsedInterface.Lines 'service-policy output type queuing ets-policy'
                        'ets mode on'                                      = Get-LineValue $VLTiUsedInterface.Lines 'ets mode on'
                        'qos-map traffic-class queue-map'                  = Get-LineValue $VLTiUsedInterface.Lines 'qos-map traffic-class queue-map'
                    }
                    $VLTiUsedInterfacesOut += Set-MissingNoteProperties $VLTiUsedInterfaceInfo
    
                }

    $VLTiUsedInterfacesEasyOut = Convert-ToSwitchComparisonTable -Interfaces $VLTiUsedInterfacesOut  | sort ShouldBe
    #$VLTiUsedInterfacesEasyOut | Show-WideTable -Title "VLTi Switch Port Comparison"
    # Add to HTML report output sections
    AddTo-HtmlReport -Title "VLTi Interfaces" `
        -Data $VLTiUsedInterfacesEasyOut `
        -Description "" `
        -Footnotes 'Highlighted in red or yellow if out of spec. <p><a href="https://infohub.delltechnologies.com/en-us/t/switch-configurations-roce-iwarp-mellanox-and-intel-e810-cards-reference-guide/" target="_blank">Ref: Switch Configurations – RoCE/iWarp Reference Guide</a></p>' `
        -IncludeTitle -IncludeDescription -IncludeFootnotes

    #endregion VLTi

    #-------------------------------------------------------------
    # vLAN INTERFACES
    #-------------------------------------------------------------
    $StoragevLANUsedInterfaces = @()
    ForEach ($StoragevLAN in $Storagevlans){
        ForEach($Interface in $ShowRunningConfigs){
                $StoragevLANUsedInterfaces += $Interface | ?{$_.Header -imatch "interface vlan"+$StoragevLAN}
        }
    }
    $StoragevLANUsedInterfacesOut = @()
    $StoragevLANUsedInterfacesInfo = ""
    foreach ($StoragevLANUsedInterface in $StoragevLANUsedInterfaces) {
        $StoragevLANUsedInterfacesInfo = [pscustomobject]@{
            FileName                                           = $StoragevLANUsedInterface.FileName
            SwHostName                                         = $StoragevLANUsedInterface.SwHostName
            Header                                             = $StoragevLANUsedInterface.Header
            PortType                                           = 'vLAN'
            Description                                        = (Get-LineValue $StoragevLANUsedInterface.Lines 'description' | select @{L="Description";E={$_ -replace "description",""}}).description
            'no shutdown'                                      = Get-LineValue $StoragevLANUsedInterface.Lines 'no shutdown'
            'MTU9216'                                          = Get-LineValue $StoragevLANUsedInterface.Lines '9216'
        }
        $StoragevLANUsedInterfacesOut += Set-MissingNoteProperties $StoragevLANUsedInterfacesInfo
    }

    $StoragevLANUsedInterfacesEasyOut = Convert-ToSwitchComparisonTable -Interfaces $StoragevLANUsedInterfacesOut | sort ShouldBe
    #$StoragevLANUsedInterfacesEasyOut | Show-WideTable -Title "vLAN Switch Port Comparison"
    # Add to HTML report output sections
    AddTo-HtmlReport -Title "Storage vLAN Interfaces" `
        -Data $StoragevLANUsedInterfacesEasyOut `
        -Description "" `
        -Footnotes 'Highlighted in red or yellow if out of spec. <p><a href="https://infohub.delltechnologies.com/en-us/t/switch-configurations-roce-iwarp-mellanox-and-intel-e810-cards-reference-guide/" target="_blank">Ref: Switch Configurations – RoCE/iWarp Reference Guide</a></p>' `
        -IncludeTitle -IncludeDescription -IncludeFootnotes



    #endregion

    # Save report
    Save-HtmlReport
}
}