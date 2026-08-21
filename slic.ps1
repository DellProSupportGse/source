<#
.SYNOPSIS
    SLIC - Switch Log InspeCtor
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
    2026/08/20:v1.37 - JG - Updated links to the validated switch configs.
    2026/08/20:v1.36 - JG - Modernized SLIC HTML report to match CluChk 2.0 visual framework.
    See GitHub pull requests for history

#>
Function Invoke-SLIC {

# Console output intentionally uses ASCII-only status markers so the script renders
# consistently in Windows PowerShell 5.1 and PowerShell ISE. HTML-only symbols are
# generated with entities/character codes and do not depend on the .ps1 file encoding.

Function EndScript{  
    break
}
$Ver="v1.37"
$ToolName = @"
$Ver
  ___ _    ___ ___ 
 / __| |  |_ _/ __|
 \__ \ |__ | | (__ 
 |___/____|___\___|
 Switch Log InspeCtor
            By: Jim Gandy
"@
Clear-Host
Write-Host $ToolName
Write-Host ""
Write-Host "[!] SLIC Compatibility Notice:" -ForegroundColor Yellow
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
    $STSLOC = Get-FileName "$env:USERPROFILE\Documents\SRs" "Please Select Show Tech-Support File(s)." "Logs (*.txt,*.log)| *.TXT;*.log"
    If(!($STSLOC)){
        Write-Host "No logs provided. Exiting script..."
        EndScript
    }Else{
        Write-Host "[+] Switch Logs:" $STSLOC -ForegroundColor Green
    }

    $SDDCPath = Read-Host "Please provide the path to the extracted SDDC"
    
    If(!(Test-Path $SDDCPath -ErrorAction SilentlyContinue)){
        Write-Host "SDDC path not found. Exiting script..." -ForegroundColor Red
        EndScript
    }Else{
        Write-Host "[+] SDDC Path:" $SDDCPath -ForegroundColor Green
    }

    #region === HTML Report System ===

    function New-HtmlReport {
        param (
            [string]$Title = "SLIC: Switch Log InspeCtor",
            [string]$Version = "",
            [string]$RunDate = (Get-Date),
            [string]$OutputPath = "$env:TEMP\SLIC_Report.html"
        )

        $script:HtmlReportSections = @()
        $script:HtmlReportTitle = $Title
        $script:HtmlReportVersion = $Version
        $script:HtmlReportRunDate = $RunDate
        $script:HtmlReportPath = $OutputPath
        $script:HtmlSectionCounter = 0
    }

    function AddTo-HtmlReport {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory)]
            [AllowEmptyCollection()]
            [array]$Data,
            [string]$Title = "Report Section",
            [string]$Description = "",
            [string]$Footnotes = "",
            [switch]$IncludeTitle,
            [switch]$IncludeDescription,
            [switch]$IncludeFootnotes
        )

        $script:HtmlSectionCounter++
        $slug = (($Title -replace '[^A-Za-z0-9]+','-').Trim('-')).ToLowerInvariant()
        if ([string]::IsNullOrWhiteSpace($slug)) { $slug = "section" }
        $sectionId = "slic-$slug-$($script:HtmlSectionCounter)"

        $html = ""
        if ($IncludeDescription -and $Description) {
            $html += "<div class='section-description'>$Description</div>`n"
        }
        if ($null -ne $Data -and $Data.Count -gt 0) {
            $html += ($Data | ConvertTo-Html -Fragment)
        }
        if ($IncludeFootnotes -and $Footnotes) {
            $html += "<div class='section-footnotes'>$Footnotes</div>`n"
        }

        $hasError = ($html -match 'RREEDD')
        $hasWarning = (($html -match 'YYEELLLLOOWW') -or ($html -match 'warning-banner'))

        # Match the softer CluChk 2.0 report palette.
        $html = $html `
            -replace '<td>RREEDD', '<td style="color:#a4262c;background-color:#fde7e9;font-weight:600">' `
            -replace '<td>YYEELLLLOOWW', '<td style="color:#5c4b00;background-color:#fff4ce;font-weight:600">'

        $status = if ($hasError) { 'error' } elseif ($hasWarning) { 'warning' } else { 'healthy' }

        $script:HtmlReportSections += [pscustomobject]@{
            Id = $sectionId
            Label = $Title
            Html = $html
            Status = $status
        }
    }

    function Save-HtmlReport {
        if (-not $script:HtmlReportSections) {
            Write-Warning "No sections added to report."
            return
        }

        $errorCount = @($script:HtmlReportSections | Where-Object Status -eq 'error').Count
        $warningCount = @($script:HtmlReportSections | Where-Object Status -eq 'warning').Count
        $healthyCount = @($script:HtmlReportSections | Where-Object Status -eq 'healthy').Count
        $totalCount = @($script:HtmlReportSections).Count

        $summaryBody = $script:HtmlReportSections | ForEach-Object {
            $warn = if ($_.Status -eq 'warning') { 1 } else { 0 }
            $err  = if ($_.Status -eq 'error') { 1 } else { 0 }
            $label = [System.Net.WebUtility]::HtmlEncode($_.Label)
            $warnClass = if ($warn -gt 0) { " class='summary-warning'" } else { '' }
            $errClass  = if ($err -gt 0)  { " class='summary-error'" } else { '' }
            "<tr><td><a href='#$($_.Id)'>$label</a></td><td$warnClass>$warn</td><td$errClass>$err</td></tr>"
        }
        $summaryHtml = "<table><thead><tr><th>Name</th><th>Warnings</th><th>Errors</th></tr></thead><tbody>$($summaryBody -join '')</tbody></table>"

        $overviewHtml = @"
<h1>$($script:HtmlReportTitle)</h1>
<div class='report-meta'>
  <span><b>Version:</b> $($script:HtmlReportVersion)</span>
  <span><b>Run Date:</b> $($script:HtmlReportRunDate)</span>
</div>
<div class='warning-banner'>&#9888; <b>SLIC Compatibility Notice:</b> This tool currently supports <b>Azure Local</b> and <b>Windows Server S2D</b> clusters only.</div>
<div class='overview-cards'>
  <div class='overview-card'><span class='overview-number'>$totalCount</span><span class='overview-label'>Sections</span></div>
  <div class='overview-card warning-card'><span class='overview-number'>$warningCount</span><span class='overview-label'>Warnings</span></div>
  <div class='overview-card error-card'><span class='overview-number'>$errorCount</span><span class='overview-label'>Errors</span></div>
  <div class='overview-card healthy-card'><span class='overview-number'>$healthyCount</span><span class='overview-label'>Healthy</span></div>
</div>
<h2>Results Summary</h2>
$summaryHtml
"@

        $allSections = @([pscustomobject]@{ Id='report-overview'; Label='Report Overview'; Html=$overviewHtml; Status='overview' }) + $script:HtmlReportSections

        $navItems = $allSections | ForEach-Object {
            $statusClass = if ($_.Status -in @('error','warning')) { " $($_.Status)" } else { '' }
            "<li><button class='tab-link$statusClass' data-status='$($_.Status)' data-target='$($_.Id)' onclick='showTab(this.dataset.target)'>$([System.Net.WebUtility]::HtmlEncode($_.Label))</button></li>"
        }

        $sectionHtml = $allSections | ForEach-Object {
            $active = if ($_.Id -eq 'report-overview') { ' active' } else { '' }
            if ($_.Id -eq 'report-overview') {
                "<section id='$($_.Id)' class='tab-panel$active'>$($_.Html)</section>"
            } else {
                "<section id='$($_.Id)' class='tab-panel$active'><h2>$([System.Net.WebUtility]::HtmlEncode($_.Label))</h2>$($_.Html)</section>"
            }
        }

        $SlicLogoLightData = 'data:image/png;base64,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'
        $SlicLogoDarkData = 'data:image/png;base64,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'

        $finalHtml = @"
<!DOCTYPE html>
<html lang='en'>
<head>
<meta charset='UTF-8'>
<meta name='viewport' content='width=device-width, initial-scale=1.0'>
<title>$($script:HtmlReportTitle)</title>
<style>
:root {
  --page-bg:#f5f5f5; --panel-bg:#ffffff; --text:#1f1f1f; --muted-text:#5f6368;
  --sidebar:#ffffff; --sidebar-hover:#f3f8fc; --sidebar-active:#e7f3fb;
  --accent:#0078d4; --accent-hover:#106ebe; --heading:#1f1f1f;
  --table-head:#e8f3fb; --table-head-text:#1f1f1f; --table-head-border:#c9e2f3;
  --row-alt:#fafafa; --row-hover:#eef7fd; --border:#e3e3e3; --link:#0078d4;
  --warning-bg:#fff4ce; --warning-text:#5c4b00; --warning-nav:#fff4ce;
  --error-bg:#fde7e9; --error-nav:#fde7e9; --success-bg:#dff6dd; --success-text:#0b6a0b;
  --input-bg:#ffffff; --input-border:#d6d6d6; --nav-text:#333333;
  --cell-border:#ededed; --scroll-thumb:#c8c8c8; --selected-text:#005a9e;
  --panel-shadow:0 1px 4px rgba(0,0,0,.06);
}
body.dark-theme {
  --page-bg:#1b1d1f; --panel-bg:#25282b; --text:#f2f2f2; --muted-text:#b8bec5;
  --sidebar:#202326; --sidebar-hover:#2b3640; --sidebar-active:#173b57;
  --accent:#4aa8e8; --accent-hover:#67b7ee; --heading:#f2f2f2;
  --table-head:#20394a; --table-head-text:#f4f8fb; --table-head-border:#35566b;
  --row-alt:#2a2d30; --row-hover:#303b44; --border:#404449; --link:#6cb8eb;
  --warning-bg:#514617; --warning-text:#ffe69a; --warning-nav:#4b4118;
  --error-bg:#542b2f; --error-nav:#4d292d; --success-bg:#173d25; --success-text:#8fe39f;
  --input-bg:#2b2e31; --input-border:#4b5055; --nav-text:#e9ecef;
  --cell-border:#3c4044; --scroll-thumb:#60656a; --selected-text:#8ed0f5;
  --panel-shadow:0 1px 5px rgba(0,0,0,.28);
}
*{box-sizing:border-box} html,body{height:100%;margin:0;padding:0}
body{display:flex;font-family:'Segoe UI',Tahoma,Geneva,Verdana,sans-serif;background:var(--page-bg);color:var(--text)}
.sidebar{width:300px;min-width:300px;background:var(--sidebar);height:100vh;overflow-y:auto;position:sticky;top:0;display:flex;flex-direction:column;border-right:1px solid var(--border)}
.brand{padding:.8rem 1rem;background:#fff;text-align:center;border-bottom:1px solid var(--border);transition:background .15s ease}
body.dark-theme .brand{background:var(--sidebar)}
.brand img{display:block;width:250px;max-width:100%;height:auto;margin:0 auto}
.sidebar-controls{padding:.65rem .8rem 0;display:grid;gap:.45rem}
.sidebar-control{display:block;width:100%;padding:.55rem .7rem;background:var(--accent);color:#fff;border:1px solid var(--accent);border-radius:5px;text-decoration:none;text-align:center;cursor:pointer;font-size:.86rem;font-weight:600}
.sidebar-control:hover{background:var(--accent-hover);border-color:var(--accent-hover)}
.issue-link-wrap{text-align:center;padding:.65rem .8rem .1rem}.issue-link{color:var(--link);font-size:.84rem;text-decoration:none}.issue-link:hover{text-decoration:underline}
#tabSearch{margin:.65rem .8rem;padding:.55rem .7rem;background:var(--input-bg);color:var(--text);border:1px solid var(--input-border);border-radius:5px;font-size:.9rem}
.sidebar ul{list-style:none;margin:0;padding:0;flex:1}.sidebar li{border-bottom:1px solid var(--border)}.sidebar li.hidden{display:none}
.sidebar button.tab-link{width:100%;padding:.75rem 1rem;background:transparent;color:var(--nav-text);border:none;text-align:left;cursor:pointer;font-size:.88rem;display:block;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.sidebar button.tab-link:hover{background:var(--sidebar-hover)}
.sidebar button.tab-link.active{background:var(--sidebar-active);color:var(--selected-text);border-left:4px solid var(--accent);border-right:5px solid var(--accent);padding-left:calc(1rem - 4px);font-weight:600}
.sidebar button.tab-link.error{background:var(--error-nav)!important;color:#a4262c;border-left:4px solid #d13438;padding-left:calc(1rem - 4px)}
body.dark-theme .sidebar button.tab-link.error{color:#ffb4ba}
.sidebar button.tab-link.warning{background:var(--warning-nav)!important;color:#7a5d00;border-left:4px solid #d9a300;padding-left:calc(1rem - 4px)}
body.dark-theme .sidebar button.tab-link.warning{color:#ffe69a}
.main-content{flex:1;padding:1.5rem;overflow-y:auto;height:100vh;background:var(--page-bg);scroll-behavior:smooth}
.tab-panel{display:none;background:var(--panel-bg);border:1px solid var(--border);border-radius:8px;padding:1.2rem 1.35rem;margin-bottom:1rem;box-shadow:var(--panel-shadow)}
.tab-panel.active{display:block}body.show-all .tab-panel{display:block}
h1{font-size:1.65rem;color:var(--heading);margin:.1rem 0 .45rem;border-bottom:3px solid var(--accent);padding-bottom:.45rem}
h2{font-size:1.15rem;color:var(--heading);margin:.15rem 0 1rem}
.report-meta{display:flex;gap:1.2rem;flex-wrap:wrap;color:var(--muted-text);margin-bottom:.9rem;font-size:.9rem}
.warning-banner{background:var(--warning-bg);color:var(--warning-text);border:1px solid #d9c866;border-radius:6px;padding:.7rem .85rem;margin:.7rem 0 1rem}
.overview-cards{display:grid;grid-template-columns:repeat(4,minmax(115px,1fr));gap:.75rem;margin:1rem 0 1.2rem}
.overview-card{background:var(--panel-bg);border:1px solid var(--border);border-left:4px solid var(--accent);border-radius:6px;padding:.8rem}.overview-number{display:block;font-size:1.55rem;font-weight:700}.overview-label{font-size:.8rem;color:var(--muted-text)}
.warning-card{border-left-color:#d9a300;background:var(--warning-bg);color:var(--warning-text)}
.error-card{border-left-color:#d13438;background:var(--error-bg);color:var(--text)}
.healthy-card{border-left-color:#107c10;background:var(--success-bg);color:var(--success-text)}
.warning-card .overview-label{color:var(--warning-text)}
.error-card .overview-label{color:inherit}
.healthy-card .overview-label{color:var(--success-text)}
.summary-warning{background:var(--warning-bg)!important;color:var(--warning-text)!important;font-weight:700}
.summary-error{background:var(--error-bg)!important;color:#a4262c!important;font-weight:700}
body.dark-theme .summary-error{color:#ffb3b8!important}
table{border-collapse:separate;border-spacing:0;width:100%;margin:.5rem 0 1rem;font-size:.88rem;border:1px solid var(--border);border-radius:6px;overflow:hidden}
th,td{padding:.55rem .65rem;border-bottom:1px solid var(--cell-border);border-right:1px solid var(--cell-border);text-align:left;vertical-align:top}th:last-child,td:last-child{border-right:none}tr:last-child td{border-bottom:none}
th{background:var(--table-head);color:var(--table-head-text);border-color:var(--table-head-border);font-weight:600;cursor:pointer;user-select:none;position:relative;padding-right:1.5rem}
th .sort-indicator{position:absolute;right:.45rem;opacity:.55;font-size:.72rem}th.sorted{color:var(--accent)}th.sorted .sort-indicator{opacity:1}
tbody tr:nth-child(even) td{background:var(--row-alt)}tbody tr:hover td{background:var(--row-hover)}
a{color:var(--link)}.section-description,.section-footnotes{margin:.5rem 0;color:var(--muted-text)}.section-footnotes{font-size:.84rem}
::-webkit-scrollbar{width:8px;height:8px}::-webkit-scrollbar-thumb{background:var(--scroll-thumb);border-radius:4px}
@media(max-width:768px){body{flex-direction:column}.sidebar{width:100%;min-width:100%;height:auto;max-height:45vh;position:relative}.main-content{height:auto;overflow-y:visible}.overview-cards{grid-template-columns:repeat(2,1fr)}}
</style>
<script>
var showAllTables=false;
function applyTheme(theme){var dark=theme==='dark';document.body.classList.toggle('dark-theme',dark);var btn=document.getElementById('themeToggle');if(btn){btn.textContent=dark?'Light Mode':'Dark Mode';btn.setAttribute('aria-pressed',dark?'true':'false');}var logo=document.getElementById('slicLogo');if(logo){var lightSrc=logo.getAttribute('data-light-src');var darkSrc=logo.getAttribute('data-dark-src');logo.setAttribute('src',dark?darkSrc:lightSrc);}}
function toggleTheme(){var next=document.body.classList.contains('dark-theme')?'light':'dark';applyTheme(next);try{localStorage.setItem('slic-theme',next)}catch(e){}}
function loadSavedTheme(){var saved=null;try{saved=localStorage.getItem('slic-theme')}catch(e){}if(saved!=='dark'&&saved!=='light')saved='light';applyTheme(saved)}
function getMainContent(){return document.querySelector('.main-content')}
function showTab(id){var el=document.getElementById(id);if(!el)return;if(showAllTables){document.querySelectorAll('.tab-link').forEach(function(b){b.classList.remove('active')});var allBtn=document.querySelector('.tab-link[data-target="'+id+'"]');if(allBtn)allBtn.classList.add('active');el.scrollIntoView({behavior:'smooth',block:'start'});return}document.querySelectorAll('.tab-panel').forEach(function(p){p.classList.remove('active')});el.classList.add('active');document.querySelectorAll('.tab-link').forEach(function(b){b.classList.remove('active')});var btn=document.querySelector('.tab-link[data-target="'+id+'"]');if(btn)btn.classList.add('active');var main=getMainContent();if(main)main.scrollTop=0}
function toggleShowAllTables(){showAllTables=!showAllTables;document.body.classList.toggle('show-all',showAllTables);var button=document.getElementById('showAllTables');if(showAllTables){if(button)button.textContent='Single Table View';document.querySelectorAll('.sidebar li.hidden').forEach(function(li){li.classList.remove('hidden')});document.querySelectorAll('.tab-link').forEach(function(b){b.classList.remove('active')});var main=getMainContent();if(main)main.scrollTop=0}else{if(button)button.textContent='Show All Tables';updateTabVisibility();showTab('report-overview')}}
function updateTabVisibility(){document.querySelectorAll('.sidebar li').forEach(function(li){var btn=li.querySelector('.tab-link');if(!btn)return;if(btn.dataset.target==='report-overview')return;if(btn.dataset.status==='healthy')li.classList.add('hidden');else li.classList.remove('hidden')})}
function toggleAllTabs(){var hidden=document.querySelectorAll('.sidebar li.hidden');if(hidden.length){hidden.forEach(function(li){li.classList.remove('hidden')})}else{updateTabVisibility()}}
function filterTabs(){var q=document.getElementById('tabSearch').value.toLowerCase().trim();document.querySelectorAll('.sidebar li').forEach(function(li){if(!q){li.style.display='';return}li.style.display=li.textContent.toLowerCase().indexOf(q)!==-1?'block':'none'})}
function sortTable(table,col){var tbody=table.tBodies&&table.tBodies[0];if(!tbody)return;var rows=Array.from(tbody.rows);if(!rows.length)return;var current=table.getAttribute('data-sort-col');var dir=(current===String(col)&&table.getAttribute('data-sort-dir')==='asc')?'desc':'asc';rows.sort(function(a,b){var av=a.cells[col]?a.cells[col].textContent.trim():'';var bv=b.cells[col]?b.cells[col].textContent.trim():'';var an=Number(av.replace(/,/g,'')),bn=Number(bv.replace(/,/g,''));var cmp=(!isNaN(an)&&!isNaN(bn)&&av!==''&&bv!=='')?(an-bn):av.localeCompare(bv,undefined,{numeric:true,sensitivity:'base'});return dir==='asc'?cmp:-cmp});rows.forEach(function(r){tbody.appendChild(r)});table.setAttribute('data-sort-col',String(col));table.setAttribute('data-sort-dir',dir);table.querySelectorAll('th').forEach(function(th,i){th.classList.toggle('sorted',i===col);var s=th.querySelector('.sort-indicator');if(s)s.textContent=String.fromCharCode(i===col?(dir==='asc'?9650:9660):8645)})}
function initTables(){document.querySelectorAll('.tab-panel table').forEach(function(table){table.querySelectorAll('th').forEach(function(th,col){if(!th.querySelector('.sort-indicator')){var span=document.createElement('span');span.className='sort-indicator';span.textContent=String.fromCharCode(8645);th.appendChild(span)}th.addEventListener('click',function(){sortTable(table,col)})})})}
document.addEventListener('DOMContentLoaded',function(){loadSavedTheme();initTables();updateTabVisibility();document.querySelector('.tab-link[data-target="report-overview"]')?.classList.add('active');document.querySelectorAll('a[href^="#"]').forEach(function(a){a.addEventListener('click',function(e){var id=this.getAttribute('href').slice(1);var target=document.getElementById(id);if(target&&target.classList.contains('tab-panel')){showTab(id);e.preventDefault()}})});var hash=window.location.hash.slice(1);if(hash)showTab(hash)});
</script>
</head>
<body>
<nav class='sidebar'>
  <div class='brand'><img id='slicLogo' src='$SlicLogoLightData' data-light-src='$SlicLogoLightData' data-dark-src='$SlicLogoDarkData' alt='SLIC' /></div>
  <div class='sidebar-controls'>
    <a href='#' id='themeToggle' class='sidebar-control' role='button' aria-pressed='false' onclick='toggleTheme();return false;'>Dark Mode</a>
    <a href='#' id='showAllTables' class='sidebar-control' onclick='toggleShowAllTables();return false;'>Show All Tables</a>
    <a href='#' id='expandAll' class='sidebar-control' onclick='toggleAllTabs();return false;'>Show / Hide Healthy Tables</a>
  </div>
  <div class='issue-link-wrap'><a class='issue-link' href='https://github.com/DellProSupportGse/source/issues' target='_blank' rel='noopener noreferrer'>Report an Issue &#8599;</a></div>
  <input type='text' id='tabSearch' oninput='filterTabs()' placeholder='Search tables...' />
  <ul>$($navItems -join '')</ul>
</nav>
<main class='main-content'>$($sectionHtml -join '')</main>
</body>
</html>
"@

        [System.IO.File]::WriteAllText($script:HtmlReportPath, $finalHtml, [System.Text.UTF8Encoding]::new($false))
        Write-Host "[+] Report saved to: $script:HtmlReportPath" -ForegroundColor Green
        Invoke-Item $script:HtmlReportPath
    }

    #endregion === HTML Report System ===

    do {
        $saveChoice = Read-Host "Save report to the SDDC folder ($SDDCPath)? [Y/N]"

        if ($saveChoice -match '^[Yy]$') {
            $OutputPath = $SDDCPath
            if (Test-Path $OutputPath) {
                Write-Host "[+] Report will be saved in: $OutputPath" -ForegroundColor Green
                $confirmed = $true
            } else {
                Write-Host "[!] The SDDC folder path does not exist: $OutputPath" -ForegroundColor Yellow
                $confirmed = $false
            }
        }
        elseif ($saveChoice -match '^[Nn]$') {
            $OutputPath = Read-Host "Please type the full folder path where you want to save the report"
            if (Test-Path $OutputPath) {
                Write-Host "[+] Report will be saved in: $OutputPath" -ForegroundColor Green
                $confirmed = $true
            } else {
                Write-Host "[!] Invalid path. Please try again or choose Y to use the SDDC folder." -ForegroundColor Yellow
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
    New-HtmlReport -Title "SLIC: Switch Log InspeCtor" -Version $Ver -RunDate (Get-Date) -OutputPath $OutputPath


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
            Write-Host "    WARNING: Could not locate the 'show running-configuration' section in $(Split-Path $Path -Leaf). Check header format in the log." -ForegroundColor Yellow
            return @()
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

    function Get-OS10InterfacesSections {
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
        #   ----------------------------------- show interface -------------------
        # then captures everything up to the next dashed "show ..." header or EOF.
        $pattern = '(?is)^\s*-{3,}\s*show\s+interface\s*-{3,}\s*\n(.*?)(?=^\s*-{3,}\s*show\s+\S.*?-{3,}\s*$|\Z)'
        $m = [regex]::Match($text, $pattern, 'IgnoreCase, Multiline, Singleline')
        #$m | ?{$_ -imatch "hostname"} | select Filename,@{L="HostName";E={($_.lines -imatch "hostname") -replace "hostname "}}
        if (-not $m.Success) {
            Write-Host "    WARNING: Could not locate the 'show interface' section in $(Split-Path $Path -Leaf). Check header format in the log." -ForegroundColor Yellow
            return @()
        }

        $run = $m.Groups[1].Value.Trim()

        # Split into sections where each section starts at a line that is just "!"
        # (Ignore blank chunks)
        $chunks = [regex]::Split($run,'(?:\r?\n){2,}') | Where-Object { $_ -and $_.Trim() -ne '' }
        #$SwitchHostname=""
        #$SwitchHostname = (((($chunks | ?{$_ -imatch "hostname"}) -split "hostname ")[-1] -split "`n")[0]).trim()

        $sections = New-Object System.Collections.Generic.List[object]
        $i = 0
        foreach ($chunk in $chunks) {
            $i++
            $lines  = ($chunk -split "`n") | ForEach-Object { $_.TrimEnd() }
            $header = ($lines | Where-Object { $_.Trim() -ne '' } | Select-Object -First 1)
            if (-not $header) { $header = "(empty)" }

            $sections.Add([pscustomobject]@{
                FileName   = $Path.split("/\")[-1]
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
            Write-Host "    WARNING: Could not locate the 'show version' section in $(Split-Path $Path -Leaf). Check header format in the log." -ForegroundColor Yellow
            return @()
        }

        $run = $m.Groups[1].Value.Trim()
        $lines   = $run -split "`n"
        $ShowVersionOut = ""
        $ShowVersionOut = ([pscustomobject]@{
            FileName   = $Path.split("/\")[-1]
            OSVersion  = (($lines | ?{$_ -imatch "OS Version:"}) -split ": ")[-1]
            SystemType = (($lines | ?{$_ -imatch "System Type:"}) -split ": ")[-1]
            UpTime     = (($lines | ?{$_ -imatch "Up Time:"}) -split ": ")[-1]
        })
        return $ShowVersionOut
    }

    $ShowVersions = @()
    $ShowVersions += $STSLOC | ForEach-Object{Get-showversion -path $_}

    #Create Ref Link for footnotes
        # Get unique system types
        $ShowVersionOut
        $SystemTypeUnique = $ShowVersions | Sort-Object SystemType -Unique | Select-Object -ExpandProperty SystemType

        # Map to reference link
        $SwitchRefLink = switch -Regex ($SystemTypeUnique) {
            "4112" { 'https://infohub.delltechnologies.com/en-us/l/switch-configurations-roce-and-iwarp-reference-guide-1/dell-networking-s4112f-on-switch-17/' ; break }
            "4148" { 'https://infohub.delltechnologies.com/en-us/l/switch-configurations-roce-and-iwarp-reference-guide-1/dell-networking-s4148f-on-switch-17/' ; break }
            "5148" { 'https://infohub.delltechnologies.com/en-us/l/switch-configurations-roce-and-iwarp-reference-guide-1/dell-networking-s5148f-on-switch-17/' ; break }
            "5212" { 'https://infohub.delltechnologies.com/en-us/l/switch-configurations-roce-and-iwarp-reference-guide-1/dell-networking-s5212f-on-switch-17/' ; break }
            "5232" { 'https://infohub.delltechnologies.com/en-us/l/switch-configurations-roce-and-iwarp-reference-guide-1/dell-networking-s5232f-on-switch-17/' ; break }
            "5248" { 'https://infohub.delltechnologies.com/en-us/l/switch-configurations-roce-and-iwarp-reference-guide-1/dell-networking-s5248f-on-switch-17/' ; break }
            default { 'https://infohub.delltechnologies.com/en-us/t/switch-configurations-roce-and-iwarp-reference-guide-1/' }
        }

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
                Write-Host "    WARNING: Could not locate the 'show lldp neighbors' section in $(Split-Path $Path -Leaf). Check header format in the log." -ForegroundColor Yellow
                return @()
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

            # Find NetIntent XML files first so a missing file can be handled cleanly.
            $NetIntentFiles = @(Get-ChildItem -Path $Path -Recurse -ErrorAction SilentlyContinue -Depth 2 -Filter GetNetIntent.XML)

            if (-not $NetIntentFiles -or $NetIntentFiles.Count -eq 0) {
                return @()
            }

            $NetIntentsXml = @($NetIntentFiles | Import-Clixml | Select-Object *,@{L="MADDR";E={$_.MacAddress -replace "-",":"}})
            return $NetIntentsXml
         }

         $GetNetIntents = @(Get-GetNetIntents -Path $SDDCPath)
         $NetIntentDataFound = ($GetNetIntents.Count -gt 0)
         $StorageNics = @()
         $MgmtNics = @()

         if (-not $NetIntentDataFound) {
            Write-Host "    [!] WARNING: No GetNetIntent.XML data was found in the SDDC." -ForegroundColor Yellow
            Write-Host "        Management and Storage intent NICs cannot be identified." -ForegroundColor Yellow
            Write-Host "        SLIC therefore cannot determine the Management and/or Storage switch ports." -ForegroundColor Yellow
         }
         else {
            $NetIntentStorageNicsInfo = @($GetNetIntents | Where-Object {$_.IsStorageIntentSet -eq $True} | Select-Object NetAdapterNamesAsList,StorageVLANs)

            foreach ($Intent in $NetIntentStorageNicsInfo) {
                for ($i = 0; $i -lt $Intent.NetAdapterNamesAsList.Count; $i++) {
                    $StorageNics += [pscustomobject]@{
                        NetAdapterName = $Intent.NetAdapterNamesAsList[$i]
                        VLAN           = $Intent.StorageVLANs[$i]
                    }
                }
            }

            $NetIntentMgmtNicsInfo = @($GetNetIntents | Where-Object {$_.IsManagementIntentSet -eq $True} | Select-Object NetAdapterNamesAsList,ManagementVLAN)

            foreach ($Intent in $NetIntentMgmtNicsInfo) {
                for ($i = 0; $i -lt $Intent.NetAdapterNamesAsList.Count; $i++) {
                    $MgmtNics += [pscustomobject]@{
                        NetAdapterName = $Intent.NetAdapterNamesAsList[$i]
                        VLAN           = $Intent.ManagementVLAN
                    }
                }
            }
         }

            # Display
            #$MgmtNics | Format-Table


         $GetNetAdapterInfos = Get-GetNetAdapterInfo -path $SDDCPath
         # Find which Qos Priorities the nodes are using to compare later
            $GetNetQOSPolicyInfo = Get-ChildItem -Path $SDDCPath -Recurse -Filter GetNetQOSPolicy.xml | Import-Clixml
            $GetNetQOSPolicyPriorities = $GetNetQOSPolicyInfo | Sort-Object PriorityValue -Unique | select PriorityValue
         $ShowRunningConfigs = $STSLOC | ForEach-Object { Get-OS10RunningConfigSections -Path $_ }
         $ShowRunningConfigs = $ShowRunningConfigs | ?{$_.hostname -ne "False"}
         $ShowInterface = $STSLOC | ForEach-Object { Get-OS10InterfacesSections -Path $_ }

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

        # If Network ATC intent data is unavailable, allow the operator to manually
        # identify the server-facing switch ports discovered through LLDP/GetNetAdapter.
        $ManualPortClassificationUsed = $false

        if ((-not $NetIntentDataFound) -and $SwPortToHostMap.Count -gt 0) {
            Write-Host ""
            Write-Host "    [!] GetNetIntent.XML is unavailable. Manual switch-port classification is required." -ForegroundColor Yellow
            Write-Host "        The following server-facing switch ports were correlated using LLDP and GetNetAdapter.xml:" -ForegroundColor Yellow
            Write-Host ""

            $ManualCandidates = @($SwPortToHostMap |
                Sort-Object SwHostName, SwLocPortId, ComputerName, ifAlias |
                Select-Object SwHostName, SwLocPortId, ComputerName, ifAlias, ifDesc, MacAddress -Unique)

            $IndexedCandidates = for ($Index = 0; $Index -lt $ManualCandidates.Count; $Index++) {
                [pscustomobject]@{
                    ID           = $Index + 1
                    Switch       = $ManualCandidates[$Index].SwHostName
                    SwitchPort   = $ManualCandidates[$Index].SwLocPortId
                    ComputerName = $ManualCandidates[$Index].ComputerName
                    Adapter      = $ManualCandidates[$Index].ifAlias
                    Description  = $ManualCandidates[$Index].ifDesc
                    MacAddress   = $ManualCandidates[$Index].MacAddress
                }
            }

            $IndexedCandidates | Format-Table -AutoSize -Wrap
            Write-Host ""
            Write-Host "    Classify each port: M = Management, S = Storage, B = Both, O = Other/Skip" -ForegroundColor Cyan
            Write-Host "    VLAN values are requested so the existing switch configuration checks can continue." -ForegroundColor Cyan
            Write-Host ""

            $ManualPortMap = @()

            foreach ($Candidate in $ManualCandidates) {
                $PromptLabel = "{0} / {1} -> {2} / {3}" -f $Candidate.SwHostName, $Candidate.SwLocPortId, $Candidate.ComputerName, $Candidate.ifAlias

                do {
                    $RoleChoice = (Read-Host "Role for $PromptLabel [M/S/B/O]").Trim().ToUpperInvariant()
                } until ($RoleChoice -match '^[MSBO]$')

                if ($RoleChoice -eq 'O') {
                    $ManualPortMap += [pscustomobject]@{
                        SwHostName      = $Candidate.SwHostName
                        SwLocPortId     = $Candidate.SwLocPortId
                        ComputerName    = $Candidate.ComputerName
                        ifAlias         = $Candidate.ifAlias
                        ifDesc          = $Candidate.ifDesc
                        MacAddress      = $Candidate.MacAddress
                        IntentType      = 'Other'
                        vLAN            = ''
                        AssignmentSource = 'Manual'
                    }
                    continue
                }

                if ($RoleChoice -in @('M','B')) {
                    do {
                        $MgmtVlan = (Read-Host "  Management VLAN for $($Candidate.SwHostName) $($Candidate.SwLocPortId)").Trim()
                        if ($MgmtVlan -notmatch '^\d{1,4}$') {
                            Write-Host "    Please enter a numeric VLAN ID (for example 201)." -ForegroundColor Yellow
                        }
                    } until ($MgmtVlan -match '^\d{1,4}$')

                    $ManualPortMap += [pscustomobject]@{
                        SwHostName      = $Candidate.SwHostName
                        SwLocPortId     = $Candidate.SwLocPortId
                        ComputerName    = $Candidate.ComputerName
                        ifAlias         = $Candidate.ifAlias
                        ifDesc          = $Candidate.ifDesc
                        MacAddress      = $Candidate.MacAddress
                        IntentType      = 'Mgmt'
                        vLAN            = $MgmtVlan
                        AssignmentSource = 'Manual'
                    }
                }

                if ($RoleChoice -in @('S','B')) {
                    do {
                        $StorageVlan = (Read-Host "  Storage VLAN for $($Candidate.SwHostName) $($Candidate.SwLocPortId)").Trim()
                        if ($StorageVlan -notmatch '^\d{1,4}$') {
                            Write-Host "    Please enter a numeric VLAN ID (for example 711)." -ForegroundColor Yellow
                        }
                    } until ($StorageVlan -match '^\d{1,4}$')

                    $ManualPortMap += [pscustomobject]@{
                        SwHostName      = $Candidate.SwHostName
                        SwLocPortId     = $Candidate.SwLocPortId
                        ComputerName    = $Candidate.ComputerName
                        ifAlias         = $Candidate.ifAlias
                        ifDesc          = $Candidate.ifDesc
                        MacAddress      = $Candidate.MacAddress
                        IntentType      = 'Storage'
                        vLAN            = $StorageVlan
                        AssignmentSource = 'Manual'
                    }
                }
            }

            $SwPortToHostMap = @($ManualPortMap)
            $ManualPortClassificationUsed = $true

            Write-Host ""
            Write-Host "    [+] Manual switch-port classification complete:" -ForegroundColor Green
            $SwPortToHostMap | Format-Table SwHostName, SwLocPortId, ComputerName, ifAlias, IntentType, vLAN -AutoSize
            Write-Host ""
        }
        # Add Interface-to-Node Map to the HTML report.
        if ((-not $NetIntentDataFound) -and $ManualPortClassificationUsed) {
            $Description = @"
<div class='warning-banner'>
  <b>WARNING: NetIntent data not found - manual port classification used.</b><br>
  SLIC could not find any <b>GetNetIntent.XML</b> data in the supplied SDDC.<br>
  The Management/Storage roles and VLANs shown below were entered manually by the operator.
</div>
"@

            AddTo-HtmlReport -Title "Interface-to-Node Map" `
                -Data $SwPortToHostMap `
                -Description $Description `
                -Footnotes "Port roles were manually assigned because GetNetIntent.XML was unavailable. LLDP and GetNetAdapter.xml were used to correlate server-facing switch ports. <p><a href='#'>Go to top</a></p>" `
                -IncludeTitle -IncludeDescription -IncludeFootnotes
        }
        elseif (-not $NetIntentDataFound) {
            $Description = @"
<div class='warning-banner'>
  <b>WARNING: NetIntent data not found and no server-facing ports could be correlated.</b><br>
  SLIC could not find any <b>GetNetIntent.XML</b> data and could not build a candidate switch-port list from LLDP/GetNetAdapter.xml.
  Management and Storage switch ports cannot be determined.
</div>
"@

            AddTo-HtmlReport -Title "Interface-to-Node Map" `
                -Data @() `
                -Description $Description `
                -Footnotes "Manual classification requires matching LLDP and GetNetAdapter.xml information. <p><a href='#'>Go to top</a></p>" `
                -IncludeTitle -IncludeDescription -IncludeFootnotes
        }
        elseif ($SwPortToHostMap) {
            AddTo-HtmlReport -Title "Interface-to-Node Map" `
                -Data $SwPortToHostMap `
                -Description "" `
                -Footnotes "Highlighted in red or yellow if out of spec. <p><a href='$SwitchRefLink' target='_blank'>Ref: Switch Configurations - RoCE/iWarp Reference Guide</a></p><p><a href='#'>Go to top</a></p>" `
                -IncludeTitle -IncludeDescription -IncludeFootnotes
        }
        else {
            Write-Host "    [!] WARNING: NetIntent data was found, but no switch-to-node port matches were found." -ForegroundColor Yellow
            Write-Host "        Suspect the supplied SDDC does not correspond to these show tech file(s), or LLDP data is incomplete." -ForegroundColor Yellow
            $Description = "<div class='warning-banner'><b>WARNING:</b> NetIntent data was found, but no switch-to-node port matches were found. Suspect the supplied SDDC does not correspond to these show tech file(s), or LLDP data is incomplete.</div>"

            AddTo-HtmlReport -Title "Interface-to-Node Map" `
                -Data @() `
                -Description $Description `
                -Footnotes "The switch-port map requires matching NetIntent, NetAdapter, and LLDP information. <p><a href='$SwitchRefLink' target='_blank'>Ref: Switch Configurations - RoCE/iWarp Reference Guide</a></p><p><a href='#'>Go to top</a></p>" `
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
                -Footnotes "Highlighted in red or yellow if out of spec. <p><a href='$SwitchRefLink' target='_blank'>Ref: Switch Configurations - RoCE/iWarp Reference Guide</a></p><p><a href='#'>Go to top</a></p>" `
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
                    -Footnotes "Highlighted in red or yellow if out of spec. <p><a href='$SwitchRefLink' target='_blank'>Ref: Switch Configurations - RoCE/iWarp Reference Guide</a></p><p><a href='#'>Go to top</a></p>" `
                    -IncludeTitle -IncludeDescription -IncludeFootnotes
            }

            #Check for #class-map type queuing Q5/7
            ####Add 5 or 7
            IF($classmaptypequeuingQ | ?{$_.lines -imatch 'class-map type queuing Q(5|7)'}){
                $classmaptypequeuingQ57 += $classmaptypequeuingQ | ?{$_.lines -imatch 'class-map type queuing Q5'} | select Filename, SwHostName,
                    @{L="class-map type queuing Q5";E={IF($_.Lines -imatch 'class-map type queuing Q5'){"Found"}Else{"RREEDDMissing"}}},
                    @{L="match queue 5";            E={IF($_.lines -imatch 'match queue 5'){"Found"}Else{"RREEDDMissing"}}}
                
                $classmaptypequeuingQ57 += $classmaptypequeuingQ | ?{$_.lines -imatch 'class-map type queuing Q7'} | select Filename, SwHostName,
                    @{L="class-map type queuing Q7";E={IF($_.Lines -imatch 'class-map type queuing Q7'){"Found"}Else{"RREEDDMissing"}}},
                    @{L="match queue 7";            E={IF($_.lines -imatch 'match queue 7'){"Found"}Else{"RREEDDMissing"}}}
            }Else{
                #Write-Host "     FAIL: Missing both class-map type queuing Q5 and Q7. Assume Q7" -ForegroundColor red
                $classmaptypequeuingQ57 += $classmaptypequeuingQ | sort FileName -Unique | select Filename, SwHostName,@{L="class-map type queuing Q5/7";E={"RREEDDMissing both"}},@{L="match queue 5/7";E={"RREEDDMissing both"}}
            }
            #$classmaptypequeuingQ7 | ft
            # Add to HTML report output sections
            if($classmaptypequeuingQ57){
                AddTo-HtmlReport -Title "class-map type queuing" `
                    -Data $classmaptypequeuingQ57 `
                    -Description "" `
                    -Footnotes "Highlighted in red or yellow if out of spec. <p><a href='$SwitchRefLink' target='_blank'>Ref: Switch Configurations - RoCE/iWarp Reference Guide</a></p><p><a href='#'>Go to top</a></p>" `
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
                -Footnotes "Highlighted in red or yellow if out of spec. <p><a href='$SwitchRefLink' target='_blank'>Ref: Switch Configurations - RoCE/iWarp Reference Guide</a></p><p><a href='#'>Go to top</a></p>" `
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
                        -Footnotes "Highlighted in red or yellow if out of spec. <p><a href='$SwitchRefLink' target='_blank'>Ref: Switch Configurations - RoCE/iWarp Reference Guide</a></p><p><a href='#'>Go to top</a></p>" `
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
                        -Footnotes "Highlighted in red or yellow if out of spec. <p><a href='$SwitchRefLink' target='_blank'>Ref: Switch Configurations - RoCE/iWarp Reference Guide</a></p><p><a href='#'>Go to top</a></p>" `
                        -IncludeTitle -IncludeDescription -IncludeFootnotes
                }
            #match qos-group 5
                $matchqosgroup5 = $classmaptypenetworkqosManagement | ?{$_.lines -imatch "match qos-group 5"}
                IF($matchqosgroup5){
                    $matchqosgroup5out = $matchqosgroup5 | sort FileName -Unique | select Filename, SwHostName,@{L=$matchqosgroup5.lines[1];E={
                        IF($_.lines -imatch "class-map type network-qos"){"Found"}Else{"RREEDDMissing"}}},
                        @{L="match qos-group 5";E={IF($_.lines -imatch "match qos-group 5"){
                            If($GetNetQOSPolicyPriorities.PriorityValue -imatch "5"){"Match Switch=Q5 Server=Q5"}
                            ElseIf($GetNetQOSPolicyPriorities.PriorityValue -imatch "7"){"RREEDDMismatch Switch=Q5 Server=Q7"}}}}
                    #$matchqosgroup5out | ft
                    # Add to HTML report output sections
                    if($matchqosgroup5out){
                        AddTo-HtmlReport -Title "class-map type network-qos group" `
                            -Data $matchqosgroup5out `
                            -Description "" `
                            -Footnotes "Highlighted in red or yellow if out of spec. <p><a href='$SwitchRefLink' target='_blank'>Ref: Switch Configurations - RoCE/iWarp Reference Guide</a></p><p><a href='#'>Go to top</a></p>" `
                            -IncludeTitle -IncludeDescription -IncludeFootnotes
                    }
                }
            #Match qos-group 7
                $matchqosgroup7 = $classmaptypenetworkqosManagement | ?{$_.lines -imatch "match qos-group 7"}
                IF($matchqosgroup7){
                    $matchqosgroup7out = $matchqosgroup7 | sort FileName -Unique | select Filename, SwHostName,@{L=$matchqosgroup7.lines[1];E={
                        IF($_.lines -imatch "class-map type network-qos"){"Found"}Else{"RREEDDMissing"}}},
                        @{L="match qos-group 7";E={IF($_.lines -imatch "match qos-group 7"){
                            If($GetNetQOSPolicyPriorities.PriorityValue -imatch "7"){"Match Switch=Q7 Server=Q7"}
                            ElseIf($GetNetQOSPolicyPriorities.PriorityValue -imatch "5"){"RREEDDMismatch Switch=Q7 Server=Q5"}}}}
                    #$matchqosgroup7out | ft
                    # Add to HTML report output sections
                    if($matchqosgroup7out){
                        AddTo-HtmlReport -Title "class-map type network-qos group" `
                            -Data $matchqosgroup7out `
                            -Description "" `
                            -Footnotes "Highlighted in red or yellow if out of spec. <p><a href='$SwitchRefLink' target='_blank'>Ref: Switch Configurations - RoCE/iWarp Reference Guide</a></p><p><a href='#'>Go to top</a></p>" `
                            -IncludeTitle -IncludeDescription -IncludeFootnotes
                    }
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
                    -Footnotes "Highlighted in red or yellow if out of spec. <p><a href='$SwitchRefLink' target='_blank'>Ref: Switch Configurations - RoCE/iWarp Reference Guide</a></p><p><a href='#'>Go to top</a></p>" `
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
                @{L="qos-group 0 dot1p";E={
                    # Check for dop1p7
                    IF($_.lines -imatch "qos-group 0 dot1p 0-2,4-6"){
                        If($GetNetQOSPolicyPriorities.PriorityValue -imatch "7"){"Match qos-group 0 dot1p 0-2,4-6"}}
                    # Check for dop1p5
                    ElseIF($_.lines -imatch "qos-group 0 dot1p 0-2,4,6-7"){
                        IF($GetNetQOSPolicyPriorities.PriorityValue -imatch "5"){"Match qos-group 0 dot1p 0-2,4,6-7"}}}},
                @{L="qos-group 3 dot1p 3";E={IF($_.lines -imatch "qos-group 3 dot1p 3"){"Match qos-group 3 dot1p 3"}Else{"RREEDDMissing"}}},
                @{L="qos-group 5/7 dot1p 5/7";E={
                    IF($_.lines -imatch "qos-group 7 dot1p 7"){
                    #Q7
                        If($GetNetQOSPolicyPriorities.PriorityValue -imatch "7"){"Match qos-group 7 dot1p 7"}
                        ElseIf($GetNetQOSPolicyPriorities.PriorityValue -imatch "7"){"RREEDDMismatch Switch=Q5 Server=Q7"}}
                    ElseIf($_.lines -imatch "qos-group 5 dot1p 5"){
                    #Q5
                        If($GetNetQOSPolicyPriorities.PriorityValue -imatch "5"){"Match qos-group 5 dot1p 5"}
                        ElseIf($GetNetQOSPolicyPriorities.PriorityValue -imatch "5"){"RREEDDMismatch Switch=Q7 Server=Q5"}}
                    #No Q5 or 7
                    ElseIf(($_.lines -inotmatch "qos-group 7 dot1p 7") -and ($_.lines -inotmatch "qos-group 5 dot1p 5")){"RREEDDMissing"}}}
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
                -Footnotes "Highlighted in red or yellow if out of spec. <p><a href='$SwitchRefLink' target='_blank'>Ref: Switch Configurations - RoCE/iWarp Reference Guide</a></p><p><a href='#'>Go to top</a></p>" `
                -IncludeTitle -IncludeDescription -IncludeFootnotes
        }

        #qos-map traffic-class queue-map
        $qosmaptrafficclassqueuemap = @()
        $qosmaptrafficclassqueuemap = $ShowRunningConfigs | ?{$_.header -imatch "qos-map traffic-class queue-map"}
        IF($qosmaptrafficclassqueuemap){
            $qosmaptrafficclassqueuemapOut = $qosmaptrafficclassqueuemap | sort FileName -Unique | select Filename, SwHostName,
                @{L="qos-map traffic-class queue-map";E={IF($_.lines -imatch "qos-map traffic-class queue-map"){"Found"}Else{"RREEDDMissing"}}},
                @{L="queue 0 qos-group 0-2,4-6/6-7";E={
                    IF($_.lines -imatch "queue 0 qos-group 0-2,4,6-7"){"Match queue 0 qos-group 0-2,4,6-7"}
                    ElseIf($_.lines -imatch "queue 0 qos-group 0-2,4-6"){"Match queue 0 qos-group 0-2,4-6"}
                    ElseIF(($_.lines -inotmatch "queue 0 qos-group 0-2,4,6-7") -and ($_.lines -inotmatch "queue 0 qos-group 0-2,4-6")){"Mismatch "+$_.Line[2]}}},
                @{L="queue 3 qos-group 3";E={IF($_.lines -imatch " queue 3 qos-group 3"){"Found"}Else{"RREEDDMissing"}}},
                @{L="queue 5/7 qos-group 5/7";E={
                    IF($_.lines -imatch "queue 5 qos-group 5"){
                        #Does Server Qos Policy Match Switch Q
                            If($GetNetQOSPolicyPriorities.PriorityValue -imatch "5"){"Match Switch=Q5 Server=Q5"}
                            ElseIf($GetNetQOSPolicyPriorities.PriorityValue -imatch "7"){"RREEDDMismatch Switch=Q5 Server=Q7"}}                 
                    ElseIf($_.lines -imatch "queue 7 qos-group 7"){
                        #Does Server Qos Policy Match Switch Q
                            If($GetNetQOSPolicyPriorities.PriorityValue -imatch "7"){"Match Switch=Q7 Server=Q7"}
                            ElseIf($GetNetQOSPolicyPriorities.PriorityValue -imatch "5"){"RREEDDMismatch Switch=Q7 Server=Q5"}}
                    ElseIf(($_.lines -inotmatch "queue 7 qos-group 7") -and ($_.lines -inotmatch "queue 5 qos-group 5")){"RREEDDMissing"}}}
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
                -Footnotes "Highlighted in red or yellow if out of spec. <p><a href='$SwitchRefLink' target='_blank'>Ref: Switch Configurations - RoCE/iWarp Reference Guide</a></p><p><a href='#'>Go to top</a></p>" `
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
                -Footnotes "Highlighted in red or yellow if out of spec. <p><a href='$SwitchRefLink' target='_blank'>Ref: Switch Configurations - RoCE/iWarp Reference Guide</a></p><p><a href='#'>Go to top</a></p>" `
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
                -Footnotes "Highlighted in red or yellow if out of spec. <p><a href='$SwitchRefLink' target='_blank'>Ref: Switch Configurations - RoCE/iWarp Reference Guide</a></p><p><a href='#'>Go to top</a></p>" `
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
            #Case 1 we have a Q5 and no Q7
            IF($classQ0357 | ?{$_.Header -imatch "Q5" -and $_.Header -inotmatch "Q7"}){
                $classQ5Out = $classQ0357 | ?{$_.Header -imatch "Q5"} | select Filename, SwHostName,
                    @{L="class Q5";E={IF($_.lines -imatch "class Q5"){
                        #Does Server Qos Policy Match Switch Q
                            If($GetNetQOSPolicyPriorities.PriorityValue -imatch "5"){"Match Switch=Q5 Server=Q5"}
                            ElseIf($GetNetQOSPolicyPriorities.PriorityValue -imatch "7"){"RREEDDMismatch Switch=Q5 Server=Q7"}}}},
                    @{L="bandwidth percent 1 or 2";E={IF($_.lines -imatch "bandwidth percent (1|2)"){"Found"}Else{"RREEDDMissing"}}}
            }
            #Case 2 we have a Q7 and no Q5
            IF($classQ0357 | ?{$_.Header -imatch "Q7" -and $_.Header -inotmatch "Q5"}){
                    $classQ7Out = $classQ0357 | ?{$_.Header -imatch "Q7"} | select Filename, SwHostName,
                        @{L="class Q7";E={IF($_.lines -imatch "class Q7"){
                        #Does Server Qos Policy Match Switch Q
                            If($GetNetQOSPolicyPriorities.PriorityValue -imatch "7"){"Match Switch=Q7 Server=Q7"}
                            ElseIf($GetNetQOSPolicyPriorities.PriorityValue -imatch "5"){"RREEDDMismatch Switch=Q7 Server=Q5"}}}},
                        @{L="bandwidth percent 1 or 2";E={IF($_.lines -imatch "bandwidth percent (1|2)"){"Found"}Else{"RREEDDMissing"}}}
            }
            #Case 3 no Q5 and no Q7
            IF(!($classQ5Out) -and ($classQ7Out)){
                $classQ57Out = $classQ0357 | select Filename, SwHostName,
                    @{L="class Q5/7";E={IF($_.lines -imatch "class Q5/7"){"Found"}Else{"RREEDDMissing Q5 and Q7"}}},
                    @{L="bandwidth percent 1 or 2";E={IF($_.lines -imatch "bandwidth percent (1|2)"){"Found"}Else{"RREEDDMissing"}}}
            }
        }Else{
            #no policy-map type queuing ets-policy
            $classQ0357Out = $ShowRunningConfigs | sort FileName | select Filename, SwHostName,@{L="class Q0|3|5|7";E={"RREEDDMissing"}}
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
                -Footnotes "Highlighted in red or yellow if out of spec. <p><a href='$SwitchRefLink' target='_blank'>Ref: Switch Configurations - RoCE/iWarp Reference Guide</a></p><p><a href='#'>Go to top</a></p>" `
                -IncludeTitle -IncludeDescription -IncludeFootnotes
        }
        if($classQ3Out){
            AddTo-HtmlReport -Title "policy-map type queuing ets-policy class Q3" `
                -Data $classQ3Out `
                -Description "" `
                -Footnotes "Highlighted in red or yellow if out of spec. <p><a href='$SwitchRefLink' target='_blank'>Ref: Switch Configurations - RoCE/iWarp Reference Guide</a></p><p><a href='#'>Go to top</a></p>" `
                -IncludeTitle -IncludeDescription -IncludeFootnotes
        }
        if($classQ57Out){
            AddTo-HtmlReport -Title "policy-map type queuing ets-policy class" `
                -Data $classQ57Out `
                -Description "" `
                -Footnotes "Highlighted in red or yellow if out of spec. <p><a href='$SwitchRefLink' target='_blank'>Ref: Switch Configurations - RoCE/iWarp Reference Guide</a></p><p><a href='#'>Go to top</a></p>" `
                -IncludeTitle -IncludeDescription -IncludeFootnotes
        }
        if($classQ5Out){
            AddTo-HtmlReport -Title "policy-map type queuing ets-policy class" `
                -Data $classQ5Out `
                -Description "" `
                -Footnotes "Highlighted in red or yellow if out of spec. <p><a href='$SwitchRefLink' target='_blank'>Ref: Switch Configurations - RoCE/iWarp Reference Guide</a></p><p><a href='#'>Go to top</a></p>" `
                -IncludeTitle -IncludeDescription -IncludeFootnotes
        }
        if($classQ7Out){
            AddTo-HtmlReport -Title "policy-map type queuing ets-policy class" `
                -Data $classQ7Out `
                -Description "" `
                -Footnotes "Highlighted in red or yellow if out of spec. <p><a href='$SwitchRefLink' target='_blank'>Ref: Switch Configurations - RoCE/iWarp Reference Guide</a></p><p><a href='#'>Go to top</a></p>" `
                -IncludeTitle -IncludeDescription -IncludeFootnotes
        }
        if($classQ0357Out){
            AddTo-HtmlReport -Title "policy-map type queuing ets-policy class Q" `
                -Data $classQ0357Out `
                -Description "" `
                -Footnotes "Highlighted in red or yellow if out of spec. <p><a href='$SwitchRefLink' target='_blank'>Ref: Switch Configurations - RoCE/iWarp Reference Guide</a></p><p><a href='#'>Go to top</a></p>" `
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
                -Footnotes "Highlighted in red or yellow if out of spec. <p><a href='$SwitchRefLink' target='_blank'>Ref: Switch Configurations - RoCE/iWarp Reference Guide</a></p><p><a href='#'>Go to top</a></p>" `
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
                -Footnotes "Highlighted in red or yellow if out of spec. <p><a href='$SwitchRefLink' target='_blank'>Ref: Switch Configurations - RoCE/iWarp Reference Guide</a></p><p><a href='#'>Go to top</a></p>" `
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
                -Footnotes "Highlighted in red or yellow if out of spec. <p><a href='$SwitchRefLink' target='_blank'>Ref: Switch Configurations - RoCE/iWarp Reference Guide</a></p><p><a href='#'>Go to top</a></p>" `
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
                Write-Host "Detected PowerShell ISE - opening '$Title' in Out-GridView..."
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

                Write-Host "=== $Title ==="
                Write-Host "Console width set to $width characters (scroll horizontally to view)."
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
                                                                       #Check for storage vlan
                                                                        if($_ -imatch [regex]::Escape($StorageUsedInterface.vLAN.ToString())){$_}else{"RREEDD"+$_}
                                                                       #We should NOT have Mgmt vLANs in storage trunk ex: switchport trunk allowed vlan 201,711-712,1701-1702,3939 where 201=Mgmt
                                                                        IF($MgmtvLans){
                                                                         IF($_ -imatch ($MgmtvLans -join '|')){"RREEDD"+$_}
                                                                        }
                                                                 }}).'switchport trunk allowed vlan'
            'MTU9216'                                          = If ((Get-LineValue $StorageUsedInterface.Lines '9216').value -ne '9216') {
                                                                    $newheader=$StorageUsedInterface.header.split(" ")[-1]
                                                                    $newheader=$newheader.substring(0,($newheader | Select-String "\d").matches[0].index) + " " + $newheader.substring(($newheader | Select-String "\d").matches[0].index)
                                                                    (($ShowInterface | ? Header -match $newheader).lines | Select-String "MTU\s(\d*)\sbytes").matches.Groups[1].Value
                                                               } else {Get-LineValue $StorageUsedInterface.Lines '9216'} 
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
    if ($StorageUsedInterfacesOut.Count -gt 0) {
        $StorageUsedInterfacesEasyOut = Convert-ToSwitchComparisonTable -Interfaces $StorageUsedInterfacesOut | sort ShouldBe
        #$StorageUsedInterfacesEasyOut | Show-WideTable -Title "Storage Switch Port Comparison"
        # Add to HTML report output sections
        AddTo-HtmlReport -Title "Storage Interfaces" `
            -Data $StorageUsedInterfacesEasyOut `
            -Description "" `
            -Footnotes "Highlighted in red or yellow if out of spec. <p><a href='$SwitchRefLink' target='_blank'>Ref: Switch Configurations - RoCE/iWarp Reference Guide</a></p><p><a href='#'>Go to top</a></p>" `
            -IncludeTitle -IncludeDescription -IncludeFootnotes
    } else {
        Write-Host "    WARNING: No storage interfaces found to analyze." -ForegroundColor Yellow
        $StorageUsedInterfacesEasyOut = @([PSCustomObject]@{Message = "No storage interfaces found"})
        AddTo-HtmlReport -Title "Storage Interfaces" `
            -Data $StorageUsedInterfacesEasyOut `
            -Description "No storage interfaces were found in the switch configurations." `
            -Footnotes "Check that the SDDC data matches the switch tech-support files." `
            -IncludeTitle -IncludeDescription -IncludeFootnotes
    }

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
                                                    #Check for Mgmt vlan
                                                     if($_ -imatch [regex]::Escape($MgmtUsedInterface.vLAN.ToString())){$_}Else{"RREEDD"+$_}
                                                    #We should NOT have storage vLANs in storage trunk ex: switchport trunk allowed vlan 201,711-712,1701-1702,3939 where 201=Mgmt
                                                     IF($Storagevlans){
                                                      if($_ -imatch ($Storagevlans -join '|')){"RREEDD"+$_}
                                                     }
                                                }}).'switchport trunk allowed vlan'
            'MTU9216'                         = If ((Get-LineValue $MgmtUsedInterface.Lines '9216') -ne '9216') {
                                                       $newheader=$MgmtUsedInterface.header.split(" ")[-1]
                                                       $newheader=$newheader.substring(0,($newheader | Select-String "\d").matches[0].index) + " " + $newheader.substring(($newheader | Select-String "\d").matches[0].index)
                                                           (($ShowInterface | ? Header -match $newheader).lines | Select-String "MTU\s(\d*)\sbytes").matches.Groups[1].Value
                                                       } else {Get-LineValue $MgmtUsedInterface.Lines '9216'}
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
    if ($MgmtUsedInterfacesOut.Count -gt 0) {
        $MgmtUsedInterfacesEasyOut = Convert-ToSwitchComparisonTable -Interfaces $MgmtUsedInterfacesOut | sort ShouldBe
        #$MgmtUsedInterfacesEasyOut | Show-WideTable -Title "Mgmt Switch Port Comparison"
        # Add to HTML report output sections
        AddTo-HtmlReport -Title "Mgmt Interfaces" `
            -Data $MgmtUsedInterfacesEasyOut `
            -Description "" `
            -Footnotes "Highlighted in red or yellow if out of spec. <p><a href='$SwitchRefLink' target='_blank'>Ref: Switch Configurations - RoCE/iWarp Reference Guide</a></p><p><a href='#'>Go to top</a></p>" `
            -IncludeTitle -IncludeDescription -IncludeFootnotes
    } else {
        Write-Host "    WARNING: No management interfaces found to analyze." -ForegroundColor Yellow
        $MgmtUsedInterfacesEasyOut = @([PSCustomObject]@{Message = "No management interfaces found"})
        AddTo-HtmlReport -Title "Mgmt Interfaces" `
            -Data $MgmtUsedInterfacesEasyOut `
            -Description "No management interfaces were found in the switch configurations." `
            -Footnotes "Check that the SDDC data matches the switch tech-support files." `
            -IncludeTitle -IncludeDescription -IncludeFootnotes
    }

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
                    Write-Host "    INFO: No VLT data found in $(Split-Path $Path -Leaf)" -ForegroundColor Yellow
                    return @()
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

            # Extract section - handle both "show port-channel summary" and "show interface port-channel summary"
            $pattern = '(?is)^\s*-{3,}\s*show\s+(?:interface\s+)?port-channel\s+summary\s*-{3,}\s*\n(.*?)(?=^\s*-{3,}\s*show\s+\S.*?-{3,}\s*$|\Z)'
            $m = [regex]::Match($text, $pattern, 'IgnoreCase, Multiline, Singleline')
            if (-not $m.Success) {
                Write-Host "    INFO: No port-channel summary section found in $(Split-Path $Path -Leaf)" -ForegroundColor Yellow
                return @()
            }

            $section = $m.Groups[1].Value.Trim()
            $lines = $section -split "`n" | ForEach-Object { $_.TrimEnd() } | Where-Object { $_ }

            # Find header
            $headerLine = $lines | Where-Object { $_ -match '^\s*Group\s+Port-Channel' }
            if (-not $headerLine) { 
                # If no data found, return empty array instead of throwing
                Write-Host "    INFO: No port-channel data found in $(Split-Path $Path -Leaf)" -ForegroundColor Yellow
                return @()
            }

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
            try {
                $showvltall += $STSLOC | ForEach-Object { Get-showvltall -Path $_ }
            } catch {
                Write-Host "    WARNING: Could not parse VLT information: $_" -ForegroundColor Yellow
                $showvltall = @()
            }

    

        #$showportchannelsummary = Get-showportchannelsummary -path "C:\Users\Jim_Gandy\Downloads\tk5tor17-01a-show-tech-20251021-134546.txt"
        try {
            $showportchannelsummary = $STSLOC | ForEach-Object { Get-showportchannelsummary -Path $_ }
            $VLTiPorts = ($showportchannelsummary | ?{$_.'Group Port-Channel' -imatch ($showvltall | select port-channel* | GM | ?{$_.MemberType -eq "NoteProperty"}).Name} | select @{L="VLTi Ports";E={$_.'Member Ports' -split " "-replace '\([A-Z]+\)', '' }}).'VLTi Ports'| sort -Unique
        } catch {
            Write-Host "    WARNING: Could not parse port-channel summary: $_" -ForegroundColor Yellow
            $showportchannelsummary = @()
            $VLTiPorts = @()
        }
        $VLTiUsedInterfaces = @()
        if ($VLTiPorts.Count -gt 0) {
            ForEach($VLTiPort in $VLTiPorts){
                ForEach($Interface in $ShowRunningConfigs){
                    $VLTiUsedInterfaces += $Interface | ?{$_.Header -imatch "interface ethernet"+$VLTiPort}
                }
            }
        } else {
            Write-Host "    INFO: No VLTi ports found in port-channel summary" -ForegroundColor Yellow
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
                        'MTU9216'                                          = If ((Get-LineValue $VLTiUsedInterface.Lines '9216') -ne '9216') {
                                                                                $newheader=$VLTiUsedInterface.header.split(" ")[-1]
                                                                                $newheader=$newheader.substring(0,($newheader | Select-String "\d").matches[0].index) + " " + $newheader.substring(($newheader | Select-String "\d").matches[0].index)
                                                                                (($ShowInterface | ? Header -match $newheader).lines | Select-String "MTU\s(\d*)\sbytes").matches.Groups[1].Value
                                                                             } else {Get-LineValue $VLTiUsedInterface.Lines '9216'}
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

    if ($VLTiUsedInterfacesOut.Count -gt 0) {
        $VLTiUsedInterfacesEasyOut = Convert-ToSwitchComparisonTable -Interfaces $VLTiUsedInterfacesOut  | sort ShouldBe
        #$VLTiUsedInterfacesEasyOut | Show-WideTable -Title "VLTi Switch Port Comparison"
        # Add to HTML report output sections
        AddTo-HtmlReport -Title "VLTi Interfaces" `
            -Data $VLTiUsedInterfacesEasyOut `
            -Description "" `
            -Footnotes "Highlighted in red or yellow if out of spec. <p><a href='$SwitchRefLink' target='_blank'>Ref: Switch Configurations - RoCE/iWarp Reference Guide</a></p><p><a href='#'>Go to top</a></p>" `
            -IncludeTitle -IncludeDescription -IncludeFootnotes
    } else {
        Write-Host "    WARNING: No VLTi interfaces found to analyze." -ForegroundColor Yellow
        $VLTiUsedInterfacesEasyOut = @([PSCustomObject]@{Message = "No VLTi interfaces found"})
        AddTo-HtmlReport -Title "VLTi Interfaces" `
            -Data $VLTiUsedInterfacesEasyOut `
            -Description "No VLTi interfaces were found in the switch configurations." `
            -Footnotes "Check that the SDDC data matches the switch tech-support files." `
            -IncludeTitle -IncludeDescription -IncludeFootnotes
    }

    #endregion VLTi

    #-------------------------------------------------------------
    # vLAN INTERFACES
    #-------------------------------------------------------------
    $StoragevLANUsedInterfaces = @()
    if ($Storagevlans.Count -gt 0) {
        ForEach ($StoragevLAN in $Storagevlans){
            ForEach($Interface in $ShowRunningConfigs){
                    $StoragevLANUsedInterfaces += $Interface | ?{$_.Header -imatch "interface vlan"+$StoragevLAN}
            }
        }
    } else {
        Write-Host "    INFO: No storage VLANs found in SDDC data" -ForegroundColor Yellow
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
            'MTU9216'                                          = If ((Get-LineValue $StoragevLANUsedInterface.Lines '9216').value -ne '9216') {
                                                                    $newheader=$StoragevLANUsedInterface.header.split(" ")[-1]
                                                                    $newheader=$newheader.substring(0,($newheader | Select-String "\d").matches[0].index) + " " + $newheader.substring(($newheader | Select-String "\d").matches[0].index)
                                                                    (($ShowInterface | ? Header -match $newheader).lines | Select-String "MTU\s(\d*)\sbytes").matches.Groups[1].Value
                                                               } else {Get-LineValue $StoragevLANUsedInterface.Lines '9216'} 
        }
        $StoragevLANUsedInterfacesOut += Set-MissingNoteProperties $StoragevLANUsedInterfacesInfo
    }

    if ($StoragevLANUsedInterfacesOut.Count -gt 0) {
        $StoragevLANUsedInterfacesEasyOut = Convert-ToSwitchComparisonTable -Interfaces $StoragevLANUsedInterfacesOut | sort ShouldBe
        #$StoragevLANUsedInterfacesEasyOut | Show-WideTable -Title "vLAN Switch Port Comparison"
        # Add to HTML report output sections
        AddTo-HtmlReport -Title "Storage vLAN Interfaces" `
            -Data $StoragevLANUsedInterfacesEasyOut `
            -Description "" `
            -Footnotes "Highlighted in red or yellow if out of spec. <p><a href='$SwitchRefLink' target='_blank'>Ref: Switch Configurations - RoCE/iWarp Reference Guide</a></p><p><a href='#'>Go to top</a></p>" `
            -IncludeTitle -IncludeDescription -IncludeFootnotes
    } else {
        Write-Host "    WARNING: No storage vLAN interfaces found to analyze." -ForegroundColor Yellow
        $StoragevLANUsedInterfacesEasyOut = @([PSCustomObject]@{Message = "No storage vLAN interfaces found"})
        AddTo-HtmlReport -Title "Storage vLAN Interfaces" `
            -Data $StoragevLANUsedInterfacesEasyOut `
            -Description "No storage vLAN interfaces were found in the switch configurations." `
            -Footnotes "Check that the SDDC data matches the switch tech-support files." `
            -IncludeTitle -IncludeDescription -IncludeFootnotes
    }



    #endregion

    # Save report
    Save-HtmlReport
}
}