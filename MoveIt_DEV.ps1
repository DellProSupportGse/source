#requires -Version 5.1
#requires -RunAsAdministrator

<#
.SYNOPSIS
    Move It - Shared-Nothing Live Migration Utility

.DESCRIPTION
    GUI utility to:
      1. Configure source cluster nodes and destination Hyper-V cluster/standalone nodes for
         Shared-Nothing Live Migration using Kerberos constrained delegation.
      2. Discover VMs located on selected source nodes.
      3. Select VMs and choose a destination host for each VM.
      4. Detect vTPM-enabled VMs and copy the source host's required
         "Shielded VM Local Certificates" certificate material to the
         selected destination host before migration.
      5. Find management IPv4 by its default gateway, exclude clustered IPs,
         and put its /32 entry
         first (metric 1), preserving every existing migration address.
      6. Add source/destination peers to WinRM TrustedHosts on both sides.
      7. Live migrate each VM to the local source node, verify ownership,
         then remove its source clustered role using local commands.
      8. Perform local Move-VM with storage from the console/RDP session.
      9. Add successfully moved VMs to the destination cluster when the destination is clustered.
     10. Remove the temporary migration configuration when finished.

    IMPORTANT:
      - vTPM certificate material is COPIED to the destination. It is not
        deleted from the source.
      - Run interactively on a selected source node. Cross-cluster Move-VM
        runs locally after intra-cluster live migration to this node.
      - All Up source nodes participate in discovery/host setup. Constrained
        delegation is configured only on the local source node running Move It
        and the selected destination nodes.
      - Management migration list changes persist after Disable. Original provider
        data is saved under ProgramData\MoveIt on each modified host.
      - Delegation is all-to-all only among the local source node running Move It
        and the selected destination nodes, including self entries. Enable preserves
        existing entries. Disable removes matching entries even if they predated Enable.
#>
Function Invoke-MoveIt {
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ---------------------------------------------------------------------------
# Globals
# ---------------------------------------------------------------------------

$script:MoveItVersion = "v34"

$script:SourceClusterNodes = @()
$script:DestClusterNodes   = @()
$script:DiscoveredVMs      = @()
$script:DestinationCSVPaths = @()
$script:DestinationType = $null   # Cluster or Node
$script:DestinationTarget = $null
$script:DestinationVMSwitchesByNode = @{}
$script:TrustedHostsAdded   = @{}

# Azure Local can render native WinForms checkbox controls incorrectly.
# Destination-node and VM selection state is therefore maintained separately
# from the visual glyphs.  The UI uses custom-drawn/text checkboxes.
$script:DestinationNodeStates = [ordered]@{}
$script:DestinationNodeCheckBoxes = @{}

$script:RequiredModules = @(
    "FailoverClusters"
    "ActiveDirectory"
    "Hyper-V"
)

# ---------------------------------------------------------------------------
# Logging / UI helpers
# ---------------------------------------------------------------------------

function Write-GuiLog {
    param(
        [Parameter(Mandatory)]
        [string]$Message,

        [ValidateSet("Info","Success","Warning","Error")]
        [string]$Type = "Info"
    )

    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

    $Prefix = switch ($Type) {
        "Success" { "[ OK ]" }
        "Warning" { "[WARN]" }
        "Error"   { "[ERR ]" }
        default   { "[INFO]" }
    }

    $Line = "$Timestamp $Prefix $Message"

    # The PowerShell host/transcript is the authoritative activity log.
    # Keep logging out of the GUI to preserve screen space.
    Write-Host $Line
}

function Show-Message {
    param(
        [string]$Text,
        [string]$Title = "Move It $($script:MoveItVersion)",
        [ValidateSet("Information","Warning","Error")]
        [string]$Type = "Information"
    )

    [System.Windows.Forms.MessageBox]::Show(
        $Text,
        $Title,
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::$Type
    ) | Out-Null
}


function Get-EmbeddedLemurImage {
    $Base64 = @'
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
'@

    $Bytes = [Convert]::FromBase64String($Base64)
    $Stream = New-Object System.IO.MemoryStream(,$Bytes)
    $TempImage = $null

    try {
        $TempImage = [System.Drawing.Image]::FromStream($Stream)
        return New-Object System.Drawing.Bitmap($TempImage)
    }
    finally {
        if ($TempImage) { $TempImage.Dispose() }
        $Stream.Dispose()
    }
}

function Set-ActionButtonsEnabled {
    param([bool]$Enabled)

    foreach ($Button in @(
        $script:btnDetect,
        $script:btnRefresh,
        $script:btnValidate,
        $script:btnEnable,
        $script:btnDisable,
        $script:btnDiscoverVMs,
        $script:btnValidateMoves,
        $script:btnMoveVMs
    )) {
        if ($Button) {
            $Button.Enabled = $Enabled
        }
    }

    [System.Windows.Forms.Application]::DoEvents()
}

# ---------------------------------------------------------------------------
# Requirements / discovery
# ---------------------------------------------------------------------------

function Initialize-RequiredModules {
    Write-GuiLog "Checking required PowerShell modules..."

    $MissingModules = @()

    foreach ($Module in $script:RequiredModules) {
        if (-not (Get-Module -ListAvailable -Name $Module)) {
            $MissingModules += $Module
        }
    }

    if ($MissingModules.Count -gt 0) {
        $Message =
            "The following required PowerShell modules are missing:`r`n`r`n" +
            ($MissingModules -join "`r`n")

        Write-GuiLog $Message "Error"
        Show-Message -Text $Message -Title "Missing PowerShell Modules" -Type Error
        return $false
    }

    try {
        foreach ($Module in $script:RequiredModules) {
            Import-Module $Module -ErrorAction Stop
        }

        Write-GuiLog "Required modules loaded successfully." "Success"
        return $true
    }
    catch {
        Write-GuiLog "Failed loading required modules: $($_.Exception.Message)" "Error"
        return $false
    }
}

function Get-LocalClusterName {
    try {
        return (Get-Cluster -ErrorAction Stop).Name
    }
    catch {
        return $null
    }
}

function Get-DomainDNSName {
    try {
        return (Get-ADDomain -ErrorAction Stop).DNSRoot
    }
    catch {
        try {
            return (Get-CimInstance Win32_ComputerSystem -ErrorAction Stop).Domain
        }
        catch {
            return $null
        }
    }
}

function Get-NodeDisplayText {
    param([Parameter(Mandatory)]$Node)
    return "{0,-30} [{1}]" -f $Node.Name, $Node.State
}

function Set-DestinationNodeChecked {
    param(
        [Parameter(Mandatory)][string]$NodeName,
        [Parameter(Mandatory)][bool]$Checked
    )

    $script:DestinationNodeStates[$NodeName] = $Checked

    if ($script:DestinationNodeCheckBoxes.ContainsKey($NodeName)) {
        $Box = $script:DestinationNodeCheckBoxes[$NodeName]
        $Box.Text = if ($Checked) { [char]0x2713 } else { "" }
    }
}

function New-DestinationNodeRow {
    param(
        [Parameter(Mandatory)]$Node,
        [Parameter(Mandatory)][int]$Top
    )

    $NodeName = [string]$Node.Name
    $IsUp = ([string]$Node.State -eq "Up")

    $Row = New-Object System.Windows.Forms.Panel
    $Row.Location = New-Object System.Drawing.Point(4, $Top)
    $Row.Size = New-Object System.Drawing.Size(370, 32)
    $Row.Tag = $NodeName
    $Row.Cursor = [System.Windows.Forms.Cursors]::Hand

    # Draw the checkbox ourselves.  Do not use CheckBox/CheckedListBox because
    # some Azure Local builds render native checkboxes as an incorrect icon.
    $Box = New-Object System.Windows.Forms.Label
    $Box.Location = New-Object System.Drawing.Point(5, 5)
    $Box.Size = New-Object System.Drawing.Size(22, 22)
    $Box.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $Box.Font = New-Object System.Drawing.Font("Segoe UI Symbol", 12, [System.Drawing.FontStyle]::Bold)
    $Box.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $Box.Tag = $NodeName
    $Box.Cursor = [System.Windows.Forms.Cursors]::Hand
    $Row.Controls.Add($Box)

    $Label = New-Object System.Windows.Forms.Label
    $Label.Location = New-Object System.Drawing.Point(36, 3)
    $Label.Size = New-Object System.Drawing.Size(325, 25)
    $Label.Text = Get-NodeDisplayText -Node $Node
    $Label.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $Label.Font = New-Object System.Drawing.Font("Consolas", 9)
    $Label.Tag = $NodeName
    $Label.Cursor = [System.Windows.Forms.Cursors]::Hand
    $Row.Controls.Add($Label)

    $script:DestinationNodeCheckBoxes[$NodeName] = $Box
    Set-DestinationNodeChecked -NodeName $NodeName -Checked $IsUp

    $Toggle = {
        param($sender, $e)

        $Name = [string]$sender.Tag
        if ([string]::IsNullOrWhiteSpace($Name)) {
            return
        }

        $Current = $false
        if ($script:DestinationNodeStates.Contains($Name)) {
            $Current = [bool]$script:DestinationNodeStates[$Name]
        }

        Set-DestinationNodeChecked -NodeName $Name -Checked (-not $Current)
    }

    $Row.Add_Click($Toggle)
    $Box.Add_Click($Toggle)
    $Label.Add_Click($Toggle)

    return $Row
}

function Set-NodeList {
    param(
        [Parameter(Mandatory)]
        [System.Windows.Forms.Panel]$List,

        [Parameter(Mandatory)]
        [array]$Nodes
    )

    $List.SuspendLayout()

    try {
        $List.Controls.Clear()
        $script:DestinationNodeStates = [ordered]@{}
        $script:DestinationNodeCheckBoxes = @{}

        $Top = 4
        foreach ($Node in $Nodes) {
            $Row = New-DestinationNodeRow -Node $Node -Top $Top
            [void]$List.Controls.Add($Row)
            $Top += 34
        }
    }
    finally {
        $List.ResumeLayout()
        $List.Refresh()
    }
}

function Get-SelectedNodeNames {
    param(
        [Parameter(Mandatory)]
        [System.Windows.Forms.Panel]$List,

        [Parameter(Mandatory)]
        [array]$ClusterNodes
    )

    $SelectedNodes = @(
        foreach ($Node in $ClusterNodes) {
            $NodeName = [string]$Node.Name
            if (
                $script:DestinationNodeStates.Contains($NodeName) -and
                [bool]$script:DestinationNodeStates[$NodeName]
            ) {
                $NodeName
            }
        }
    )

    return @($SelectedNodes | Sort-Object -Unique)
}

function Select-AllUpNodes {
    param(
        [Parameter(Mandatory)]
        [System.Windows.Forms.Panel]$List,

        [Parameter(Mandatory)]
        [array]$ClusterNodes
    )

    foreach ($Node in $ClusterNodes) {
        $NodeName = [string]$Node.Name
        $IsUp = ([string]$Node.State -eq "Up")
        Set-DestinationNodeChecked -NodeName $NodeName -Checked $IsUp
    }

    $List.Refresh()
}

function Clear-NodeSelections {
    param(
        [Parameter(Mandatory)]
        [System.Windows.Forms.Panel]$List
    )

    foreach ($NodeName in @($script:DestinationNodeStates.Keys)) {
        Set-DestinationNodeChecked -NodeName ([string]$NodeName) -Checked $false
    }

    $List.Refresh()
}

function Refresh-Clusters {
    $SourceCluster = Get-LocalClusterName
    if ([string]::IsNullOrWhiteSpace($SourceCluster)) {
        Show-Message -Text 'Run Move It while logged into a source cluster node (console/RDP).' -Title 'Local Source Required' -Type Error
        return $null
    }

    $script:txtSourceCluster.Text = $SourceCluster
    $DestTarget = $script:txtDestCluster.Text.Trim()

    if ([string]::IsNullOrWhiteSpace($DestTarget)) {
        Show-Message -Text "Enter a destination cluster name or standalone Hyper-V node." -Type Warning
        return $false
    }

    if ($SourceCluster -ieq $DestTarget) {
        Show-Message -Text "The source cluster and destination cannot be the same." -Type Error
        return $false
    }

    Write-GuiLog "Querying source cluster: $SourceCluster"
    try {
        $script:SourceClusterNodes = @(Get-ClusterNode -Cluster $SourceCluster -ErrorAction Stop | Sort-Object Name)
        Write-GuiLog "Source cluster contains $($script:SourceClusterNodes.Count) node(s)." "Success"
    }
    catch {
        Write-GuiLog "Unable to query source cluster $SourceCluster : $($_.Exception.Message)" "Error"
        Show-Message -Text "Unable to query source cluster $SourceCluster.`r`n`r`n$($_.Exception.Message)" -Type Error
        return $false
    }

    # Destination may be a failover cluster name or a standalone Hyper-V host.
    $script:DestinationType = $null
    $script:DestinationTarget = $DestTarget
    $script:DestClusterNodes = @()
    $script:DestinationCSVPaths = @()

    Write-GuiLog "Detecting destination target: $DestTarget"

    try {
        $ClusterNodes = @(Get-ClusterNode -Cluster $DestTarget -ErrorAction Stop | Sort-Object Name)
        if ($ClusterNodes.Count -gt 0) {
            $script:DestinationType = 'Cluster'
            $script:DestClusterNodes = @($ClusterNodes)
            Write-GuiLog "Destination '$DestTarget' detected as a cluster with $($script:DestClusterNodes.Count) node(s)." "Success"

            $script:DestinationCSVPaths = @(Get-DestinationCSVPaths -ClusterName $DestTarget)
            if ($script:DestinationCSVPaths.Count -eq 0) {
                Show-Message -Text "No Cluster Shared Volumes were discovered on destination cluster $DestTarget.`r`n`r`nCluster-to-cluster VM moves require a destination CSV." -Title "No Destination CSV Found" -Type Error
                return $false
            }
        }
    }
    catch {
        Write-GuiLog "'$DestTarget' was not detected as a destination cluster. Checking for a standalone Hyper-V node..." "Info"
    }

    if (-not $script:DestinationType) {
        try {
            $NodeInfo = Invoke-Command -ComputerName $DestTarget -ErrorAction Stop -ScriptBlock {
                Import-Module Hyper-V -ErrorAction Stop
                $HostInfo = Get-VMHost -ErrorAction Stop
                [PSCustomObject]@{
                    ComputerName       = $env:COMPUTERNAME
                    VirtualMachinePath = [string]$HostInfo.VirtualMachinePath
                }
            }

            $NodeName = if ([string]::IsNullOrWhiteSpace([string]$NodeInfo.ComputerName)) { $DestTarget } else { [string]$NodeInfo.ComputerName }
            $script:DestinationType = 'Node'
            $script:DestClusterNodes = @([PSCustomObject]@{ Name = $NodeName; State = 'Up' })

            $StorageRoot = [string]$NodeInfo.VirtualMachinePath
            if ([string]::IsNullOrWhiteSpace($StorageRoot)) {
                $StorageRoot = 'C:\ProgramData\Microsoft\Windows\Hyper-V'
            }
            $script:DestinationCSVPaths = @($StorageRoot)

            Write-GuiLog "Destination '$DestTarget' detected as standalone Hyper-V node $NodeName." "Success"
            Write-GuiLog "Standalone destination VM storage root: $StorageRoot" "Success"
        }
        catch {
            Write-GuiLog "Unable to use destination '$DestTarget' as a cluster or standalone Hyper-V node: $($_.Exception.Message)" "Error"
            Show-Message -Text "Unable to use '$DestTarget' as either a destination cluster or a standalone Hyper-V node.`r`n`r`n$($_.Exception.Message)" -Title "Destination Not Found" -Type Error
            return $false
        }
    }

    Set-NodeList -List $script:lstDestNodes -Nodes $script:DestClusterNodes

    foreach ($Node in $script:SourceClusterNodes) {
        $Type = if ([string]$Node.State -eq "Up") { "Success" } else { "Warning" }
        Write-GuiLog "SOURCE $($Node.Name) : $($Node.State)" $Type
    }
    foreach ($Node in $script:DestClusterNodes) {
        $Type = if ([string]$Node.State -eq "Up") { "Success" } else { "Warning" }
        $Prefix = if ($script:DestinationType -eq 'Cluster') { 'DEST CLUSTER' } else { 'DEST NODE   ' }
        Write-GuiLog "$Prefix $($Node.Name) : $($Node.State)" $Type
    }

    $script:DiscoveredVMs = @()
    if ($script:gridVMs) { $script:gridVMs.Rows.Clear() }

    Write-GuiLog "Destination discovery completed. Type=$($script:DestinationType). All available destination nodes were selected automatically." "Success"
    return $true
}


# ---------------------------------------------------------------------------
# Node validation
# ---------------------------------------------------------------------------

function Test-SelectedNodesReady {
    param(
        [Parameter(Mandatory)]
        [string]$ClusterName,

        [Parameter(Mandatory)]
        [string[]]$SelectedNodes,

        [Parameter(Mandatory)]
        [string]$Label
    )

    Write-GuiLog "Re-validating selected $Label nodes..."

    try {
        $CurrentNodes = @(Get-ClusterNode -Cluster $ClusterName -ErrorAction Stop)
    }
    catch {
        Write-GuiLog "Unable to re-query $Label cluster: $($_.Exception.Message)" "Error"
        return $false
    }

    $Problems = @()

    foreach ($NodeName in $SelectedNodes) {
        $Node =
            $CurrentNodes |
            Where-Object { $_.Name -ieq $NodeName } |
            Select-Object -First 1

        if (-not $Node) {
            $Problems += "$NodeName = Not found in cluster"
            Write-GuiLog "$Label node $NodeName was not found." "Error"
            continue
        }

        # State=Paused also catches the normal cluster pause/drain condition.
        if ([string]$Node.State -ne "Up") {
            $Problems += "$NodeName = $($Node.State)"
            Write-GuiLog "$Label node $NodeName is $($Node.State)." "Error"
        }
        else {
            Write-GuiLog "$Label node $NodeName is Up." "Success"
        }
    }

    if ($Problems.Count -gt 0) {
        $Message =
            "The following selected $Label nodes are not ready:`r`n`r`n" +
            ($Problems -join "`r`n") +
            "`r`n`r`nSelected nodes must be Up and not Paused."

        Show-Message -Text $Message -Title "Cluster Node Validation Failed" -Type Error
        return $false
    }

    Write-GuiLog "All selected $Label nodes are Up." "Success"
    return $true
}

function Test-StandaloneDestinationNodeReady {
    param([Parameter(Mandatory)][string]$NodeName)

    Write-GuiLog "Re-validating standalone destination node $NodeName..."
    try {
        $Result = Invoke-Command -ComputerName $NodeName -ErrorAction Stop -ScriptBlock {
            Import-Module Hyper-V -ErrorAction Stop
            [bool](Get-VMHost -ErrorAction Stop)
        }
        if (-not $Result) { throw 'Get-VMHost did not return a valid Hyper-V host.' }
        Write-GuiLog "Standalone destination node $NodeName is reachable and Hyper-V is available." "Success"
        return $true
    }
    catch {
        Write-GuiLog "Standalone destination node $NodeName failed validation: $($_.Exception.Message)" "Error"
        Show-Message -Text "Standalone destination node '$NodeName' is not ready.`r`n`r`n$($_.Exception.Message)" -Title "Destination Node Validation Failed" -Type Error
        return $false
    }
}

function Get-ValidatedSelections {
    $SourceCluster = Get-LocalClusterName
    if ([string]::IsNullOrWhiteSpace($SourceCluster)) {
        Show-Message -Text 'Run Move It while logged into a source cluster node (console/RDP).' -Title 'Local Source Required' -Type Error
        return $null
    }
    $script:txtSourceCluster.Text = $SourceCluster

    $DestTarget = $script:txtDestCluster.Text.Trim()
    $DomainName = $script:txtDomain.Text.Trim()

    if ($script:SourceClusterNodes.Count -eq 0 -or $script:DestClusterNodes.Count -eq 0 -or [string]::IsNullOrWhiteSpace($script:DestinationType) -or $script:DestinationTarget -ine $DestTarget) {
        Write-GuiLog "Source/destination nodes have not been discovered for the current target. Refreshing..."
        if (-not (Refresh-Clusters)) { return $null }
    }

    $SelectedSourceNodes = @($script:SourceClusterNodes | Where-Object { [string]$_.State -eq 'Up' } | ForEach-Object { [string]$_.Name })
    $SelectedDestNodes = @(Get-SelectedNodeNames -List $script:lstDestNodes -ClusterNodes $script:DestClusterNodes)

    if ($SelectedSourceNodes.Count -eq 0) {
        Show-Message -Text "No Up nodes were found in the local source cluster." -Type Warning
        return $null
    }
    if ($SelectedDestNodes.Count -eq 0) {
        Show-Message -Text "Select at least one destination node." -Type Warning
        return $null
    }
    if ([string]::IsNullOrWhiteSpace($DomainName)) {
        Show-Message -Text "The Active Directory DNS domain is required." -Type Warning
        return $null
    }

    if (-not (Test-SelectedNodesReady -ClusterName $SourceCluster -SelectedNodes $SelectedSourceNodes -Label "source")) { return $null }

    if ($script:DestinationType -eq 'Cluster') {
        if (-not (Test-SelectedNodesReady -ClusterName $DestTarget -SelectedNodes $SelectedDestNodes -Label "destination")) { return $null }
    }
    else {
        foreach ($NodeName in $SelectedDestNodes) {
            if (-not (Test-StandaloneDestinationNodeReady -NodeName $NodeName)) { return $null }
        }
    }

    return [PSCustomObject]@{
        SourceCluster   = $SourceCluster
        DestTarget      = $DestTarget
        DestType        = $script:DestinationType
        DestCluster     = if ($script:DestinationType -eq 'Cluster') { $DestTarget } else { $null }
        DomainName      = $DomainName
        LocalSourceNode = $env:COMPUTERNAME
        SourceNodes     = @($SelectedSourceNodes)
        DestNodes       = @($SelectedDestNodes)
    }
}

# ---------------------------------------------------------------------------
# Kerberos constrained delegation
# ---------------------------------------------------------------------------

function Get-MigrationDelegationSPNs {
    param(
        [Parameter(Mandatory)]
        [string]$RemoteNode,

        [Parameter(Mandatory)]
        [string]$DomainName
    )

    return @(
        "Microsoft Virtual System Migration Service/$RemoteNode.$DomainName"
        "cifs/$RemoteNode.$DomainName"
        "Microsoft Virtual System Migration Service/$RemoteNode"
        "cifs/$RemoteNode"
    )
}

function Add-MigrationDelegation {
    param(
        [Parameter(Mandatory)]
        [string]$Node,

        [Parameter(Mandatory)]
        [string[]]$RemoteNodes,

        [Parameter(Mandatory)]
        [string]$DomainName
    )

    Write-GuiLog "Reading constrained delegation for $Node..."

    try {
        $Computer =
            Get-ADComputer `
                -Identity $Node `
                -Properties "msDS-AllowedToDelegateTo" `
                -ErrorAction Stop

        $Current = @($Computer.'msDS-AllowedToDelegateTo')
    }
    catch {
        Write-GuiLog "Unable to read AD computer $Node : $($_.Exception.Message)" "Error"
        return $false
    }

    $Success = $true

    foreach ($RemoteNode in $RemoteNodes) {
        foreach ($SPN in (Get-MigrationDelegationSPNs -RemoteNode $RemoteNode -DomainName $DomainName)) {
            $Existing =
                $Current |
                Where-Object { $_ -ieq $SPN } |
                Select-Object -First 1

            if ($Existing) {
                Write-GuiLog "$Node already contains delegation: $SPN"
                continue
            }

            try {
                Set-ADComputer `
                    -Identity $Node `
                    -Add @{ "msDS-AllowedToDelegateTo" = $SPN } `
                    -ErrorAction Stop

                $Current += $SPN
                Write-GuiLog "$Node added delegation: $SPN" "Success"
            }
            catch {
                Write-GuiLog "$Node FAILED adding delegation $SPN : $($_.Exception.Message)" "Error"
                $Success = $false
            }
        }
    }

    return $Success
}

function Remove-MigrationDelegation {
    param(
        [Parameter(Mandatory)]
        [string]$Node,

        [Parameter(Mandatory)]
        [string[]]$RemoteNodes,

        [Parameter(Mandatory)]
        [string]$DomainName
    )

    Write-GuiLog "Reading constrained delegation for $Node..."

    try {
        $Computer =
            Get-ADComputer `
                -Identity $Node `
                -Properties "msDS-AllowedToDelegateTo" `
                -ErrorAction Stop

        $Current = @($Computer.'msDS-AllowedToDelegateTo')
    }
    catch {
        Write-GuiLog "Unable to read AD computer $Node : $($_.Exception.Message)" "Error"
        return $false
    }

    $Success = $true

    foreach ($RemoteNode in $RemoteNodes) {
        foreach ($SPN in (Get-MigrationDelegationSPNs -RemoteNode $RemoteNode -DomainName $DomainName)) {
            $Existing =
                $Current |
                Where-Object { $_ -ieq $SPN } |
                Select-Object -First 1

            if (-not $Existing) {
                Write-GuiLog "$Node delegation not present: $SPN"
                continue
            }

            try {
                Set-ADComputer `
                    -Identity $Node `
                    -Remove @{ "msDS-AllowedToDelegateTo" = $Existing } `
                    -ErrorAction Stop

                $Current = @($Current | Where-Object { $_ -ine $Existing })
                Write-GuiLog "$Node removed delegation: $Existing" "Success"
            }
            catch {
                Write-GuiLog "$Node FAILED removing delegation $SPN : $($_.Exception.Message)" "Error"
                $Success = $false
            }
        }
    }

    return $Success
}

# ---------------------------------------------------------------------------
# WinRM TrustedHosts helpers
# ---------------------------------------------------------------------------

function Add-MigrationTrustedHosts {
    param(
        [Parameter(Mandatory)][string]$Node,
        [Parameter(Mandatory)][string[]]$RemoteNodes,
        [Parameter(Mandatory)][string]$DomainName
    )

    $PeerNames = @(
        foreach ($RemoteNode in $RemoteNodes) {
            if ($Node -ieq $RemoteNode) { continue }
            $RemoteNode
            "$RemoteNode.$DomainName"
        }
    ) | Sort-Object -Unique

    if ($PeerNames.Count -eq 0) { return $true }

    Write-GuiLog "Adding WinRM TrustedHosts on $Node for: $($PeerNames -join ', ')"

    try {
        $Result = Invoke-Command -ComputerName $Node -ErrorAction Stop -ScriptBlock {
            param([string[]]$Peers)

            $Path = 'WSMan:\localhost\Client\TrustedHosts'
            $CurrentValue = [string](Get-Item -Path $Path -ErrorAction Stop).Value
            $Current = @(
                $CurrentValue -split ',' |
                ForEach-Object { $_.Trim() } |
                Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
            )

            $Added = @()
            foreach ($Peer in $Peers) {
                if (-not ($Current | Where-Object { $_ -ieq $Peer })) {
                    $Current += $Peer
                    $Added += $Peer
                }
            }

            if ($Added.Count -gt 0) {
                Set-Item -Path $Path -Value ($Current -join ',') -Force -ErrorAction Stop
            }

            [PSCustomObject]@{
                Added = @($Added)
                Value = [string](Get-Item -Path $Path -ErrorAction Stop).Value
            }
        } -ArgumentList (,$PeerNames)

        $AddedHere = @($Result.Added)
        if ($AddedHere.Count -gt 0) {
            if (-not $script:TrustedHostsAdded.ContainsKey($Node)) {
                $script:TrustedHostsAdded[$Node] = @()
            }
            $script:TrustedHostsAdded[$Node] = @(
                $script:TrustedHostsAdded[$Node] + $AddedHere | Sort-Object -Unique
            )
            Write-GuiLog "$Node TrustedHosts added: $($AddedHere -join ', ')" "Success"
        }
        else {
            Write-GuiLog "$Node already trusted all required migration peers." "Success"
        }

        return $true
    }
    catch {
        Write-GuiLog "$Node TrustedHosts configuration FAILED: $($_.Exception.Message)" "Error"
        return $false
    }
}

function Remove-MigrationTrustedHosts {
    param([Parameter(Mandatory)][string]$Node)

    if (-not $script:TrustedHostsAdded.ContainsKey($Node)) {
        Write-GuiLog "$Node has no TrustedHosts entries added by this session; leaving TrustedHosts unchanged."
        return $true
    }

    $EntriesToRemove = @($script:TrustedHostsAdded[$Node])
    if ($EntriesToRemove.Count -eq 0) { return $true }

    Write-GuiLog "Removing session-added WinRM TrustedHosts entries from ${Node}: $($EntriesToRemove -join ', ')"

    try {
        Invoke-Command -ComputerName $Node -ErrorAction Stop -ScriptBlock {
            param([string[]]$Entries)

            $Path = 'WSMan:\localhost\Client\TrustedHosts'
            $CurrentValue = [string](Get-Item -Path $Path -ErrorAction Stop).Value
            $Current = @(
                $CurrentValue -split ',' |
                ForEach-Object { $_.Trim() } |
                Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
            )

            $Remaining = @(
                $Current |
                Where-Object {
                    $Item = $_
                    -not ($Entries | Where-Object { $_ -ieq $Item })
                }
            )

            Set-Item -Path $Path -Value ($Remaining -join ',') -Force -ErrorAction Stop
        } -ArgumentList (,$EntriesToRemove)

        $script:TrustedHostsAdded.Remove($Node)
        Write-GuiLog "$Node session-added TrustedHosts entries removed." "Success"
        return $true
    }
    catch {
        Write-GuiLog "$Node TrustedHosts cleanup FAILED: $($_.Exception.Message)" "Error"
        return $false
    }
}

# ---------------------------------------------------------------------------
# Host setup / cleanup
# ---------------------------------------------------------------------------

# Runs locally on each selected host through Invoke-Command.
$script:ManagementMigrationNetwork = {
    param([bool]$Configure = $false)

    $ErrorActionPreference = 'Stop'

    Import-Module Hyper-V -ErrorAction Stop

    # A destination may be either a cluster node or a standalone Hyper-V host.
    # Only query Failover Clustering when this computer is actually participating
    # in a running cluster. A standalone Hyper-V host may have Failover Clustering
    # installed while ClusSvc is stopped, which must not block shared-nothing moves.
    $ClusterIPs = @()
    $ClusterService = Get-Service -Name ClusSvc -ErrorAction SilentlyContinue

    if ($ClusterService -and $ClusterService.Status -eq 'Running') {
        try {
            Import-Module FailoverClusters -ErrorAction Stop
            $null = Get-Cluster -ErrorAction Stop

            # A single-node Azure Local cluster can own the cluster IP on the
            # same adapter as the node management IP.  Both can therefore appear
            # in Get-NetIPConfiguration as addresses on an interface that has a
            # default gateway.  Explicitly collect clustered IPv4 addresses and
            # subtract them before choosing the node management address.
            #
            # Read each resource independently.  One unusual/non-IPv4 IP resource
            # must not cause the entire exclusion list to be discarded.
            $ClusterIPResources = @(
                Get-ClusterResource -ErrorAction Stop |
                    Where-Object {
                        $_.Name -eq 'Cluster IP Address' -or
                        [string]$_.ResourceType -eq 'IP Address'
                    }
            )

            $DiscoveredClusterIPs = @()
            foreach ($ClusterIPResource in $ClusterIPResources) {
                try {
                    $AddressParameter = $ClusterIPResource |
                        Get-ClusterParameter -Name Address -ErrorAction Stop

                    $Address = [string]$AddressParameter.Value
                    $ParsedAddress = $null
                    if (
                        -not [string]::IsNullOrWhiteSpace($Address) -and
                        [System.Net.IPAddress]::TryParse($Address, [ref]$ParsedAddress) -and
                        $ParsedAddress.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetwork
                    ) {
                        $DiscoveredClusterIPs += $Address
                    }
                }
                catch {
                    # Ignore only this resource.  Do not throw away cluster IPs
                    # successfully discovered from other resources.
                }
            }

            $ClusterIPs = @($DiscoveredClusterIPs | Sort-Object -Unique)
        }
        catch {
            # If this computer is not actually a usable cluster member, continue
            # as a standalone Hyper-V host.
            $ClusterIPs = @()
        }
    }

    $ManagementConfigs = @(
        Get-NetIPConfiguration |
            Where-Object { $null -ne $_.IPv4DefaultGateway }
    )

    $ManagementIPs = @(
        $ManagementConfigs |
            ForEach-Object { $_.IPv4Address.IPAddress } |
            Where-Object {
                -not [string]::IsNullOrWhiteSpace($_) -and
                [string]$_ -notin $ClusterIPs
            } |
            Sort-Object -Unique
    )

    if ($ManagementIPs.Count -ne 1) {
        throw "Expected one node IPv4 address with a default gateway after excluding clustered IPs [$($ClusterIPs -join ', ')]; found $($ManagementIPs.Count): $($ManagementIPs -join ', ')."
    }

    $ManagementIP = [string]$ManagementIPs[0]
    $GatewayIPs = @(
        $ManagementConfigs |
            ForEach-Object { $_.IPv4Address.IPAddress } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            Sort-Object -Unique
    )
    $SelectedManagementConfigs = @(
        $ManagementConfigs |
            Where-Object { @($_.IPv4Address.IPAddress) -contains $ManagementIP }
    )
    $NetworkAction = 'VerifyOnly'
    $Namespace = 'root/virtualization/v2'
    $Class = 'Msvm_VirtualSystemMigrationNetworkSettingData'

    if ($Configure) {
        # Back up the raw Hyper-V provider entries before changing metrics.
        $BackupDir = Join-Path $env:ProgramData 'MoveIt'
        New-Item -Path $BackupDir -ItemType Directory -Force | Out-Null
        $BackupPath = Join-Path $BackupDir ('MigrationNetworks-{0}.xml' -f (Get-Date -Format 'yyyyMMdd-HHmmss-fff'))

        $Networks = @(Get-WmiObject -Namespace $Namespace -Class $Class -ErrorAction Stop)

        [PSCustomObject]@{
            ComputerName              = $env:COMPUTERNAME
            UseAnyNetworkForMigration = (Get-VMHost).UseAnyNetworkForMigration
            Networks                  = @($Networks | ForEach-Object { $_.GetText(1) })
        } | Export-Clixml -LiteralPath $BackupPath

        # Manage ONLY the node management /32 with the Hyper-V cmdlets.
        # Azure Local owns the Microsoft:ClusterManaged entries; do not rewrite
        # those objects through WMI/VMMS because their Tags are provider-owned.
        $ManagementSubnet = "$ManagementIP/32"

        $ManagementNetwork = @(
            Get-VMMigrationNetwork -ErrorAction Stop |
                Where-Object {
                    ([string]$_.Subnet -eq $ManagementSubnet) -or
                    ([string]$_.Subnet -eq $ManagementIP)
                } |
                Select-Object -First 1
        )

        if ($ManagementNetwork.Count -eq 0) {
            Add-VMMigrationNetwork `
                -Subnet $ManagementSubnet `
                -Priority 1 `
                -ErrorAction Stop
            $NetworkAction = "Added $ManagementSubnet priority 1"
        }
        elseif ([uint32]$ManagementNetwork[0].Priority -ne 1) {
            $OldPriority = [uint32]$ManagementNetwork[0].Priority
            Set-VMMigrationNetwork `
                -Subnet ([string]$ManagementNetwork[0].Subnet) `
                -NewPriority 1 `
                -ErrorAction Stop
            $NetworkAction = "Changed $ManagementSubnet priority $OldPriority -> 1"
        }
        else {
            $NetworkAction = "$ManagementSubnet already priority 1"
        }

        # Leave every other migration network untouched. On Azure Local these
        # commonly carry Microsoft:ClusterManaged and a metric/priority such as
        # 5000. Move It must not retag, renumber, remove, or recreate them.

        # Confirm none of the entries present before configuration were lost.
        $After = @(Get-CimInstance -Namespace $Namespace -ClassName $Class -ErrorAction Stop)
        foreach ($Original in $Networks) {
            if (-not @($After | Where-Object {
                $_.InstanceID -eq $Original.InstanceID -and
                $_.SubnetNumber -eq $Original.SubnetNumber -and
                [int]$_.PrefixLength -eq [int]$Original.PrefixLength
            }).Count) {
                throw "Existing migration entry was not preserved: $($Original.SubnetNumber)/$($Original.PrefixLength). Backup: $BackupPath"
            }
        }

        Set-VMHost -UseAnyNetworkForMigration $false -ErrorAction Stop
    }
    else {
        $BackupPath = $null
    }

    # Verify directly against the Hyper-V provider. This intentionally avoids
    # Get/Set/Remove-VMMigrationNetwork for verification so provider-visible
    # entries that are not addressable by the cmdlets do not create false failures.
    $Verified = @(
        Get-CimInstance -Namespace $Namespace -ClassName $Class -ErrorAction Stop |
            Sort-Object Metric, InstanceID
    )

    $ManagementFirst = @(
        $Verified |
            Where-Object {
                $_.SubnetNumber -eq $ManagementIP -and
                [int]$_.PrefixLength -eq 32 -and
                [uint32]$_.Metric -eq 1
            }
    )

    $OtherFirst = @(
        $Verified |
            Where-Object {
                ($_.SubnetNumber -ne $ManagementIP -or [int]$_.PrefixLength -ne 32) -and
                [uint32]$_.Metric -le 1
            }
    )

    $UseAnyNetwork = [bool](Get-VMHost).UseAnyNetworkForMigration
    if ($UseAnyNetwork -or $ManagementFirst.Count -eq 0 -or $OtherFirst.Count -gt 0) {
        $VerifiedSummary = @(
            $Verified |
                ForEach-Object {
                    "{0}/{1}:metric={2}:tags={3}" -f $_.SubnetNumber,$_.PrefixLength,$_.Metric,(@($_.Tags) -join ',')
                }
        ) -join '; '

        throw "Management-first migration verification failed. ManagementIP=$ManagementIP; ClusterIPs=[$($ClusterIPs -join ', ')]; GatewayIPs=[$($GatewayIPs -join ', ')]; UseAnyNetwork=$UseAnyNetwork; ManagementMatches=$($ManagementFirst.Count); OtherPriorityLE1=$($OtherFirst.Count); Provider=[$VerifiedSummary]"
    }

    [PSCustomObject]@{
        ManagementIP       = $ManagementIP
        GatewayIPs         = ($GatewayIPs -join ', ')
        ExcludedClusterIPs = ($ClusterIPs -join ', ')
        InterfaceAlias     = ($SelectedManagementConfigs.InterfaceAlias -join ', ')
        Priority           = [uint32]$ManagementFirst[0].Metric
        NetworkAction      = $NetworkAction
        BackupPath         = $BackupPath
    }
}

function Enable-MigrationHost {
    param([Parameter(Mandatory)][string]$Node)

    Write-GuiLog "Configuring host $Node for Kerberos Live Migration..."

    try {
        $Result =
            Invoke-Command `
                -ComputerName $Node `
                -ErrorAction Stop `
                -ScriptBlock {

                    $Service = Get-Service -Name NetworkATC -ErrorAction SilentlyContinue

                    if ($Service) {
                        Set-Service -Name NetworkATC -StartupType Disabled -ErrorAction Stop
                        Stop-Service -Name NetworkATC -Force -ErrorAction SilentlyContinue
                    }

                    # Use the explicit migration list, with management first.
                    Set-VMHost `
                        -VirtualMachineMigrationAuthenticationType Kerberos `
                        -UseAnyNetworkForMigration $false `
                        -ErrorAction Stop

                    $VMHost = Get-VMHost

                    [PSCustomObject]@{
                        ComputerName             = $env:COMPUTERNAME
                        LiveMigrationAuth        = [string]$VMHost.VirtualMachineMigrationAuthenticationType
                        UseAnyNetworkForMigration = [bool]$VMHost.UseAnyNetworkForMigration
                        NetworkATCStatus         = if ($Service) { [string](Get-Service NetworkATC).Status } else { "Service not found" }
                    }
                }

        $NetworkResult = Invoke-Command -ComputerName $Node -ScriptBlock $script:ManagementMigrationNetwork -ArgumentList $true -ErrorAction Stop
        Write-GuiLog "$Node gateway IPv4 candidates: $($NetworkResult.GatewayIPs)"
        Write-GuiLog "$Node excluded clustered IPs: $($NetworkResult.ExcludedClusterIPs)"
        Write-GuiLog "$Node selected management IP: $($NetworkResult.ManagementIP)"
        Write-GuiLog "$Node migration network action: $($NetworkResult.NetworkAction)" "Success"
        Write-GuiLog "$Node management migration entry: $($NetworkResult.ManagementIP)/32, priority=$($NetworkResult.Priority), adapter=$($NetworkResult.InterfaceAlias). Backup=$($NetworkResult.BackupPath)" "Success"
        Write-GuiLog "$Node configured. Auth=$($Result.LiveMigrationAuth), UseAnyMigrationNetwork=$($Result.UseAnyNetworkForMigration), NetworkATC=$($Result.NetworkATCStatus)" "Success"
        return $true
    }
    catch {
        Write-GuiLog "$Node host configuration FAILED: $($_.Exception.Message)" "Error"
        return $false
    }
}

function Disable-MigrationHost {
    param([Parameter(Mandatory)][string]$Node)

    Write-GuiLog "Returning host $Node to CredSSP / NetworkATC enabled..."

    try {
        $Result =
            Invoke-Command `
                -ComputerName $Node `
                -ErrorAction Stop `
                -ScriptBlock {

                    Set-VMHost `
                        -VirtualMachineMigrationAuthenticationType CredSSP `
                        -UseAnyNetworkForMigration $false `
                        -ErrorAction Stop

                    $Service = Get-Service -Name NetworkATC -ErrorAction SilentlyContinue

                    if ($Service) {
                        Set-Service -Name NetworkATC -StartupType Automatic -ErrorAction Stop
                        Start-Service -Name NetworkATC -ErrorAction Stop
                    }

                    $VMHost = Get-VMHost

                    [PSCustomObject]@{
                        ComputerName             = $env:COMPUTERNAME
                        LiveMigrationAuth        = [string]$VMHost.VirtualMachineMigrationAuthenticationType
                        UseAnyNetworkForMigration = [bool]$VMHost.UseAnyNetworkForMigration
                        NetworkATCStatus         = if ($Service) { [string](Get-Service NetworkATC).Status } else { "Service not found" }
                    }
                }

        Write-GuiLog "$Node restored. Auth=$($Result.LiveMigrationAuth), UseAnyMigrationNetwork=$($Result.UseAnyNetworkForMigration), NetworkATC=$($Result.NetworkATCStatus)" "Success"
        return $true
    }
    catch {
        Write-GuiLog "$Node cleanup FAILED: $($_.Exception.Message)" "Error"
        return $false
    }
}

function Test-CurrentSelection {
    $Selection = Get-ValidatedSelections

    if (-not $Selection) {
        return
    }

    Write-GuiLog "========================================"
    Write-GuiLog "Selected node validation PASSED." "Success"
    Write-GuiLog "SOURCE: $($Selection.SourceNodes -join ', ')"
    Write-GuiLog "DESTINATION: $($Selection.DestNodes -join ', ')"
    Write-GuiLog "DOMAIN: $($Selection.DomainName)"
    Write-GuiLog "========================================"

    Show-Message -Text "All selected nodes are Up and ready." -Title "Validation Passed" -Type Information
}

function Enable-SharedNothingMigration {
    $Selection = Get-ValidatedSelections

    if (-not $Selection) {
        return
    }

    $SourceList = $Selection.SourceNodes -join ", "
    $DestList   = $Selection.DestNodes -join ", "
    $SetupNodes = @(@($Selection.LocalSourceNode) + @($Selection.DestNodes)) | Sort-Object -Unique

    $ConfirmationText = @"
Enable Move It using the following configuration?

Source Migration Host (configured): $($Selection.LocalSourceNode)
Source Cluster Nodes (discovery only): $SourceList

Destination Nodes (configured):
$DestList

This action will ONLY modify the local source node running Move It and the selected destination node(s):

- Disable Network ATC
- Configure Live Migration to use Kerberos authentication
- Disable the management network for migration traffic and exclude cluster IP addresses
- Prioritize the management IP while preserving existing migration IPs
- Back up the current migration network configuration
- Configure constrained delegation for CIFS and Live Migration
- Update WinRM TrustedHosts

Other source-cluster nodes are used only for VM discovery/staging and are NOT modified by Enable Move It.

Continue?
"@

    $Response =
        [System.Windows.Forms.MessageBox]::Show(
            $ConfirmationText,
            "Enable Move It",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )

    if ($Response -ne [System.Windows.Forms.DialogResult]::Yes) {
        Write-GuiLog "Enable operation cancelled by user." "Warning"
        return
    }

    Set-ActionButtonsEnabled -Enabled $false

    try {
        $HostErrors    = 0
        $ADErrors      = 0
        $WinRMErrors   = 0

        # Enable/prepare ONLY the interactive local source host and the
        # selected destination host(s). Other source-cluster nodes are used
        # for VM discovery/staging only and must not be modified here.
        if (-not (Add-MigrationTrustedHosts -Node $Selection.LocalSourceNode -RemoteNodes $Selection.DestNodes -DomainName $Selection.DomainName)) {
            $WinRMErrors++
        }

        foreach ($Node in $Selection.DestNodes) {
            if (-not (Add-MigrationTrustedHosts -Node $Node -RemoteNodes @($Selection.LocalSourceNode) -DomainName $Selection.DomainName)) {
                $WinRMErrors++
            }
        }

        foreach ($Node in $SetupNodes) {
            if (-not (Enable-MigrationHost -Node $Node)) {
                $HostErrors++
            }
        }

        $DelegationNodes = @($SetupNodes)

        foreach ($Node in $DelegationNodes) {
            if (-not (
                Add-MigrationDelegation `
                    -Node $Node `
                    -RemoteNodes $DelegationNodes `
                    -DomainName $Selection.DomainName
            )) {
                $ADErrors++
            }
        }

        if ($HostErrors -eq 0 -and $ADErrors -eq 0 -and $WinRMErrors -eq 0) {
            Write-GuiLog "ENABLE completed successfully." "Success"
            Show-Message -Text "'Move It' was enabled successfully for the selected nodes." -Title "Move It Enabled"
        }
        else {
            Write-GuiLog "ENABLE completed with errors. Host errors=$HostErrors AD errors=$ADErrors WinRM errors=$WinRMErrors" "Warning"
            Show-Message -Text "Enable completed with one or more errors. Review the PowerShell window or transcript log." -Type Warning
        }
    }
    finally {
        Set-ActionButtonsEnabled -Enabled $true
    }
}

function Disable-SharedNothingMigration {
    $Selection = Get-ValidatedSelections

    if (-not $Selection) {
        return
    }

    $SourceList = $Selection.SourceNodes -join ", "
    $DestList   = $Selection.DestNodes -join ", "
    $SetupNodes = @(@($Selection.LocalSourceNode) + @($Selection.DestNodes)) | Sort-Object -Unique

    $ConfirmationText = @"
DISABLE 'Move It' configuration?

LOCAL MIGRATION SOURCE (configured): $($Selection.LocalSourceNode)
SOURCE CLUSTER NODES (discovery only):
$SourceList

DESTINATION NODES (configured):
$DestList

This will ONLY modify the local source node running Move It and the selected destination node(s):
  - Remove migration delegation from those nodes
  - Remove WinRM TrustedHosts entries added by this session
  - Set Live Migration authentication back to CredSSP
  - Enable and start NetworkATC
  - Management migration entries are retained; NetworkATC may reapply its policy

Other source-cluster nodes are not modified by Disable Move It.

Continue?
"@

    $Response =
        [System.Windows.Forms.MessageBox]::Show(
            $ConfirmationText,
            "Disable Move It",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )

    if ($Response -ne [System.Windows.Forms.DialogResult]::Yes) {
        Write-GuiLog "Disable operation cancelled by user." "Warning"
        return
    }

    Set-ActionButtonsEnabled -Enabled $false

    try {
        $HostErrors  = 0
        $ADErrors    = 0
        $WinRMErrors = 0

        $DelegationNodes =
            @(@($env:COMPUTERNAME) + @($Selection.DestNodes)) |
            Sort-Object -Unique

        # Match Enable: remove delegation only from the local source and the
        # selected destination nodes.
        foreach ($Node in $DelegationNodes) {
            if (-not (
                Remove-MigrationDelegation `
                    -Node $Node `
                    -RemoteNodes $DelegationNodes `
                    -DomainName $Selection.DomainName
            )) {
                $ADErrors++
            }
        }

        foreach ($Node in $SetupNodes) {
            if (-not (Remove-MigrationTrustedHosts -Node $Node)) {
                $WinRMErrors++
            }
        }

        foreach ($Node in $SetupNodes) {
            if (-not (Disable-MigrationHost -Node $Node)) {
                $HostErrors++
            }
        }

        if ($HostErrors -eq 0 -and $ADErrors -eq 0 -and $WinRMErrors -eq 0) {
            Write-GuiLog "DISABLE completed successfully." "Success"
            Show-Message -Text "'Move It' was disabled successfully for the selected nodes." -Title "Move It Disabled"
        }
        else {
            Write-GuiLog "DISABLE completed with errors. Host errors=$HostErrors AD errors=$ADErrors WinRM errors=$WinRMErrors" "Warning"
            Show-Message -Text "Disable completed with one or more errors. Review the PowerShell window or transcript log." -Type Warning
        }
    }
    finally {
        Set-ActionButtonsEnabled -Enabled $true
    }
}


function Get-DestinationCSVPaths {
    param(
        [Parameter(Mandatory)]
        [string]$ClusterName
    )

    Write-GuiLog "Discovering Cluster Shared Volumes on destination cluster $ClusterName..."

    try {
        # Use the destination cluster CSV FriendlyVolumeName values directly.
        # Example:
        # (Get-ClusterSharedVolume -Cluster AZL760AGNClus28).SharedVolumeInfo.FriendlyVolumeName
        $CSVs =
            @(
                (Get-ClusterSharedVolume `
                    -Cluster $ClusterName `
                    -ErrorAction Stop
                ).SharedVolumeInfo.FriendlyVolumeName |
                Where-Object {
                    -not [string]::IsNullOrWhiteSpace($_)
                } |
                Sort-Object -Unique |
                Sort-Object `
                    @{ Expression = { if ((Split-Path $_ -Leaf) -like 'Infrastructure_*') { 1 } else { 0 } } }, `
                    @{ Expression = { $_ } }
            )

        if ($CSVs.Count -eq 0) {
            Write-GuiLog "No Cluster Shared Volumes were found on $ClusterName." "Error"
            return @()
        }

        foreach ($CSV in $CSVs) {
            Write-GuiLog "Destination CSV: $CSV" "Success"
        }

        return @($CSVs)
    }
    catch {
        Write-GuiLog "Failed to discover destination Cluster Shared Volumes: $($_.Exception.Message)" "Error"
        return @()
    }
}

function Get-DestinationVMSwitchNames {
    param(
        [Parameter(Mandatory)]
        [string]$Node
    )

    Write-GuiLog "Discovering virtual switches on destination $Node..."

    try {
        $Switches =
            @(
                Invoke-Command `
                    -ComputerName $Node `
                    -ErrorAction Stop `
                    -ScriptBlock {
                        Get-VMSwitch -ErrorAction Stop |
                            Select-Object -ExpandProperty Name |
                            Sort-Object -Unique
                    }
            )

        foreach ($Switch in $Switches) {
            Write-GuiLog "Destination VMSwitch on ${Node}: $Switch" "Success"
        }

        return @($Switches)
    }
    catch {
        Write-GuiLog "Failed to discover virtual switches on ${Node}: $($_.Exception.Message)" "Error"
        return @()
    }
}

# ---------------------------------------------------------------------------
# VM discovery
# ---------------------------------------------------------------------------

function Get-VMInventoryFromNode {
    param(
        [Parameter(Mandatory)]
        [string]$Node
    )

    Write-GuiLog "Discovering VMs on $Node..."

    try {
        $VMs =
            Invoke-Command `
                -ComputerName $Node `
                -ErrorAction Stop `
                -ScriptBlock {

                    Get-VM -ErrorAction Stop |
                    ForEach-Object {
                        $VM = $_

                        $Security =
                            Get-VMSecurity `
                                -VMName $VM.Name `
                                -ErrorAction SilentlyContinue

                        [PSCustomObject]@{
                            Name       = $VM.Name
                            Id         = $VM.Id.Guid
                            State      = [string]$VM.State
                            Generation = $VM.Generation
                            vTPM       = [bool]($Security.TpmEnabled)
                            SourceNode = $env:COMPUTERNAME
                        }
                    }
                }

        Write-GuiLog "$Node returned $(@($VMs).Count) VM(s)." "Success"
        return @($VMs)
    }
    catch {
        Write-GuiLog "Failed discovering VMs on $Node : $($_.Exception.Message)" "Error"
        return @()
    }
}

function Discover-VMs {
    $Selection = Get-ValidatedSelections

    if (-not $Selection) {
        return
    }

    $script:gridVMs.Rows.Clear()
    $script:DiscoveredVMs = @()

    foreach ($SourceNode in $Selection.SourceNodes) {
        $script:DiscoveredVMs += @(Get-VMInventoryFromNode -Node $SourceNode)
    }

    $script:DiscoveredVMs =
        @(
            $script:DiscoveredVMs |
            Sort-Object SourceNode, Name
        )

    $Destinations = @($Selection.DestNodes | Sort-Object)

    $script:DestinationVMSwitchesByNode = @{}
    foreach ($Destination in $Destinations) {
        $script:DestinationVMSwitchesByNode[$Destination] =
            @(Get-DestinationVMSwitchNames -Node $Destination)
    }

    foreach ($VM in $script:DiscoveredVMs) {
        $RowIndex = $script:gridVMs.Rows.Add()

        $Row = $script:gridVMs.Rows[$RowIndex]

        $Row.Cells["Move"].Value       = [char]0x2610
        $Row.Cells["Move"].Tag         = $false
        $Row.Cells["VMName"].Value     = $VM.Name
        $Row.Cells["SourceNode"].Value = $VM.SourceNode
        $Row.Cells["State"].Value      = $VM.State
        $Row.Cells["vTPM"].Value       = if ($VM.vTPM) { "Yes" } else { "No" }

        $DestinationCell =
            New-Object System.Windows.Forms.DataGridViewComboBoxCell

        [void]$DestinationCell.Items.AddRange([object[]]$Destinations)

        if ($Destinations.Count -gt 0) {
            $DestinationCell.Value = $Destinations[0]
        }

        $Row.Cells["DestinationNode"] = $DestinationCell

        $SwitchCell = New-Object System.Windows.Forms.DataGridViewComboBoxCell
        $DefaultDestination = [string]$DestinationCell.Value
        $SwitchChoices = @($script:DestinationVMSwitchesByNode[$DefaultDestination])
        if ($SwitchChoices.Count -gt 0) {
            [void]$SwitchCell.Items.AddRange([object[]]$SwitchChoices)
            $SwitchCell.Value = $SwitchChoices[0]
        }
        $Row.Cells["DestinationVMSwitch"] = $SwitchCell

        $StorageCell = New-Object System.Windows.Forms.DataGridViewComboBoxCell

        $SafeVMName = ($VM.Name -replace '[\\/:*?"<>|]', '_')

        $VMStorageChoices = @(
            $script:DestinationCSVPaths |
            ForEach-Object { Join-Path $_ $SafeVMName }
        )

        if ($VMStorageChoices.Count -gt 0) {
            [void]$StorageCell.Items.AddRange([object[]]$VMStorageChoices)
            $DefaultStorage =
                $VMStorageChoices |
                Where-Object { (Split-Path (Split-Path $_ -Parent) -Leaf) -notlike 'Infrastructure_*' } |
                Select-Object -First 1

            # Never default to an Infrastructure_* CSV. It remains available
            # as the last choice if the user explicitly wants it.
            if ($DefaultStorage) {
                $StorageCell.Value = $DefaultStorage
            }
        }

        $Row.Cells["DestinationStorage"] = $StorageCell
        $Row.Tag = $VM

        if ($VM.vTPM) {
            $Row.DefaultCellStyle.Font =
                New-Object System.Drawing.Font(
                    $script:gridVMs.Font,
                    [System.Drawing.FontStyle]::Bold
                )
        }
    }

    if ($script:DiscoveredVMs.Count -eq 0) {
        Show-Message -Text "No VMs were found on the selected source nodes." -Type Warning
        return
    }

    Write-GuiLog "VM discovery completed. Found $($script:DiscoveredVMs.Count) VM(s)." "Success"

    $script:tabControl.SelectedTab = $script:tabVMs
}

function Get-SelectedVMMoves {
    $Moves = @()

    foreach ($Row in $script:gridVMs.Rows) {
        if ($Row.IsNewRow) {
            continue
        }

        $DoMove = [bool]$Row.Cells["Move"].Tag

        if (-not $DoMove) {
            continue
        }

        $VMObject = $Row.Tag
        $DestNode = [string]$Row.Cells["DestinationNode"].Value

        $Moves += [PSCustomObject]@{
            VMName             = [string]$Row.Cells["VMName"].Value
            SourceNode         = [string]$Row.Cells["SourceNode"].Value
            State              = [string]$Row.Cells["State"].Value
            vTPM               = ([string]$Row.Cells["vTPM"].Value -eq "Yes")
            DestinationNode     = $DestNode
            DestinationVMSwitch = [string]$Row.Cells["DestinationVMSwitch"].Value
            DestinationStorage  = [string]$Row.Cells["DestinationStorage"].Value
            VMObject           = $VMObject
        }
    }

    return @($Moves)
}

# ---------------------------------------------------------------------------
# vTPM certificate handling
# ---------------------------------------------------------------------------

function Test-VMHasVTPM {
    param(
        [Parameter(Mandatory)]
        [string]$VMName,

        [Parameter(Mandatory)]
        [string]$SourceNode
    )

    try {
        return [bool](
            Invoke-Command `
                -ComputerName $SourceNode `
                -ErrorAction Stop `
                -ScriptBlock {
                    param($Name)
                    $Security = Get-VMSecurity -VMName $Name -ErrorAction Stop
                    [bool]$Security.TpmEnabled
                } `
                -ArgumentList $VMName
        )
    }
    catch {
        throw "Unable to determine vTPM state for '$VMName' on '$SourceNode': $($_.Exception.Message)"
    }
}

function Sync-VTPMCertificates {
    param(
        [Parameter(Mandatory)]
        [string]$SourceNode,

        [Parameter(Mandatory)]
        [string]$DestinationNode,

        [Parameter(Mandatory)]
        [string]$VMName
    )

    Write-GuiLog "vTPM detected on '$VMName'. Checking certificate material $SourceNode -> $DestinationNode..." "Warning"

    # vTPM-enabled Hyper-V VMs commonly rely on local guardian signing/encryption
    # certificates in the "Shielded VM Local Certificates" machine store.
    #
    # To avoid making assumptions about a particular guardian naming scheme, copy
    # only certificates WITH private keys from that store that are missing on the
    # destination by thumbprint.
    #
    # This copies certificate material. It does NOT remove anything from source.

    $StorePath = "Cert:\LocalMachine\Shielded VM Local Certificates"

    try {
        $SourceCerts =
            @(
                Invoke-Command `
                    -ComputerName $SourceNode `
                    -ErrorAction Stop `
                    -ScriptBlock {
                        param($Path)

                        Get-ChildItem -Path $Path -ErrorAction Stop |
                        Where-Object { $_.HasPrivateKey } |
                        Select-Object `
                            Thumbprint,
                            Subject,
                            FriendlyName,
                            NotAfter,
                            HasPrivateKey
                    } `
                    -ArgumentList $StorePath
            )
    }
    catch {
        Write-GuiLog "Unable to read vTPM certificate store on $SourceNode : $($_.Exception.Message)" "Error"
        return $false
    }

    if ($SourceCerts.Count -eq 0) {
        Write-GuiLog "No exportable candidate certificates with private keys were found in '$StorePath' on $SourceNode." "Error"
        return $false
    }

    try {
        $DestThumbprints =
            @(
                Invoke-Command `
                    -ComputerName $DestinationNode `
                    -ErrorAction Stop `
                    -ScriptBlock {
                        param($Path)

                        if (Test-Path $Path) {
                            Get-ChildItem -Path $Path -ErrorAction SilentlyContinue |
                            Select-Object -ExpandProperty Thumbprint
                        }
                    } `
                    -ArgumentList $StorePath
            )
    }
    catch {
        Write-GuiLog "Unable to read vTPM certificate store on $DestinationNode : $($_.Exception.Message)" "Error"
        return $false
    }

    $MissingCerts =
        @(
            $SourceCerts |
            Where-Object {
                $DestThumbprints -notcontains $_.Thumbprint
            }
        )

    if ($MissingCerts.Count -eq 0) {
        Write-GuiLog "Destination $DestinationNode already has all source vTPM certificate thumbprints." "Success"
        return $true
    }

    Write-GuiLog "$($MissingCerts.Count) vTPM certificate(s) need to be copied to $DestinationNode."

    # Strong random temporary PFX password. It is kept only in memory and never logged.
    $PasswordPlain =
        [Convert]::ToBase64String(
            (New-Object byte[] 48 | ForEach-Object { $_ })
        )

    # The line above creates zero bytes in Windows PowerShell; replace contents
    # with cryptographically strong random bytes.
    $RandomBytes = New-Object byte[] 48
    [System.Security.Cryptography.RandomNumberGenerator]::Create().GetBytes($RandomBytes)
    $PasswordPlain = [Convert]::ToBase64String($RandomBytes)

    foreach ($Cert in $MissingCerts) {
        $Token = [Guid]::NewGuid().Guid
        $SourceTemp = "C:\Windows\Temp\MoveIt_$Token.pfx"
        $DestTemp   = "C:\Windows\Temp\MoveIt_$Token.pfx"

        try {
            Write-GuiLog "Exporting vTPM certificate $($Cert.Thumbprint) from $SourceNode..."

            $Bytes =
                Invoke-Command `
                    -ComputerName $SourceNode `
                    -ErrorAction Stop `
                    -ScriptBlock {
                        param(
                            $Thumbprint,
                            $Store,
                            $Path,
                            $PasswordText
                        )

                        $SecurePassword =
                            ConvertTo-SecureString `
                                -String $PasswordText `
                                -AsPlainText `
                                -Force

                        $CertPath =
                            Join-Path `
                                $Store `
                                $Thumbprint

                        try {
                            Export-PfxCertificate `
                                -Cert $CertPath `
                                -FilePath $Path `
                                -Password $SecurePassword `
                                -CryptoAlgorithmOption AES256_SHA256 `
                                -ErrorAction Stop |
                                Out-Null

                            [System.IO.File]::ReadAllBytes($Path)
                        }
                        finally {
                            Remove-Item $Path -Force -ErrorAction SilentlyContinue
                        }
                    } `
                    -ArgumentList `
                        $Cert.Thumbprint,
                        $StorePath,
                        $SourceTemp,
                        $PasswordPlain

            if (-not $Bytes -or $Bytes.Count -eq 0) {
                throw "The exported PFX contained no data."
            }

            Write-GuiLog "Importing vTPM certificate $($Cert.Thumbprint) on $DestinationNode..."

            Invoke-Command `
                -ComputerName $DestinationNode `
                -ErrorAction Stop `
                -ScriptBlock {
                    param(
                        [byte[]]$PfxBytes,
                        $Store,
                        $Path,
                        $PasswordText
                    )

                    $SecurePassword =
                        ConvertTo-SecureString `
                            -String $PasswordText `
                            -AsPlainText `
                            -Force

                    try {
                        [System.IO.File]::WriteAllBytes($Path, $PfxBytes)

                        Import-PfxCertificate `
                            -FilePath $Path `
                            -CertStoreLocation $Store `
                            -Password $SecurePassword `
                            -ErrorAction Stop |
                            Out-Null
                    }
                    finally {
                        Remove-Item $Path -Force -ErrorAction SilentlyContinue
                    }
                } `
                -ArgumentList `
                    (,$Bytes),
                    $StorePath,
                    $DestTemp,
                    $PasswordPlain

            $Imported =
                Invoke-Command `
                    -ComputerName $DestinationNode `
                    -ErrorAction Stop `
                    -ScriptBlock {
                        param($Store, $Thumbprint)

                        $Cert =
                            Get-ChildItem -Path $Store -ErrorAction Stop |
                            Where-Object { $_.Thumbprint -eq $Thumbprint } |
                            Select-Object -First 1

                        [bool]($Cert -and $Cert.HasPrivateKey)
                    } `
                    -ArgumentList `
                        $StorePath,
                        $Cert.Thumbprint

            if (-not $Imported) {
                throw "Certificate appeared to import, but the destination does not show the certificate with a private key."
            }

            Write-GuiLog "vTPM certificate $($Cert.Thumbprint) copied and verified on $DestinationNode." "Success"
        }
        catch {
            Write-GuiLog "FAILED copying vTPM certificate $($Cert.Thumbprint) : $($_.Exception.Message)" "Error"
            return $false
        }
    }

    return $true
}

# ---------------------------------------------------------------------------
# VM cluster role helpers
# ---------------------------------------------------------------------------

function Get-SourceVMClusterGroup {
    param([Parameter(Mandatory)][guid]$VMId)
    $Matches = @(Get-ClusterResource -ErrorAction Stop | Where-Object {
        [string]$_.ResourceType -eq 'Virtual Machine'
    } | Where-Object {
        $IdParameter = $_ | Get-ClusterParameter -Name VmId -ErrorAction Stop
        [guid]$IdParameter.Value -eq $VMId
    })
    if ($Matches.Count -gt 1) { throw "Multiple clustered roles found for VM $VMId." }
    if ($Matches.Count -eq 1) {
        Get-ClusterGroup -Name ([string]$Matches[0].OwnerGroup) -ErrorAction Stop
    }
}

function Move-VMToLocalSourceNode {
    param([Parameter(Mandatory)]$Move)
    $VMId = [guid]$Move.VMObject.Id
    $LocalNode = $env:COMPUTERNAME
    $Group = Get-SourceVMClusterGroup -VMId $VMId
    $UseQuickMigration = ([string]$Move.State -in @('Off','Offline')) -or ($Group -and [string]$Group.State -eq 'Offline')

    if ($Group -and ([string]$Group.OwnerNode).Split('.')[0] -ine $LocalNode) {
        $Owner = [string]$Group.OwnerNode
        if ($Move.vTPM -and -not (Sync-VTPMCertificates -SourceNode $Owner -DestinationNode $LocalNode -VMName $Move.VMName)) {
            throw 'vTPM certificate preparation for the local source node failed.'
        }

        $MigrationType = if ($UseQuickMigration) { 'Quick' } else { 'Live' }
        Write-GuiLog "$MigrationType migrating '$($Move.VMName)' within the source cluster: $Owner -> $LocalNode."

        # Invoke locally in the interactive session, never through WinRM.
        Move-ClusterVirtualMachineRole -Name $Group.Name -Node $LocalNode -MigrationType $MigrationType -Wait 1800 -ErrorAction Stop | Out-Null
        $Group = Get-SourceVMClusterGroup -VMId $VMId

        if (-not $Group -or ([string]$Group.OwnerNode).Split('.')[0] -ine $LocalNode) {
            throw "The clustered VM has not completed its $($MigrationType.ToLower()) migration to this node; its role will not be removed."
        }
        if (-not $UseQuickMigration -and [string]$Group.State -ne 'Online') {
            throw 'The clustered VM has not completed its live migration to this node; its role will not be removed.'
        }
    }

    $VM = Get-VM -Id $VMId -ErrorAction Stop
    if ($Group -and -not $UseQuickMigration -and [string]$Group.State -ne 'Online') {
        throw 'Source clustered role is not Online.'
    }

    $Move.SourceNode = $LocalNode
    $Move.State = [string]$VM.State
    Write-GuiLog "'$($VM.Name)' verified on local source node $LocalNode (State=$($VM.State))." 'Success'
}

function Remove-VMFromSourceCluster {
    param(
        [Parameter(Mandatory)][string]$ClusterName,
        [Parameter(Mandatory)][string]$SourceNode,
        [Parameter(Mandatory)][string]$VMName,
        [switch]$AllowOffline
    )

    try {
        if ($SourceNode.Split('.')[0] -ine $env:COMPUTERNAME) {
            throw 'Cluster role removal must run on the local source node.'
        }
        $VM = Get-VM -Name $VMName -ErrorAction Stop
        $Group = Get-SourceVMClusterGroup -VMId $VM.Id
        if ($Group) {
            if (([string]$Group.OwnerNode).Split('.')[0] -ine $env:COMPUTERNAME) {
                throw 'VM role must be owned by this node before removal.'
            }
            if (-not $AllowOffline -and [string]$Group.State -ne 'Online') {
                throw 'VM role must be Online before removal.'
            }
            $Group | Remove-ClusterGroup -RemoveResources -Force -ErrorAction Stop
        }
        $Result = [PSCustomObject]@{ Removed = [bool]$Group }

        if (-not $Result.Removed) {
            Write-GuiLog "'$VMName' is not currently a clustered role on source cluster $ClusterName; nothing to remove." "Warning"
            return [PSCustomObject]@{ Success = $true; Removed = $false }
        }

        Write-GuiLog "'$VMName' was removed from source cluster $ClusterName on $SourceNode." "Success"
        return [PSCustomObject]@{ Success = $true; Removed = $true }
    }
    catch {
        Write-GuiLog "FAILED removing '$VMName' from source cluster $ClusterName on $SourceNode : $($_.Exception.Message)" "Error"
        return [PSCustomObject]@{ Success = $false; Removed = $false }
    }
}

function Add-VMToCluster {
    param(
        [Parameter(Mandatory)][string]$ClusterName,
        [string]$Node,
        [Parameter(Mandatory)][string]$VMName,
        [string]$Label = 'destination'
    )

    try {
        Import-Module FailoverClusters -ErrorAction Stop
        $Existing = Get-ClusterGroup -Cluster $ClusterName -Name $VMName -ErrorAction SilentlyContinue
        if ($Existing) {
            Write-GuiLog "'$VMName' is already clustered on $Label cluster $ClusterName." "Success"
            return $true
        }

        Add-ClusterVirtualMachineRole -VMName $VMName -Cluster $ClusterName -ErrorAction Stop | Out-Null
        $OwnerText = if ([string]::IsNullOrWhiteSpace($Node)) { '' } else { " (destination host $Node)" }
        Write-GuiLog "'$VMName' was added to $Label cluster $ClusterName$OwnerText." "Success"
        return $true
    }
    catch {
        Write-GuiLog "FAILED adding '$VMName' to $Label cluster $ClusterName : $($_.Exception.Message)" "Error"
        return $false
    }
}

# ---------------------------------------------------------------------------
# Move validation / execution
# ---------------------------------------------------------------------------

function Test-VMMovePlan {
    param(
        [switch]$ShowSuccessMessage
    )

    $Selection = Get-ValidatedSelections

    if (-not $Selection) {
        return $null
    }

    try {
        if (Get-Variable -Name PSSenderInfo -ValueOnly -ErrorAction SilentlyContinue) {
            throw 'Run Move It from an interactive console/RDP session on a source cluster node, not PowerShell remoting.'
        }
        $LocalCluster = Get-Cluster -ErrorAction Stop
        $SourceCluster = Get-Cluster -Name $Selection.SourceCluster -ErrorAction Stop
        if ($LocalCluster.Name -ine $SourceCluster.Name) {
            throw 'This computer is not a member of the selected source cluster.'
        }
        $LocalNode = Get-ClusterNode -Name $env:COMPUTERNAME -ErrorAction Stop
        if ([string]$LocalNode.State -ne 'Up') { throw 'The local source node must be Up and not Paused.' }
        if (-not @($Selection.SourceNodes | Where-Object { $_.Split('.')[0] -ieq $env:COMPUTERNAME }).Count) {
            throw 'The local node running Move It must be Up in the source cluster.'
        }
    }
    catch {
        Show-Message -Text $_.Exception.Message -Title 'Local Source Node Required' -Type Error
        return $null
    }

    $Moves = @(Get-SelectedVMMoves)

    if ($Moves.Count -eq 0) {
        Show-Message -Text "Select at least one VM to move." -Type Warning
        return $null
    }

    $Problems = @()

    foreach ($Move in $Moves) {
        if ([string]::IsNullOrWhiteSpace($Move.DestinationNode)) {
            $Problems += "$($Move.VMName): No destination node selected."
            continue
        }

        if ([string]::IsNullOrWhiteSpace($Move.DestinationVMSwitch)) {
            $Problems += "$($Move.VMName): No destination VMSwitch selected."
            continue
        }

        if ([string]::IsNullOrWhiteSpace($Move.DestinationStorage)) {
            $Problems += "$($Move.VMName): No destination storage path selected."
            continue
        }

        $StorageIsValid = $false

        foreach ($StorageRoot in $script:DestinationCSVPaths) {
            if ($Move.DestinationStorage.StartsWith($StorageRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
                $StorageIsValid = $true
                break
            }
        }

        if (-not $StorageIsValid) {
            $StorageLabel = if ($Selection.DestType -eq 'Cluster') { 'destination Cluster Shared Volume' } else { 'standalone destination storage root' }
            $Problems += "$($Move.VMName): Destination storage '$($Move.DestinationStorage)' is not under a discovered $StorageLabel."
            continue
        }

        if ($Selection.DestNodes -notcontains $Move.DestinationNode) {
            $Problems += "$($Move.VMName): Destination $($Move.DestinationNode) is no longer selected."
            continue
        }

        if ($Move.SourceNode -ieq $Move.DestinationNode) {
            $Problems += "$($Move.VMName): Source and destination node are the same."
            continue
        }

        # Re-query the VM so vTPM status is current.
        try {
            $CurrentVTPM =
                Test-VMHasVTPM `
                    -VMName $Move.VMName `
                    -SourceNode $Move.SourceNode

            $Move.vTPM = $CurrentVTPM
        }
        catch {
            $Problems += $_.Exception.Message
        }

        # Confirm the selected virtual switch still exists on this destination.
        try {
            $SwitchExists =
                Invoke-Command `
                    -ComputerName $Move.DestinationNode `
                    -ErrorAction Stop `
                    -ScriptBlock {
                        param($SwitchName)
                        [bool](Get-VMSwitch -Name $SwitchName -ErrorAction SilentlyContinue)
                    } `
                    -ArgumentList $Move.DestinationVMSwitch

            if (-not $SwitchExists) {
                $Problems += "$($Move.VMName): VMSwitch '$($Move.DestinationVMSwitch)' does not exist on $($Move.DestinationNode)."
            }
        }
        catch {
            $Problems += "$($Move.VMName): Could not validate VMSwitch '$($Move.DestinationVMSwitch)' on $($Move.DestinationNode): $($_.Exception.Message)"
        }

        # Ensure no VM with same name already exists on destination.
        try {
            $Exists =
                Invoke-Command `
                    -ComputerName $Move.DestinationNode `
                    -ErrorAction Stop `
                    -ScriptBlock {
                        param($Name)
                        [bool](Get-VM -Name $Name -ErrorAction SilentlyContinue)
                    } `
                    -ArgumentList $Move.VMName

            if ($Exists) {
                $Problems += "$($Move.VMName): A VM with this name already exists on $($Move.DestinationNode)."
            }
        }
        catch {
            $Problems += "$($Move.VMName): Could not query destination $($Move.DestinationNode): $($_.Exception.Message)"
        }
    }

    if ($Problems.Count -gt 0) {
        Write-GuiLog "VM move plan validation FAILED." "Error"

        foreach ($Problem in $Problems) {
            Write-GuiLog $Problem "Error"
        }

        Show-Message `
            -Text ("The VM move plan has problems:`r`n`r`n" + ($Problems -join "`r`n")) `
            -Title "Move Validation Failed" `
            -Type Error

        return $null
    }

    foreach ($Move in $Moves) {
        Write-GuiLog "MOVE PLAN: '$($Move.VMName)' $($Move.SourceNode) -> $($Move.DestinationNode) | VMSwitch=$($Move.DestinationVMSwitch) | vTPM=$($Move.vTPM)" "Success"
    }

    if ($ShowSuccessMessage) {
        Show-Message `
            -Text "The selected VM move plan passed validation." `
            -Title "Move Plan Validated" `
            -Type Information
    }

    return [PSCustomObject]@{
        Selection = $Selection
        Moves     = $Moves
    }
}

function Move-SelectedVMs {
    $Plan = Test-VMMovePlan

    if (-not $Plan) {
        return
    }

    $MoveSummary =
        @(
            $Plan.Moves |
            ForEach-Object {
                "VM: $($_.VMName)`r`nPath: $($_.SourceNode) -> LOCAL $env:COMPUTERNAME -> $($_.DestinationNode)"
            }
        )

    $ConfirmationText = @"
$($MoveSummary -join "`r`n`r`n")

- Running VMs use Live Migration; stopped VMs use Quick Migration.
- The move must be started from the source node (console or RDP).
- The VM will be removed from the source cluster and moved locally to the selected destination.
$(if ($Plan.Selection.DestType -eq 'Cluster') { '- After migration, the VM will be added to the destination cluster.' } else { '- The destination is standalone, so the VM will remain non-clustered after migration.' })
- Required vTPM/Shielded VM certificates will be copied before the move.
- If the move fails, the utility will attempt to restore the VM to the source cluster.

Continue?
"@

    $Response =
        [System.Windows.Forms.MessageBox]::Show(
            $ConfirmationText,
            "Move Selected VM(s)",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )

    if ($Response -ne [System.Windows.Forms.DialogResult]::Yes) {
        Write-GuiLog "VM move cancelled by user." "Warning"
        return
    }

    Set-ActionButtonsEnabled -Enabled $false

    try {
        # Management-first validation is required only for the local migration
        # source and the destination node(s) used by this move plan.
        $MigrationNodes = @(@($env:COMPUTERNAME) + @($Plan.Moves | ForEach-Object { $_.DestinationNode }) | Sort-Object -Unique)
        foreach ($MigrationNode in $MigrationNodes) {
            try {
                $NetworkCheck = Invoke-Command -ComputerName $MigrationNode -ScriptBlock $script:ManagementMigrationNetwork -ArgumentList $false -ErrorAction Stop
                Write-GuiLog "$MigrationNode management-first migration verified: $($NetworkCheck.ManagementIP)/32" "Success"
            }
            catch {
                Write-GuiLog "Move blocked: migration network verification failed on ${MigrationNode}: $($_.Exception.Message)" "Error"
                return
            }
        }
        foreach ($Move in $Plan.Moves) {
            Write-GuiLog "========================================"
            Write-GuiLog "Preparing '$($Move.VMName)' : $($Move.SourceNode) -> $($Move.DestinationNode)"

            try {
                Move-VMToLocalSourceNode -Move $Move
            }
            catch {
                Write-GuiLog "Skipping '$($Move.VMName)': local staging failed. $($_.Exception.Message)" 'Error'
                continue
            }

            if ($Move.vTPM) {
                if (-not (
                    Sync-VTPMCertificates `
                        -SourceNode $Move.SourceNode `
                        -DestinationNode $Move.DestinationNode `
                        -VMName $Move.VMName
                )) {
                    Write-GuiLog "Skipping '$($Move.VMName)' because vTPM certificate preparation failed." "Error"
                    continue
                }
            }

            $DestinationPath = $Move.DestinationStorage.Trim()

            Write-GuiLog "Removing '$($Move.VMName)' from source cluster before migration..."
            $AllowOfflineRemoval = ([string]$Move.State -in @('Off','Offline'))
            $SourceClusterResult = Remove-VMFromSourceCluster `
                -ClusterName $Plan.Selection.SourceCluster `
                -SourceNode $Move.SourceNode `
                -VMName $Move.VMName `
                -AllowOffline:$AllowOfflineRemoval

            if (-not $SourceClusterResult.Success) {
                Write-GuiLog "Skipping '$($Move.VMName)' because it could not be removed from the source cluster." "Error"
                continue
            }

            $RemovedFromSourceCluster = [bool]$SourceClusterResult.Removed

            $CrossClusterMoveType = if ([string]$Move.State -in @('Off','Offline')) { 'offline migration' } else { 'live migration' }
            Write-GuiLog "Starting $CrossClusterMoveType of '$($Move.VMName)' to $($Move.DestinationNode)."
            Write-GuiLog "Destination storage path: $DestinationPath"

            try {
                Write-GuiLog "Preparing compatibility report and mapping all VM adapters to destination VMSwitch '$($Move.DestinationVMSwitch)'..."

                $CompatibilityReport =
                    Compare-VM `
                        -Name $Move.VMName `
                        -DestinationHost $Move.DestinationNode `
                        -IncludeStorage `
                        -DestinationStoragePath $DestinationPath `
                        -ErrorAction Stop

                $ReportAdapters = @($CompatibilityReport.VM.NetworkAdapters)
                if ($ReportAdapters.Count -eq 0) {
                    Write-GuiLog "'$($Move.VMName)' has no virtual network adapters to map." "Warning"
                }
                else {
                    foreach ($Adapter in $ReportAdapters) {
                        Connect-VMNetworkAdapter `
                            -VMNetworkAdapter $Adapter `
                            -SwitchName $Move.DestinationVMSwitch `
                            -ErrorAction Stop
                    }
                    Write-GuiLog "Mapped $($ReportAdapters.Count) adapter(s) to destination VMSwitch '$($Move.DestinationVMSwitch)'." "Success"
                }

                Move-VM `
                    -CompatibilityReport $CompatibilityReport `
                    -ErrorAction Stop

                Write-GuiLog "'$($Move.VMName)' migration command completed successfully." "Success"

                # Verify VM exists on destination.
                $Verified =
                    Invoke-Command `
                        -ComputerName $Move.DestinationNode `
                        -ErrorAction Stop `
                        -ScriptBlock {
                            param($VMName)
                            [bool](Get-VM -Name $VMName -ErrorAction SilentlyContinue)
                        } `
                        -ArgumentList $Move.VMName

                if ($Verified) {
                    Write-GuiLog "'$($Move.VMName)' is present on destination $($Move.DestinationNode)." "Success"

                    if ($Plan.Selection.DestType -eq 'Cluster') {
                        if (-not (Add-VMToCluster `
                            -Cluster $Plan.Selection.DestCluster `
                            -VMName $Move.VMName )) {
                            Write-GuiLog "'$($Move.VMName)' moved successfully but is NOT clustered on destination. Manual cluster registration is required." "Error"
                        }
                    }
                    else {
                        Write-GuiLog "'$($Move.VMName)' moved successfully to standalone destination $($Move.DestinationNode); cluster registration skipped." "Success"
                    }
                }
                else {
                    Write-GuiLog "'$($Move.VMName)' was not found on destination after Move-VM returned." "Warning"
                }
            }
            catch {
                Write-GuiLog "MOVE FAILED for '$($Move.VMName)' : $($_.Exception.Message)" "Error"

                if ($RemovedFromSourceCluster) {
                    Write-GuiLog "Attempting to restore '$($Move.VMName)' to source cluster $($Plan.Selection.SourceCluster)..." "Warning"
                    if (-not (Add-VMToCluster `
                        -Cluster $Plan.Selection.SourceCluster `
                        -VMName $Move.VMName `
                        -Label 'source')) {
                        Write-GuiLog "ROLLBACK FAILED: '$($Move.VMName)' is not registered with the source cluster. Manual recovery is required." "Error"
                    }
                }
            }
        }

        Write-GuiLog "========================================"
        Write-GuiLog "Selected VM move processing finished."

        # Refresh inventory so successfully moved VMs disappear from source list.
        Discover-VMs
    }
    finally {
        Set-ActionButtonsEnabled -Enabled $true
    }
}

# ---------------------------------------------------------------------------
# GUI
# ---------------------------------------------------------------------------

$form = New-Object System.Windows.Forms.Form
$form.Text = "Move It $($script:MoveItVersion) - Shared-Nothing Live Migration Utility"
$form.StartPosition = "CenterScreen"
$form.Size = New-Object System.Drawing.Size(980, 590)
$form.MinimumSize = New-Object System.Drawing.Size(980, 590)
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text = "Move It $($script:MoveItVersion)"
$lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 20, [System.Drawing.FontStyle]::Bold)
$lblTitle.Location = New-Object System.Drawing.Point(25, 15)
$lblTitle.AutoSize = $true
$form.Controls.Add($lblTitle)

$lblSubTitle = New-Object System.Windows.Forms.Label
$lblSubTitle.Text = "Shared-Nothing Live Migration Utility"
$lblSubTitle.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Italic)
$lblSubTitle.Location = New-Object System.Drawing.Point(30, 53)
$lblSubTitle.AutoSize = $true
$form.Controls.Add($lblSubTitle)

$picLemur = New-Object System.Windows.Forms.PictureBox
$picLemur.Location = New-Object System.Drawing.Point(855, 8)
$picLemur.Size = New-Object System.Drawing.Size(95, 58)
$picLemur.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
$picLemur.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$picLemur.BackColor = [System.Drawing.Color]::Transparent
$picLemur.BorderStyle = [System.Windows.Forms.BorderStyle]::None

try {
    $picLemur.Image = Get-EmbeddedLemurImage
}
catch {
    Write-GuiLog "Unable to load embedded lemur image: $($_.Exception.Message)" "Warning"
}

$form.Controls.Add($picLemur)


# Cluster header fields

$lblSourceCluster = New-Object System.Windows.Forms.Label
$lblSourceCluster.Text = "Source Cluster"
$lblSourceCluster.Location = New-Object System.Drawing.Point(25, 82)
$lblSourceCluster.AutoSize = $true
$form.Controls.Add($lblSourceCluster)

$script:txtSourceCluster = New-Object System.Windows.Forms.TextBox
$script:txtSourceCluster.ReadOnly = $true
$script:txtSourceCluster.Location = New-Object System.Drawing.Point(140, 79)
$script:txtSourceCluster.Size = New-Object System.Drawing.Size(285, 25)
$form.Controls.Add($script:txtSourceCluster)

$script:btnDetect = New-Object System.Windows.Forms.Button
$script:btnDetect.Text = "Detect"
$script:btnDetect.Location = New-Object System.Drawing.Point(435, 76)
$script:btnDetect.Size = New-Object System.Drawing.Size(80, 30)
$form.Controls.Add($script:btnDetect)

$lblDestCluster = New-Object System.Windows.Forms.Label
$lblDestCluster.Text = "Destination Cluster / Node"
$lblDestCluster.Location = New-Object System.Drawing.Point(520, 82)
$lblDestCluster.AutoSize = $true
$form.Controls.Add($lblDestCluster)

$script:txtDestCluster = New-Object System.Windows.Forms.TextBox
$script:txtDestCluster.Location = New-Object System.Drawing.Point(685, 79)
$script:txtDestCluster.Size = New-Object System.Drawing.Size(140, 25)
$form.Controls.Add($script:txtDestCluster)

$script:btnRefresh = New-Object System.Windows.Forms.Button
$script:btnRefresh.Text = "Discover"
$script:btnRefresh.Location = New-Object System.Drawing.Point(835, 76)
$script:btnRefresh.Size = New-Object System.Drawing.Size(75, 30)
$form.Controls.Add($script:btnRefresh)

$lblDomain = New-Object System.Windows.Forms.Label
$lblDomain.Text = "AD DNS Domain"
$lblDomain.Location = New-Object System.Drawing.Point(25, 117)
$lblDomain.AutoSize = $true
$form.Controls.Add($lblDomain)

$script:txtDomain = New-Object System.Windows.Forms.TextBox
$script:txtDomain.Location = New-Object System.Drawing.Point(140, 114)
$script:txtDomain.Size = New-Object System.Drawing.Size(350, 25)
$form.Controls.Add($script:txtDomain)


# Tabs

$script:tabControl = New-Object System.Windows.Forms.TabControl
$script:tabControl.Location = New-Object System.Drawing.Point(25, 150)
$script:tabControl.Size = New-Object System.Drawing.Size(925, 395)
$form.Controls.Add($script:tabControl)

$tabSetup = New-Object System.Windows.Forms.TabPage
$tabSetup.Text = "1. Setup / Cleanup"
$script:tabControl.TabPages.Add($tabSetup)

$script:tabVMs = New-Object System.Windows.Forms.TabPage
$script:tabVMs.Text = "2. VM Moves"
$script:tabControl.TabPages.Add($script:tabVMs)

# Setup tab - source nodes

$grpSource = New-Object System.Windows.Forms.GroupBox
$grpSource.Text = "Source Node (this computer)"
$grpSource.Location = New-Object System.Drawing.Point(12, 12)
$grpSource.Size = New-Object System.Drawing.Size(430, 245)
$tabSetup.Controls.Add($grpSource)

$script:txtLocalSourceNode = New-Object System.Windows.Forms.TextBox
$script:txtLocalSourceNode.Text = $env:COMPUTERNAME
$script:txtLocalSourceNode.ReadOnly = $true
$script:txtLocalSourceNode.Location = New-Object System.Drawing.Point(15, 28)
$script:txtLocalSourceNode.Size = New-Object System.Drawing.Size(400, 28)
$script:txtLocalSourceNode.Font = New-Object System.Drawing.Font("Consolas", 11)
$grpSource.Controls.Add($script:txtLocalSourceNode)

$lblLocalSourceHelp = New-Object System.Windows.Forms.Label
$lblLocalSourceHelp.Text = "Run Move It while logged into this source node (console/RDP). The source is fixed to this computer.`r`n`r`nVMs on other source-cluster nodes are moved here first (Live when running, Quick when Off/Offline). Only then is their cluster role removed and Move-VM run locally.`r`n`r`nAll Up source nodes are included automatically for discovery and setup."
$lblLocalSourceHelp.Location = New-Object System.Drawing.Point(15, 68)
$lblLocalSourceHelp.Size = New-Object System.Drawing.Size(400, 165)
$grpSource.Controls.Add($lblLocalSourceHelp)

# Setup tab - destination nodes

$grpDest = New-Object System.Windows.Forms.GroupBox
$grpDest.Text = "Destination Node(s)"
$grpDest.Location = New-Object System.Drawing.Point(455, 12)
$grpDest.Size = New-Object System.Drawing.Size(430, 245)
$tabSetup.Controls.Add($grpDest)

# Custom destination-node checkbox surface.
# Native CheckedListBox checkmarks render incorrectly on some Azure Local builds.
$script:lstDestNodes = New-Object System.Windows.Forms.Panel
$script:lstDestNodes.Location = New-Object System.Drawing.Point(15, 28)
$script:lstDestNodes.Size = New-Object System.Drawing.Size(400, 155)
$script:lstDestNodes.AutoScroll = $true
$script:lstDestNodes.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$script:lstDestNodes.BackColor = [System.Drawing.SystemColors]::Window
$grpDest.Controls.Add($script:lstDestNodes)

$btnDestSelectUp = New-Object System.Windows.Forms.Button
$btnDestSelectUp.Text = "Select All Up"
$btnDestSelectUp.Location = New-Object System.Drawing.Point(15, 190)
$btnDestSelectUp.Size = New-Object System.Drawing.Size(120, 30)
$grpDest.Controls.Add($btnDestSelectUp)

$btnDestClear = New-Object System.Windows.Forms.Button
$btnDestClear.Text = "Clear"
$btnDestClear.Location = New-Object System.Drawing.Point(145, 190)
$btnDestClear.Size = New-Object System.Drawing.Size(90, 30)
$grpDest.Controls.Add($btnDestClear)

$script:btnValidate = New-Object System.Windows.Forms.Button
$script:btnValidate.Text = "Validate Nodes"
$script:btnValidate.Location = New-Object System.Drawing.Point(120, 275)
$script:btnValidate.Size = New-Object System.Drawing.Size(160, 42)
$tabSetup.Controls.Add($script:btnValidate)

$script:btnEnable = New-Object System.Windows.Forms.Button
$script:btnEnable.Text = "ENABLE Move It"
$script:btnEnable.Location = New-Object System.Drawing.Point(305, 275)
$script:btnEnable.Size = New-Object System.Drawing.Size(175, 42)
$tabSetup.Controls.Add($script:btnEnable)

$script:btnDisable = New-Object System.Windows.Forms.Button
$script:btnDisable.Text = "DISABLE Move It"
$script:btnDisable.Location = New-Object System.Drawing.Point(500, 275)
$script:btnDisable.Size = New-Object System.Drawing.Size(175, 42)
$tabSetup.Controls.Add($script:btnDisable)

$script:btnDiscoverVMs = New-Object System.Windows.Forms.Button
$script:btnDiscoverVMs.Text = "Discover VMs >"
$script:btnDiscoverVMs.Location = New-Object System.Drawing.Point(695, 275)
$script:btnDiscoverVMs.Size = New-Object System.Drawing.Size(160, 42)
$tabSetup.Controls.Add($script:btnDiscoverVMs)

# VM Moves tab

$lblVMHelp = New-Object System.Windows.Forms.Label
$lblVMHelp.Text = "Important: VM will be first be moved to the Source Node (running VMs use Live Migration; stopped VMs use Quick Migration). The cross-cluster move is performed from the source node and does not use remote migration commands."
$lblVMHelp.Location = New-Object System.Drawing.Point(15, 15)
$lblVMHelp.AutoSize = $false
$lblVMHelp.Size = New-Object System.Drawing.Size(895, 62)
$lblVMHelp.ForeColor = [System.Drawing.Color]::DarkRed
$script:tabVMs.Controls.Add($lblVMHelp)

$script:gridVMs = New-Object System.Windows.Forms.DataGridView
$script:gridVMs.Location = New-Object System.Drawing.Point(15, 83)
$script:gridVMs.Size = New-Object System.Drawing.Size(895, 205)
$script:gridVMs.AllowUserToAddRows = $false
$script:gridVMs.AllowUserToDeleteRows = $false
$script:gridVMs.RowHeadersVisible = $false
$script:gridVMs.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::DisplayedCells
$script:gridVMs.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
$script:gridVMs.MultiSelect = $false
$script:tabVMs.Controls.Add($script:gridVMs)



# Text checkbox column: Azure Local can render DataGridViewCheckBoxColumn
# as an incorrect icon.  Use Unicode checkbox glyphs with state in Cell.Tag.
$colMove = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colMove.Name = "Move"
$colMove.HeaderText = "Move"
$colMove.ReadOnly = $true
$colMove.FillWeight = 45
$colMove.DefaultCellStyle.Alignment = [System.Windows.Forms.DataGridViewContentAlignment]::MiddleCenter
$colMove.DefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI Symbol", 13)
[void]$script:gridVMs.Columns.Add($colMove)

$colVM = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colVM.Name = "VMName"
$colVM.HeaderText = "VM Name"
$colVM.ReadOnly = $true
$colVM.FillWeight = 150
[void]$script:gridVMs.Columns.Add($colVM)

$colSource = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colSource.Name = "SourceNode"
$colSource.HeaderText = "Current Owner"
$colSource.ReadOnly = $true
$colSource.FillWeight = 115
[void]$script:gridVMs.Columns.Add($colSource)

$colState = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colState.Name = "State"
$colState.HeaderText = "State"
$colState.ReadOnly = $true
$colState.FillWeight = 75
[void]$script:gridVMs.Columns.Add($colState)

$colTPM = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colTPM.Name = "vTPM"
$colTPM.HeaderText = "vTPM"
$colTPM.ReadOnly = $true
$colTPM.FillWeight = 55
[void]$script:gridVMs.Columns.Add($colTPM)

$colDest = New-Object System.Windows.Forms.DataGridViewComboBoxColumn
$colDest.Name = "DestinationNode"
$colDest.HeaderText = "Destination Node"
$colDest.FillWeight = 125
[void]$script:gridVMs.Columns.Add($colDest)

$colDestSwitch = New-Object System.Windows.Forms.DataGridViewComboBoxColumn
$colDestSwitch.Name = "DestinationVMSwitch"
$colDestSwitch.HeaderText = "Destination VMSwitch"
$colDestSwitch.FillWeight = 135
$colDestSwitch.DisplayStyle = [System.Windows.Forms.DataGridViewComboBoxDisplayStyle]::DropDownButton
[void]$script:gridVMs.Columns.Add($colDestSwitch)

$colDestStorage = New-Object System.Windows.Forms.DataGridViewComboBoxColumn
$colDestStorage.Name = "DestinationStorage"
$colDestStorage.HeaderText = "Destination Storage"
$colDestStorage.FillWeight = 185
$colDestStorage.DisplayStyle = [System.Windows.Forms.DataGridViewComboBoxDisplayStyle]::DropDownButton
[void]$script:gridVMs.Columns.Add($colDestStorage)

$btnSelectAllVMs = New-Object System.Windows.Forms.Button
$btnSelectAllVMs.Text = "Select All VMs"
$btnSelectAllVMs.Location = New-Object System.Drawing.Point(15, 300)
$btnSelectAllVMs.Size = New-Object System.Drawing.Size(120, 35)
$script:tabVMs.Controls.Add($btnSelectAllVMs)

$btnClearVMs = New-Object System.Windows.Forms.Button
$btnClearVMs.Text = "Clear"
$btnClearVMs.Location = New-Object System.Drawing.Point(145, 300)
$btnClearVMs.Size = New-Object System.Drawing.Size(85, 35)
$script:tabVMs.Controls.Add($btnClearVMs)

$script:btnValidateMoves = New-Object System.Windows.Forms.Button
$script:btnValidateMoves.Text = "Validate Move Plan"
$script:btnValidateMoves.Location = New-Object System.Drawing.Point(545, 330)
$script:btnValidateMoves.Size = New-Object System.Drawing.Size(170, 35)
$script:tabVMs.Controls.Add($script:btnValidateMoves)

$script:btnMoveVMs = New-Object System.Windows.Forms.Button
$script:btnMoveVMs.Text = "MOVE Selected VMs"
$script:btnMoveVMs.Location = New-Object System.Drawing.Point(725, 330)
$script:btnMoveVMs.Size = New-Object System.Drawing.Size(190, 35)
$script:tabVMs.Controls.Add($script:btnMoveVMs)

# ---------------------------------------------------------------------------
# Events
# ---------------------------------------------------------------------------

$script:btnDetect.Add_Click({
    $DetectedCluster = Get-LocalClusterName

    if ($DetectedCluster) {
        $script:txtSourceCluster.Text = $DetectedCluster
        Write-GuiLog "Detected local cluster: $DetectedCluster" "Success"
    }
    else {
        Write-GuiLog "Unable to detect a local failover cluster." "Warning"
        Show-Message -Text "Unable to detect a local failover cluster." -Type Warning
    }
})

$script:btnRefresh.Add_Click({
    Refresh-Clusters | Out-Null
})



$btnDestSelectUp.Add_Click({
    Select-AllUpNodes -List $script:lstDestNodes -ClusterNodes $script:DestClusterNodes
})

$btnDestClear.Add_Click({
    Clear-NodeSelections -List $script:lstDestNodes
})

$script:btnValidate.Add_Click({
    Test-CurrentSelection
})

$script:btnEnable.Add_Click({
    Enable-SharedNothingMigration
})

$script:btnDisable.Add_Click({
    Disable-SharedNothingMigration
})

$script:btnDiscoverVMs.Add_Click({
    Discover-VMs
})

$btnSelectAllVMs.Add_Click({
    foreach ($Row in $script:gridVMs.Rows) {
        if (-not $Row.IsNewRow) {
            $Row.Cells["Move"].Tag = $true
            $Row.Cells["Move"].Value = [char]0x2611
        }
    }

    $script:gridVMs.Refresh()
})

$btnClearVMs.Add_Click({
    foreach ($Row in $script:gridVMs.Rows) {
        if (-not $Row.IsNewRow) {
            $Row.Cells["Move"].Tag = $false
            $Row.Cells["Move"].Value = [char]0x2610
        }
    }

    $script:gridVMs.Refresh()
})

# Toggle the custom VM checkbox.  Only the Move column toggles selection.
$script:gridVMs.Add_CellClick({
    param($sender, $e)

    if ($e.RowIndex -lt 0 -or $e.ColumnIndex -lt 0) {
        return
    }

    if ($sender.Columns[$e.ColumnIndex].Name -ne "Move") {
        return
    }

    $Cell = $sender.Rows[$e.RowIndex].Cells["Move"]
    $Selected = -not [bool]$Cell.Tag
    $Cell.Tag = $Selected
    $Cell.Value = if ($Selected) { [char]0x2611 } else { [char]0x2610 }

    $sender.InvalidateCell($Cell)
})

$script:gridVMs.Add_CellValueChanged({
    param($sender, $e)

    if ($e.RowIndex -lt 0 -or $e.ColumnIndex -lt 0) {
        return
    }

    if ($sender.Columns[$e.ColumnIndex].Name -ne "DestinationNode") {
        return
    }

    $Row = $sender.Rows[$e.RowIndex]
    $VMName = [string]$Row.Cells["VMName"].Value

    if ([string]::IsNullOrWhiteSpace($VMName)) {
        return
    }

    $DestinationNode = [string]$Row.Cells["DestinationNode"].Value

    # VMSwitch choices are host-specific, so changing the destination node
    # immediately refreshes the switch list for that row.
    $SwitchCell = $Row.Cells["DestinationVMSwitch"]
    $SwitchCell.Items.Clear()
    $SwitchCell.Value = $null

    if (-not [string]::IsNullOrWhiteSpace($DestinationNode)) {
        $SwitchChoices = @($script:DestinationVMSwitchesByNode[$DestinationNode])

        if ($SwitchChoices.Count -eq 0) {
            $SwitchChoices = @(Get-DestinationVMSwitchNames -Node $DestinationNode)
            $script:DestinationVMSwitchesByNode[$DestinationNode] = @($SwitchChoices)
        }

        if ($SwitchChoices.Count -gt 0) {
            [void]$SwitchCell.Items.AddRange([object[]]$SwitchChoices)
            $SwitchCell.Value = $SwitchChoices[0]
        }
    }

    # Destination storage roots are discovered for the target. Cluster targets use CSVs;
    # standalone targets use the Hyper-V VirtualMachinePath. Preserve the selected path.
    $StorageCell = $Row.Cells["DestinationStorage"]

    if ($StorageCell.Items.Count -eq 0) {
        $SafeVMName = ($VMName -replace '[\\/:*?"<>|]', '_')
        $Choices = @(
            $script:DestinationCSVPaths |
            ForEach-Object { Join-Path $_ $SafeVMName }
        )

        if ($Choices.Count -gt 0) {
            [void]$StorageCell.Items.AddRange([object[]]$Choices)
            $DefaultStorage =
                $Choices |
                Where-Object { (Split-Path (Split-Path $_ -Parent) -Leaf) -notlike 'Infrastructure_*' } |
                Select-Object -First 1

            if ($DefaultStorage) {
                $StorageCell.Value = $DefaultStorage
            }
        }
    }
})

$script:gridVMs.Add_DataError({
    param($sender, $e)
    $e.ThrowException = $false
})

$script:gridVMs.Add_CurrentCellDirtyStateChanged({
    if (
        $script:gridVMs.IsCurrentCellDirty -and
        $script:gridVMs.CurrentCell -is [System.Windows.Forms.DataGridViewComboBoxCell]
    ) {
        [void]$script:gridVMs.CommitEdit(
            [System.Windows.Forms.DataGridViewDataErrorContexts]::Commit
        )
    }
})

$script:btnValidateMoves.Add_Click({
    [void](Test-VMMovePlan -ShowSuccessMessage)
})

$script:btnMoveVMs.Add_Click({
    Move-SelectedVMs
})

$form.Add_Shown({
    Write-GuiLog "Starting Move It $($script:MoveItVersion)..."

    if (-not (Initialize-RequiredModules)) {
        Set-ActionButtonsEnabled -Enabled $false
        return
    }

    $DetectedCluster = Get-LocalClusterName

    if ($DetectedCluster) {
        $script:txtSourceCluster.Text = $DetectedCluster
        Write-GuiLog "Detected source cluster: $DetectedCluster" "Success"
    }

    $DomainName = Get-DomainDNSName

    if ($DomainName) {
        $script:txtDomain.Text = $DomainName
        Write-GuiLog "Detected AD DNS domain: $DomainName" "Success"
    }

    Write-GuiLog "Enter a destination cluster name or standalone Hyper-V node and click Discover."
})

[void]$form.ShowDialog()
}
