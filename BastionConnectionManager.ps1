Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

# Add code to hide the console window more robustly
Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("kernel32.dll")]
public static extern IntPtr GetConsoleWindow();

[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

[DllImport("user32.dll")]
public static extern bool SetForegroundWindow(IntPtr hWnd);

public static void HideConsoleWindow() {
    IntPtr hWnd = GetConsoleWindow();
    if (hWnd != IntPtr.Zero) {
        ShowWindow(hWnd, 0); // 0 = SW_HIDE
        // Ensure the console window does not regain focus
        SetForegroundWindow(IntPtr.Zero);
    }
}
'

# Hide the console window immediately
[Console.Window]::HideConsoleWindow()

# Function to ensure the console remains hidden even if it tries to reappear
$hideConsoleTimer = New-Object System.Windows.Forms.Timer
$hideConsoleTimer.Interval = 100  # Check every 100ms
$hideConsoleTimer.Add_Tick({
    [Console.Window]::HideConsoleWindow()
})
$hideConsoleTimer.Start()

# Create the main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Azure Bastion Connection Manager"
$form.Size = New-Object System.Drawing.Size(600, 580)  # Height set to 580 to fully display buttons
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable  # Allow resizing
$form.MaximizeBox = $true  # Enable maximize button

# Create a tooltip object
$tooltip = New-Object System.Windows.Forms.ToolTip

# Subscription ComboBox (for selecting Azure subscription)
$subscriptionLabel = New-Object System.Windows.Forms.Label
$subscriptionLabel.Text = "Subscription:"
$subscriptionLabel.Location = New-Object System.Drawing.Point(10, 20)
$subscriptionLabel.Size = New-Object System.Drawing.Size(100, 20)
$subscriptionLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
$form.Controls.Add($subscriptionLabel)

$subscriptionComboBox = New-Object System.Windows.Forms.ComboBox
$subscriptionComboBox.Location = New-Object System.Drawing.Point(120, 20)
$subscriptionComboBox.Size = New-Object System.Drawing.Size(400, 20)
$subscriptionComboBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$tooltip.SetToolTip($subscriptionComboBox, "Select an Azure subscription to work with.")
$form.Controls.Add($subscriptionComboBox)

# Checkbox to include "Azure for Students" subscriptions
$includeStudentsCheckBox = New-Object System.Windows.Forms.CheckBox
$includeStudentsCheckBox.Text = "Include 'Azure for Students'"
$includeStudentsCheckBox.Location = New-Object System.Drawing.Point(120, 50)
$includeStudentsCheckBox.Size = New-Object System.Drawing.Size(200, 20)
$includeStudentsCheckBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
$tooltip.SetToolTip($includeStudentsCheckBox, "Check to include 'Azure for Students' subscriptions in the list.")
$form.Controls.Add($includeStudentsCheckBox)

# Resource Group ComboBox
$resourceGroupLabel = New-Object System.Windows.Forms.Label
$resourceGroupLabel.Text = "Resource Group:"
$resourceGroupLabel.Location = New-Object System.Drawing.Point(10, 80)
$resourceGroupLabel.Size = New-Object System.Drawing.Size(100, 20)
$resourceGroupLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
$form.Controls.Add($resourceGroupLabel)

$resourceGroupComboBox = New-Object System.Windows.Forms.ComboBox
$resourceGroupComboBox.Location = New-Object System.Drawing.Point(120, 80)
$resourceGroupComboBox.Size = New-Object System.Drawing.Size(400, 20)
$resourceGroupComboBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$tooltip.SetToolTip($resourceGroupComboBox, "Select a resource group within the chosen subscription.")
$form.Controls.Add($resourceGroupComboBox)

# Virtual Machine Name ComboBox
$vmNameLabel = New-Object System.Windows.Forms.Label
$vmNameLabel.Text = "VM Name:"
$vmNameLabel.Location = New-Object System.Drawing.Point(10, 110)
$vmNameLabel.Size = New-Object System.Drawing.Size(100, 20)
$vmNameLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
$form.Controls.Add($vmNameLabel)

$vmNameComboBox = New-Object System.Windows.Forms.ComboBox
$vmNameComboBox.Location = New-Object System.Drawing.Point(120, 110)
$vmNameComboBox.Size = New-Object System.Drawing.Size(400, 20)
$vmNameComboBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$tooltip.SetToolTip($vmNameComboBox, "Select a virtual machine within the chosen resource group.")
$form.Controls.Add($vmNameComboBox)

# Bastion Name ComboBox
$bastionNameLabel = New-Object System.Windows.Forms.Label
$bastionNameLabel.Text = "Bastion Name:"
$bastionNameLabel.Location = New-Object System.Drawing.Point(10, 140)
$bastionNameLabel.Size = New-Object System.Drawing.Size(100, 20)
$bastionNameLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
$form.Controls.Add($bastionNameLabel)

$bastionNameComboBox = New-Object System.Windows.Forms.ComboBox
$bastionNameComboBox.Location = New-Object System.Drawing.Point(120, 140)
$bastionNameComboBox.Size = New-Object System.Drawing.Size(400, 20)
$bastionNameComboBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$tooltip.SetToolTip($bastionNameComboBox, "Select a Bastion host to use for the connection.")
$form.Controls.Add($bastionNameComboBox)

# Status TextBox (for output and errors)
$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Text = "Status:"
$statusLabel.Location = New-Object System.Drawing.Point(10, 170)
$statusLabel.Size = New-Object System.Drawing.Size(100, 20)
$statusLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
$form.Controls.Add($statusLabel)

$statusTextBox = New-Object System.Windows.Forms.TextBox
$statusTextBox.Location = New-Object System.Drawing.Point(10, 190)
$statusTextBox.Size = New-Object System.Drawing.Size(560, 300)
$statusTextBox.Multiline = $true
$statusTextBox.ScrollBars = "Vertical"
$statusTextBox.ReadOnly = $true
$statusTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$tooltip.SetToolTip($statusTextBox, "Displays logs, errors, and status messages.")
$form.Controls.Add($statusTextBox)

# Function to append text to the status box or update in-place
function Write-Status {
    param(
        [string]$Message,
        [switch]$UpdateInPlace  # If true, updates the last line instead of appending
    )
    if ($UpdateInPlace) {
        # Update the last line in the status text box
        $lines = $statusTextBox.Text -split "`r`n"
        if ($lines.Count -gt 1) {
            $lines[$lines.Count - 2] = $Message  # Update the last non-empty line
            $statusTextBox.Text = $lines -join "`r`n"
        } else {
            $statusTextBox.Text = $Message
        }
    } else {
        # Append a new line to the status text box
        $statusTextBox.AppendText("$Message`r`n")
    }
}

# Function to check Azure CLI login status
function Check-AzureLogin {
    try {
        $account = & az account show --output json 2>&1
        if ($LASTEXITCODE -ne 0) {
            Write-Status "You are not logged into Azure CLI. Please run 'az login' and try again."
            [System.Windows.Forms.MessageBox]::Show("You are not logged into Azure CLI. Please run 'az login' in a terminal and restart this application.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return $false
        }
        Write-Status "Azure CLI is authenticated."
        return $true
    } catch {
        Write-Status "Error checking Azure CLI login: $_"
        return $false
    }
}

# Function to validate VM connection type
function Validate-VMConnection {
    param($ResourceGroup, $VMName, $ConnectionType)
    try {
        Write-Status "Validating VM connection type for $VMName in $ResourceGroup..." -UpdateInPlace
        $osType = & az vm show --resource-group $ResourceGroup --name $VMName --query "storageProfile.osDisk.osType" --output tsv 2>&1
        if ($LASTEXITCODE -ne 0) {
            Write-Status "Error retrieving VM OS type: $osType"
            return $false
        }
        if ($ConnectionType -eq "RDP" -and $osType -ne "Windows") {
            Write-Status "RDP is only supported for Windows VMs. Selected VM OS: $osType"
            return $false
        } elseif ($ConnectionType -eq "SSH" -and $osType -ne "Linux") {
            Write-Status "SSH is only supported for Linux VMs. Selected VM OS: $osType"
            return $false
        }
        Write-Status "VM connection type validated successfully."
        return $true
    } catch {
        Write-Status "Error validating VM connection: $_"
        return $false
    }
}

# Function to populate subscriptions (configurable filtering)
function Populate-Subscriptions {
    try {
        Write-Status "Retrieving available Azure subscriptions..." -UpdateInPlace
        $job = Start-Job -ScriptBlock {
            & az account list --query "[].{Name:name, Id:id, IsDefault:isDefault}" --output json
        }
        while ($job.State -eq "Running") {
            Start-Sleep -Milliseconds 100
            Write-Status "Retrieving available Azure subscriptions... (please wait)" -UpdateInPlace
        }
        $subscriptions = Receive-Job -Job $job | ConvertFrom-Json
        Remove-Job -Job $job
        if ($LASTEXITCODE -ne 0) {
            Write-Status "Error retrieving subscriptions. Please ensure you are logged into Azure CLI using 'az login'."
            return
        }
        if (-not $includeStudentsCheckBox.Checked) {
            $subscriptions = $subscriptions | Where-Object { $_.Name -notlike "*Azure for Students*" }
        }
        if (-not $subscriptions) {
            Write-Status "No subscriptions found. Please check your subscription list or filtering options."
            return
        }
        $subscriptionComboBox.Items.Clear()
        foreach ($sub in $subscriptions) {
            $displayName = "$($sub.Name) ($($sub.Id))"
            if ($sub.IsDefault) {
                $displayName += " [Default]"
            }
            $subscriptionComboBox.Items.Add($displayName) | Out-Null
            if ($sub.IsDefault -and (-not $includeStudentsCheckBox.Checked -or $sub.Name -notlike "*Azure for Students*")) {
                $subscriptionComboBox.SelectedIndex = $subscriptionComboBox.Items.Count - 1
            }
        }
        if ($subscriptionComboBox.SelectedIndex -eq -1 -and $subscriptionComboBox.Items.Count -gt 0) {
            $subscriptionComboBox.SelectedIndex = 0  # Select the first subscription if no default is available
        }
        Write-Status "Successfully retrieved Azure subscriptions."
    } catch {
        Write-Status "Error retrieving subscriptions: $_"
    }
}

# Function to populate resource groups
function Populate-ResourceGroups {
    param($SubscriptionId)
    try {
        Write-Status "Retrieving resource groups for subscription $SubscriptionId..." -UpdateInPlace
        $job = Start-Job -ScriptBlock {
            param($subId)
            & az group list --subscription $subId --query "[].name" --output tsv
        } -ArgumentList $SubscriptionId
        while ($job.State -eq "Running") {
            Start-Sleep -Milliseconds 100
            Write-Status "Retrieving resource groups for subscription $SubscriptionId... (please wait)" -UpdateInPlace
        }
        $resourceGroups = Receive-Job -Job $job
        Remove-Job -Job $job
        if ($LASTEXITCODE -ne 0) {
            Write-Status "Error retrieving resource groups: $resourceGroups"
            return
        }
        $resourceGroupComboBox.Items.Clear()
        foreach ($rg in $resourceGroups) {
            $resourceGroupComboBox.Items.Add($rg) | Out-Null
        }
        Write-Status "Successfully retrieved resource groups for subscription $SubscriptionId."
    } catch {
        Write-Status "Error retrieving resource groups: $_"
    }
}

# Function to populate VMs
function Populate-VMs {
    param($ResourceGroup)
    try {
        Write-Status "Retrieving VMs for resource group $ResourceGroup..." -UpdateInPlace
        $job = Start-Job -ScriptBlock {
            param($rg)
            & az vm list --resource-group $rg --query "[].name" --output tsv
        } -ArgumentList $ResourceGroup
        while ($job.State -eq "Running") {
            Start-Sleep -Milliseconds 100
            Write-Status "Retrieving VMs for resource group $ResourceGroup... (please wait)" -UpdateInPlace
        }
        $vms = Receive-Job -Job $job
        Remove-Job -Job $job
        if ($LASTEXITCODE -ne 0) {
            Write-Status "Error retrieving VMs: $vms"
            return
        }
        $vmNameComboBox.Items.Clear()
        foreach ($vm in $vms) {
            $vmNameComboBox.Items.Add($vm) | Out-Null
        }
        Write-Status "Successfully retrieved VMs for resource group $ResourceGroup."
    } catch {
        Write-Status "Error retrieving VMs: $_"
    }
}

# Function to populate Bastion hosts (subscription-wide)
function Populate-Bastions {
    param($SubscriptionId)
    try {
        Write-Status "Retrieving all Bastion hosts in subscription $SubscriptionId..." -UpdateInPlace
        $job = Start-Job -ScriptBlock {
            param($subId)
            & az network bastion list --subscription $subId --query "[].{Name:name, ResourceGroup:resourceGroup, Location:location}" --output json
        } -ArgumentList $SubscriptionId
        while ($job.State -eq "Running") {
            Start-Sleep -Milliseconds 100
            Write-Status "Retrieving all Bastion hosts in subscription $SubscriptionId... (please wait)" -UpdateInPlace
        }
        $bastions = Receive-Job -Job $job | ConvertFrom-Json
        Remove-Job -Job $job
        if ($LASTEXITCODE -ne 0) {
            Write-Status "Error retrieving Bastion hosts: $bastions"
            return
        }
        $bastionNameComboBox.Items.Clear()
        $script:bastionMapping = @{}
        foreach ($bastion in $bastions) {
            $displayName = "$($bastion.Name) (RG: $($bastion.ResourceGroup), Location: $($bastion.Location))"
            $bastionNameComboBox.Items.Add($displayName) | Out-Null
            $script:bastionMapping[$displayName] = @{
                Name = $bastion.Name
                ResourceGroup = $bastion.ResourceGroup
            }
        }
        Write-Status "Successfully retrieved Bastion hosts for subscription $SubscriptionId."
    } catch {
        Write-Status "Error retrieving Bastion hosts: $_"
    }
}

# Function to get VM resource ID
function Get-VMResourceId {
    param($ResourceGroup, $VMName)
    try {
        Write-Status "Retrieving resource ID for VM $VMName in resource group $ResourceGroup..." -UpdateInPlace
        $job = Start-Job -ScriptBlock {
            param($rg, $vm)
            & az vm show --resource-group $rg --name $vm --query id --output tsv 2>&1
        } -ArgumentList $ResourceGroup, $VMName
        while ($job.State -eq "Running") {
            Start-Sleep -Milliseconds 100
            Write-Status "Retrieving resource ID for VM $VMName in resource group $ResourceGroup... (please wait)" -UpdateInPlace
        }
        $vmResourceId = Receive-Job -Job $job
        Remove-Job -Job $job
        if ($LASTEXITCODE -ne 0) {
            Write-Status "Error retrieving VM resource ID: $vmResourceId"
            return $null
        }
        Write-Status "Successfully retrieved VM resource ID."
        return $vmResourceId
    } catch {
        Write-Status "Error retrieving VM resource ID: $_"
        return $null
    }
}

# Function to validate Bastion connectivity (basic check)
function Validate-BastionConnectivity {
    param($ResourceGroup, $VMName, $BastionName, $BastionResourceGroup)
    try {
        Write-Status "Validating Bastion connectivity for VM $VMName and Bastion $BastionName..." -UpdateInPlace
        $vmVnet = & az vm show --resource-group $ResourceGroup --name $VMName --query "networkProfile.networkInterfaces[0].id" --output tsv | ForEach-Object { & az network nic show --ids $_ --query "ipConfigurations[0].subnet.id" --output tsv } | ForEach-Object { & az network vnet show --ids $_ --query "name" --output tsv }
        $bastionVnet = & az network bastion show --name $BastionName --resource-group $BastionResourceGroup --query "ipConfigurations[0].subnet.id" --output tsv | ForEach-Object { & az network vnet show --ids $_ --query "name" --output tsv }
        if ($vmVnet -ne $bastionVnet) {
            Write-Status "Warning: VM and Bastion are in different VNets ($vmVnet vs $bastionVnet). Ensure VNet peering or correct configuration is in place."
            return $false
        }
        Write-Status "Bastion connectivity validated successfully."
        return $true
    } catch {
        Write-Status "Error validating Bastion connectivity: $_"
        return $false
    }
}

# Event handler for subscription selection
$subscriptionComboBox.Add_SelectedIndexChanged({
    $selectedSubscription = $subscriptionComboBox.SelectedItem -replace " \[Default\]", ""
    $subscriptionId = ($selectedSubscription -split " \(")[1] -replace "\)", ""
    Write-Status "Setting active subscription to $subscriptionId..." -UpdateInPlace
    try {
        $output = & az account set --subscription $subscriptionId 2>&1
        if ($LASTEXITCODE -ne 0) {
            Write-Status "Error setting subscription: $output"
            return
        }
        Write-Status "Successfully set active subscription to $subscriptionId."
        Populate-ResourceGroups -SubscriptionId $subscriptionId
        Populate-Bastions -SubscriptionId $subscriptionId
    } catch {
        Write-Status "Error setting subscription: $_"
    }
})

# Event handler for resource group selection
$resourceGroupComboBox.Add_SelectedIndexChanged({
    $selectedResourceGroup = $resourceGroupComboBox.SelectedItem
    if ($selectedResourceGroup) {
        Populate-VMs -ResourceGroup $selectedResourceGroup
    }
})

# Event handler for "Include Azure for Students" checkbox
$includeStudentsCheckBox.Add_CheckedChanged({
    Populate-Subscriptions
})

# RDP Button
$rdpButton = New-Object System.Windows.Forms.Button
$rdpButton.Text = "Connect via RDP"
$rdpButton.Location = New-Object System.Drawing.Point(10, 500)
$rdpButton.Size = New-Object System.Drawing.Size(150, 30)
$rdpButton.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$rdpButton.Add_Click({
    $resourceGroup = $resourceGroupComboBox.SelectedItem
    $vmName = $vmNameComboBox.SelectedItem
    $bastionDisplayName = $bastionNameComboBox.SelectedItem

    if (-not $resourceGroup -or -not $vmName -or -not $bastionDisplayName) {
        [System.Windows.Forms.MessageBox]::Show("All fields (Resource Group, VM Name, Bastion Name) are required!", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    if (-not (Validate-VMConnection -ResourceGroup $resourceGroup -VMName $vmName -ConnectionType "RDP")) {
        return
    }

    $bastionInfo = $script:bastionMapping[$bastionDisplayName]
    $bastionName = $bastionInfo.Name
    $bastionResourceGroup = $bastionInfo.ResourceGroup

    if (-not (Validate-BastionConnectivity -ResourceGroup $resourceGroup -VMName $vmName -BastionName $bastionName -BastionResourceGroup $bastionResourceGroup)) {
        return
    }

    $vmResourceId = Get-VMResourceId -ResourceGroup $resourceGroup -VMName $vmName
    if (-not $vmResourceId) {
        return
    }

    Write-Status "Initiating RDP connection via Bastion ($bastionName in $bastionResourceGroup)..." -UpdateInPlace
    try {
        $cmd = "az network bastion rdp --name $bastionName --resource-group $bastionResourceGroup --target-resource-id $vmResourceId"
        $output = & powershell -Command $cmd 2>&1
        if ($LASTEXITCODE -ne 0) {
            Write-Status "Error initiating RDP: $output"
        } else {
            Write-Status "RDP connection initiated successfully."
        }
    } catch {
        Write-Status "Error initiating RDP: $_"
    }
})
$tooltip.SetToolTip($rdpButton, "Initiates an RDP connection to the selected VM via Azure Bastion.")
$form.Controls.Add($rdpButton)

# SSH Button
$sshButton = New-Object System.Windows.Forms.Button
$sshButton.Text = "Connect via SSH"
$sshButton.Location = New-Object System.Drawing.Point(170, 500)
$sshButton.Size = New-Object System.Drawing.Size(150, 30)
$sshButton.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$sshButton.Add_Click({
    $resourceGroup = $resourceGroupComboBox.SelectedItem
    $vmName = $vmNameComboBox.SelectedItem
    $bastionDisplayName = $bastionNameComboBox.SelectedItem

    if (-not $resourceGroup -or -not $vmName -or -not $bastionDisplayName) {
        [System.Windows.Forms.MessageBox]::Show("All fields (Resource Group, VM Name, Bastion Name) are required!", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    if (-not (Validate-VMConnection -ResourceGroup $resourceGroup -VMName $vmName -ConnectionType "SSH")) {
        return
    }

    $bastionInfo = $script:bastionMapping[$bastionDisplayName]
    $bastionName = $bastionInfo.Name
    $bastionResourceGroup = $bastionInfo.ResourceGroup

    if (-not (Validate-BastionConnectivity -ResourceGroup $resourceGroup -VMName $vmName -BastionName $bastionName -BastionResourceGroup $bastionResourceGroup)) {
        return
    }

    $vmResourceId = Get-VMResourceId -ResourceGroup $resourceGroup -VMName $vmName
    if (-not $vmResourceId) {
        return
    }

    Write-Status "Initiating SSH connection via Bastion ($bastionName in $bastionResourceGroup)..." -UpdateInPlace
    try {
        $cmd = "az network bastion ssh --name $bastionName --resource-group $bastionResourceGroup --target-resource-id $vmResourceId"
        $output = & powershell -Command $cmd 2>&1
        if ($LASTEXITCODE -ne 0) {
            Write-Status "Error initiating SSH: $output"
        } else {
            Write-Status "SSH connection initiated successfully."
        }
    } catch {
        Write-Status "Error initiating SSH: $_"
    }
})
$tooltip.SetToolTip($sshButton, "Initiates an SSH connection to the selected VM via Azure Bastion.")
$form.Controls.Add($sshButton)

# Refresh Button
$refreshButton = New-Object System.Windows.Forms.Button
$refreshButton.Text = "Refresh"
$refreshButton.Location = New-Object System.Drawing.Point(330, 500)
$refreshButton.Size = New-Object System.Drawing.Size(100, 30)
$refreshButton.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$refreshButton.Add_Click({
    # Clear all dropdowns and reset selections
    $subscriptionComboBox.Items.Clear()
    $subscriptionComboBox.SelectedIndex = -1
    $resourceGroupComboBox.Items.Clear()
    $resourceGroupComboBox.SelectedIndex = -1
    $vmNameComboBox.Items.Clear()
    $vmNameComboBox.SelectedIndex = -1
    $bastionNameComboBox.Items.Clear()
    $bastionNameComboBox.SelectedIndex = -1
    Write-Status "Refreshing form..."
    Populate-Subscriptions
})
$tooltip.SetToolTip($refreshButton, "Refreshes the list of subscriptions, resource groups, VMs, and Bastion hosts.")
$form.Controls.Add($refreshButton)

# Check Azure CLI login before proceeding
if (-not (Check-AzureLogin)) {
    $hideConsoleTimer.Stop()
    exit
}

# Initialize the GUI
Populate-Subscriptions

# Show the form and stop the timer when the form is closed
$form.Add_FormClosed({
    $hideConsoleTimer.Stop()
})
$form.ShowDialog()
