Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

# Create the main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Azure Bastion Connection Manager"
$form.Size = New-Object System.Drawing.Size(600, 500)
$form.StartPosition = "CenterScreen"

# Subscription ComboBox (for selecting Azure subscription)
$subscriptionLabel = New-Object System.Windows.Forms.Label
$subscriptionLabel.Text = "Subscription:"
$subscriptionLabel.Location = New-Object System.Drawing.Point(10, 20)
$subscriptionLabel.Size = New-Object System.Drawing.Size(100, 20)
$form.Controls.Add($subscriptionLabel)

$subscriptionComboBox = New-Object System.Windows.Forms.ComboBox
$subscriptionComboBox.Location = New-Object System.Drawing.Point(120, 20)
$subscriptionComboBox.Size = New-Object System.Drawing.Size(400, 20)
$form.Controls.Add($subscriptionComboBox)

# Resource Group ComboBox
$resourceGroupLabel = New-Object System.Windows.Forms.Label
$resourceGroupLabel.Text = "Resource Group:"
$resourceGroupLabel.Location = New-Object System.Drawing.Point(10, 60)
$resourceGroupLabel.Size = New-Object System.Drawing.Size(100, 20)
$form.Controls.Add($resourceGroupLabel)

$resourceGroupComboBox = New-Object System.Windows.Forms.ComboBox
$resourceGroupComboBox.Location = New-Object System.Drawing.Point(120, 60)
$resourceGroupComboBox.Size = New-Object System.Drawing.Size(400, 20)
$form.Controls.Add($resourceGroupComboBox)

# Virtual Machine Name ComboBox
$vmNameLabel = New-Object System.Windows.Forms.Label
$vmNameLabel.Text = "VM Name:"
$vmNameLabel.Location = New-Object System.Drawing.Point(10, 100)
$vmNameLabel.Size = New-Object System.Drawing.Size(100, 20)
$form.Controls.Add($vmNameLabel)

$vmNameComboBox = New-Object System.Windows.Forms.ComboBox
$vmNameComboBox.Location = New-Object System.Drawing.Point(120, 100)
$vmNameComboBox.Size = New-Object System.Drawing.Size(400, 20)
$form.Controls.Add($vmNameComboBox)

# Bastion Name ComboBox
$bastionNameLabel = New-Object System.Windows.Forms.Label
$bastionNameLabel.Text = "Bastion Name:"
$bastionNameLabel.Location = New-Object System.Drawing.Point(10, 140)
$bastionNameLabel.Size = New-Object System.Drawing.Size(100, 20)
$form.Controls.Add($bastionNameLabel)

$bastionNameComboBox = New-Object System.Windows.Forms.ComboBox
$bastionNameComboBox.Location = New-Object System.Drawing.Point(120, 140)
$bastionNameComboBox.Size = New-Object System.Drawing.Size(400, 20)
$form.Controls.Add($bastionNameComboBox)

# Status TextBox (for output and errors)
$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Text = "Status:"
$statusLabel.Location = New-Object System.Drawing.Point(10, 180)
$statusLabel.Size = New-Object System.Drawing.Size(100, 20)
$form.Controls.Add($statusLabel)

$statusTextBox = New-Object System.Windows.Forms.TextBox
$statusTextBox.Location = New-Object System.Drawing.Point(10, 200)
$statusTextBox.Size = New-Object System.Drawing.Size(560, 200)
$statusTextBox.Multiline = $true
$statusTextBox.ScrollBars = "Vertical"
$statusTextBox.ReadOnly = $true
$form.Controls.Add($statusTextBox)

# Function to append text to the status box
function Write-Status {
    param($Message)
    $statusTextBox.AppendText("$Message`r`n")
}

# Function to populate subscriptions
function Populate-Subscriptions {
    try {
        Write-Status "Retrieving available Azure subscriptions..."
        $subscriptions = & az account list --query "[].{Name:name, Id:id, IsDefault:isDefault}" --output json | ConvertFrom-Json
        if ($LASTEXITCODE -ne 0) {
            Write-Status "Error retrieving subscriptions. Please ensure you are logged into Azure CLI using 'az login'."
            return
        }
        $subscriptionComboBox.Items.Clear()
        foreach ($sub in $subscriptions) {
            $displayName = "$($sub.Name) ($($sub.Id))"
            if ($sub.IsDefault) {
                $displayName += " [Default]"
            }
            $subscriptionComboBox.Items.Add($displayName) | Out-Null
            if ($sub.IsDefault) {
                $subscriptionComboBox.SelectedIndex = $subscriptionComboBox.Items.Count - 1
            }
        }
    } catch {
        Write-Status "Error retrieving subscriptions: $_"
    }
}

# Function to populate resource groups
function Populate-ResourceGroups {
    param($SubscriptionId)
    try {
        Write-Status "Retrieving resource groups for subscription $SubscriptionId..."
        $resourceGroups = & az group list --subscription $SubscriptionId --query "[].name" --output tsv
        if ($LASTEXITCODE -ne 0) {
            Write-Status "Error retrieving resource groups: $resourceGroups"
            return
        }
        $resourceGroupComboBox.Items.Clear()
        foreach ($rg in $resourceGroups) {
            $resourceGroupComboBox.Items.Add($rg) | Out-Null
        }
    } catch {
        Write-Status "Error retrieving resource groups: $_"
    }
}

# Function to populate VMs
function Populate-VMs {
    param($ResourceGroup)
    try {
        Write-Status "Retrieving VMs for resource group $ResourceGroup..."
        $vms = & az vm list --resource-group $ResourceGroup --query "[].name" --output tsv
        if ($LASTEXITCODE -ne 0) {
            Write-Status "Error retrieving VMs: $vms"
            return
        }
        $vmNameComboBox.Items.Clear()
        foreach ($vm in $vms) {
            $vmNameComboBox.Items.Add($vm) | Out-Null
        }
    } catch {
        Write-Status "Error retrieving VMs: $_"
    }
}

# Function to populate Bastion hosts (subscription-wide)
function Populate-Bastions {
    param($SubscriptionId)
    try {
        Write-Status "Retrieving all Bastion hosts in subscription $SubscriptionId..."
        $bastions = & az network bastion list --subscription $SubscriptionId --query "[].{Name:name, ResourceGroup:resourceGroup, Location:location}" --output json | ConvertFrom-Json
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
    } catch {
        Write-Status "Error retrieving Bastion hosts: $_"
    }
}

# Function to get VM resource ID
function Get-VMResourceId {
    param($ResourceGroup, $VMName)
    try {
        Write-Status "Retrieving resource ID for VM $VMName in resource group $ResourceGroup..."
        $vmResourceId = & az vm show --resource-group $ResourceGroup --name $VMName --query id --output tsv 2>&1
        if ($LASTEXITCODE -ne 0) {
            Write-Status "Error retrieving VM resource ID: $vmResourceId"
            return $null
        }
        Write-Status "Successfully retrieved VM resource ID: $vmResourceId"
        return $vmResourceId
    } catch {
        Write-Status "Error retrieving VM resource ID: $_"
        return $null
    }
}

# Event handler for subscription selection
$subscriptionComboBox.Add_SelectedIndexChanged({
    $selectedSubscription = $subscriptionComboBox.SelectedItem -replace " \[Default\]", ""
    $subscriptionId = ($selectedSubscription -split " \(")[1] -replace "\)", ""
    Write-Status "Setting active subscription to $subscriptionId..."
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

# RDP Button
$rdpButton = New-Object System.Windows.Forms.Button
$rdpButton.Text = "Connect via RDP"
$rdpButton.Location = New-Object System.Drawing.Point(10, 420)
$rdpButton.Size = New-Object System.Drawing.Size(150, 30)
$rdpButton.Add_Click({
    $resourceGroup = $resourceGroupComboBox.SelectedItem
    $vmName = $vmNameComboBox.SelectedItem
    $bastionDisplayName = $bastionNameComboBox.SelectedItem

    if (-not $resourceGroup -or -not $vmName -or -not $bastionDisplayName) {
        [System.Windows.Forms.MessageBox]::Show("All fields are required!", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    $bastionInfo = $script:bastionMapping[$bastionDisplayName]
    $bastionName = $bastionInfo.Name
    $bastionResourceGroup = $bastionInfo.ResourceGroup

    $vmResourceId = Get-VMResourceId -ResourceGroup $resourceGroup -VMName $vmName
    if (-not $vmResourceId) {
        return
    }

    Write-Status "Initiating RDP connection via Bastion ($bastionName in $bastionResourceGroup)..."
    try {
        $output = & az network bastion rdp --name $bastionName --resource-group $bastionResourceGroup --target-resource-id $vmResourceId 2>&1
        if ($LASTEXITCODE -ne 0) {
            Write-Status "Error initiating RDP: $output"
        } else {
            Write-Status "RDP connection initiated successfully."
        }
    } catch {
        Write-Status "Error initiating RDP: $_"
    }
})
$form.Controls.Add($rdpButton)

# SSH Button
$sshButton = New-Object System.Windows.Forms.Button
$sshButton.Text = "Connect via SSH"
$sshButton.Location = New-Object System.Drawing.Point(170, 420)
$sshButton.Size = New-Object System.Drawing.Size(150, 30)
$sshButton.Add_Click({
    $resourceGroup = $resourceGroupComboBox.SelectedItem
    $vmName = $vmNameComboBox.SelectedItem
    $bastionDisplayName = $bastionNameComboBox.SelectedItem

    if (-not $resourceGroup -or -not $vmName -or -not $bastionDisplayName) {
        [System.Windows.Forms.MessageBox]::Show("All fields are required!", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    $bastionInfo = $script:bastionMapping[$bastionDisplayName]
    $bastionName = $bastionInfo.Name
    $bastionResourceGroup = $bastionInfo.ResourceGroup

    $vmResourceId = Get-VMResourceId -ResourceGroup $resourceGroup -VMName $vmName
    if (-not $vmResourceId) {
        return
    }

    Write-Status "Initiating SSH connection via Bastion ($bastionName in $bastionResourceGroup)..."
    try {
        $output = & az network bastion ssh --name $bastionName --resource-group $bastionResourceGroup --target-resource-id $vmResourceId 2>&1
        if ($LASTEXITCODE -ne 0) {
            Write-Status "Error initiating SSH: $output"
        } else {
            Write-Status "SSH connection initiated successfully."
        }
    } catch {
        Write-Status "Error initiating SSH: $_"
    }
})
$form.Controls.Add($sshButton)

# Initialize the GUI
Populate-Subscriptions

# Show the form
$form.ShowDialog()