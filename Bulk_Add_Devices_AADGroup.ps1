# Install AzureAD module if not already installed
Set-ExecutionPolicy -ExecutionPolicy Unrestricted
Install-Module AzureAD

# Connect to Azure AD
Connect-AzureAD

# Define the path to the CSV or text file containing device names or IDs
$devicesFile = "C:\scripts\devices.csv" # Update with your file path

# Define the name or object ID of the Azure AD security group
$groupName = "APP - Zoom Meetings Client" # Update with your security group name or object ID

# Read the devices from the file
$devices = Get-Content $devicesFile

# Get the Azure AD security group
$group = Get-AzureADGroup -Filter "displayName eq '$groupName'" 

if ($group -eq $null) {
    Write-Error "Security group '$groupName' not found."
    Exit
}

# Add devices to the security group
foreach ($device in $devices) {
    $deviceObject = Get-AzureADDevice -Filter "displayName eq '$device'"
    if ($deviceObject -eq $null) {
        Write-Warning "Device '$device' not found in Azure AD."
        continue
    }

    $existingMembers = Get-AzureADGroupMember -ObjectId $group.ObjectId | Select-Object -ExpandProperty ObjectId
    if ($existingMembers -notcontains $deviceObject.ObjectId) {
        Add-AzureADGroupMember -ObjectId $group.ObjectId -RefObjectId $deviceObject.ObjectId
        Write-Host "Added device $device to security group $groupName"
    } else {
        Write-Host "Device $device is already a member of security group $groupName"
    }
}

# Disconnect from Azure AD
Disconnect-AzureAD