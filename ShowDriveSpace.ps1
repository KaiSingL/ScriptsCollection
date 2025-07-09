# Script to display computer name and drive space information in a tab-delimited format for easy copying into a Word table
Write-Debug "Starting script to gather drive space information"

# Initialize array to store drive information
$driveInfo = @()

# Get computer name
try {
    Write-Debug "Retrieving computer name"
    $computerName = [System.Environment]::MachineName
}
catch {
    Write-Error "Error retrieving computer name: $_"
    Write-Debug "Exception details: $($_.Exception.Message)"
    $computerName = "Unknown"
}

# Get all drives with Win32_LogicalDisk, including fixed (3), network (4), and removable (2) drives
try {
    Write-Debug "Querying Win32_LogicalDisk for drive information"
    $drives = Get-WmiObject -Class Win32_LogicalDisk | Where-Object { $_.DriveType -in (2, 3, 4) }

    foreach ($drive in $drives) {
        Write-Debug "Processing drive: $($drive.DeviceID) (Type: $($drive.DriveType))"

        # Skip drives with no size information (e.g., inaccessible drives)
        if ($null -eq $drive.Size -or $drive.Size -eq 0) {
            Write-Debug "Skipping drive $($drive.DeviceID) due to missing or zero size"
            continue
        }

        # Calculate sizes in GB
        $totalSpaceGB = [math]::Round($drive.Size / 1GB, 1)
        $freeSpaceGB = [math]::Round($drive.FreeSpace / 1GB, 1)

        # Calculate used percentage
        $usedPercent = if ($totalSpaceGB -eq 0) { 
            Write-Debug "Total space is 0 for $($drive.DeviceID), setting used % to 0"
            0 
        } else { 
            [math]::Round((($totalSpaceGB - $freeSpaceGB) / $totalSpaceGB) * 100, 1) 
        }

        # Determine drive type for clarity (not displayed but used for debugging)
        $driveType = switch ($drive.DriveType) {
            2 { "Removable" }
            3 { "Fixed" }
            4 { "Network" }
            default { "Unknown" }
        }

        # Create custom object for drive info
        $driveInfo += [PSCustomObject]@{
            'DriveLetter'    = $drive.DeviceID
            'TotalSpaceGB'   = $totalSpaceGB
            'FreeSpaceGB'    = $freeSpaceGB
            'UsedPercent'    = $usedPercent
            'DriveType'      = $driveType
        }
    }
}
catch {
    Write-Error "Error retrieving drive information: $_"
    Write-Debug "Exception details: $($_.Exception.Message)"
    exit 1
}

# Check if any drives were found
if ($driveInfo.Count -eq 0) {
    Write-Warning "No accessible drives found."
    Write-Debug "No drives with DriveType 2, 3, or 4 detected or accessible"
    exit 0
}

# Sort drives by DriveLetter for consistent output
$driveInfo = $driveInfo | Sort-Object DriveLetter

# Display computer name
Write-Debug "Outputting computer name: $computerName"
Write-Host "Computer Name: $computerName"
Write-Host "" # Empty line for readability

# Output table in tab-delimited format
Write-Debug "Outputting table in tab-delimited format"
# Header
Write-Host "Drive Letter`tTotal Space (GB)`tFree Space (GB)`tUsed %"

# Data rows
foreach ($info in $driveInfo) {
    Write-Debug "Displaying info for drive: $($info.DriveLetter) ($($info.DriveType))"
    # Format FreeSpaceGB and UsedPercent to 0.0
    $formattedFreeSpace = "{0:F1}" -f $info.FreeSpaceGB
    $formattedUsedPercent = "{0:F1}" -f $info.UsedPercent
    Write-Host "$($info.DriveLetter)`t$($info.TotalSpaceGB)`t$formattedFreeSpace`t$formattedUsedPercent"
}

Write-Debug "Script execution completed"