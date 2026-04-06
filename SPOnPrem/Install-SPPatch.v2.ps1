# Define the variables
$SPServers = "FE1", "FE2", "FE3", "FE4", "DCH1", "DCH2", "SRCH1", "SRCH2", "APP1", "APP2"
$PatchPath = "C:\SharePointPatches\"
$LogFile = "C:\SharePointPatches\CU.log"

# Check if the patch folder exists
if (!(Test-Path $PatchPath)) {
    Write-Host "Patch folder does not exist. Please create the folder and copy the patches to the folder." -ForegroundColor Red
    Exit
}

# Loop through each server and check the patch version
foreach ($SPServer in $SPServers) {
    Write-Host "Checking patch version on $SPServer..." -ForegroundColor Yellow
    Invoke-Command -ComputerName $SPServer -ScriptBlock {
        Get-HotFix | Where-Object {$_.Description -like "*SharePoint*"} | Select-Object -Property Description, HotFixID, InstalledOn
    }
}

# Update the farm with zero downtime patching
Write-Host "Updating the farm with zero downtime patching..." -ForegroundColor Yellow
Invoke-Command -ComputerName $SPServers -ScriptBlock {
    # Stop the Distributed Cache service
    Stop-Service AppFabricCachingService

    # Stop the SharePoint Timer service
    Stop-Service SPTimerV4

    # Install the patches
    $Patches = Get-ChildItem -Path $using:PatchPath -Filter *.exe
    foreach ($Patch in $Patches) {
        Write-Host "Installing $($Patch.Name)..." -ForegroundColor Yellow
        Start-Process -FilePath $Patch.FullName -ArgumentList "/quiet /norestart" -Wait
    }

    # Start the SharePoint Timer service
    Start-Service SPTimerV4

    # Start the Distributed Cache service
    Start-Service AppFabricCachingService
}

# Log the patching process
Write-Host "Logging the patching process..." -ForegroundColor Yellow
Invoke-Command -ComputerName $SPServers -ScriptBlock {
    Get-HotFix | Where-Object {$_.Description -like "*SharePoint*"} | Select-Object -Property Description, HotFixID, InstalledOn
} | Out-File $LogFile

# Done
Write-Host "Done." -ForegroundColor Green
