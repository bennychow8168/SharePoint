[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)] [string] $backupDir
)

# Define Backup Location
#$backupDir = "C:\SharePointInstallerBackup"
if (!(Test-Path $backupDir)) { New-Item -ItemType Directory -Path $backupDir }

# Get all MSP/MSI files in Windows Installer
$installerFiles = Get-ChildItem -Path "C:\Windows\Installer" -Include *.msp, *.msi -Recurse

foreach ($file in $installerFiles) {
    # Check if the file contains "SharePoint" (can be tweaked for specific product names)
    if (Get-ItemProperty -Path $file.FullName | Select-Object -ExpandProperty VersionString -ErrorAction SilentlyContinue | Select-String -Pattern "SharePoint") {
        Copy-Item -Path $file.FullName -Destination $backupDir -Force
        Write-Host "Backed up: $($file.Name)" -ForegroundColor Green
    }
}
Write-Host "Backup Complete." -ForegroundColor Cyan
