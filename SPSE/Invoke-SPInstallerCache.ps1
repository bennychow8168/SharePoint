#Requires -RunAsAdministrator
<#
.SYNOPSIS
    SharePoint Windows Installer Cache Manager

.DESCRIPTION
    Retrieves SharePoint-related binaries from the Windows Installer Cache
    (C:\Windows\Installer), extracts them to a feed-in location, scans for
    missing entries, and restores binaries back to the cache from a source feed.

.PARAMETER Action
    The operation to perform. Accepted values:

      Export       - Copy SharePoint binaries from Installer Cache to FeedLocation
      Scan         - Detect missing SharePoint binaries; exits 1 if any are found
      Restore      - Scan then restore only missing files from FeedLocation to Cache
      RestoreAll   - Restore ALL feed entries regardless of current cache state
      ScanRestore  - Scan then auto-restore missing files in one shot
      List         - Display all SharePoint products found in the registry

.PARAMETER FeedLocation
    Override the default feed/extract path. Accepts local or UNC paths.
    Default: D:\Binaries\InstallerCache

.PARAMETER Force
    Suppress the restore confirmation prompt. Required for unattended/automated runs.

.PARAMETER ExportReport
    When combined with -Action Scan or ScanRestore, writes a timestamped CSV
    report to FeedLocation.

.PARAMETER LogDirectory
    Directory where the log file is written.
    Log file name is auto-generated as: <ScriptName>_<yyyyMMdd_HHmmss>.log
    Default: C:\Temp

.EXAMPLE
    # Export binaries from Installer Cache to default feed path
    .\Invoke-SPInstallerCache.ps1 -Action Export

.EXAMPLE
    # Export to a custom UNC share
    .\Invoke-SPInstallerCache.ps1 -Action Export -FeedLocation "\\fileserver\SPFeed"

.EXAMPLE
    # Scan and save a CSV report; exits 1 if binaries are missing
    .\Invoke-SPInstallerCache.ps1 -Action Scan -ExportReport

.EXAMPLE
    # Restore only missing files (interactive confirmation)
    .\Invoke-SPInstallerCache.ps1 -Action Restore

.EXAMPLE
    # Restore only missing files silently - ideal for RMM / automation
    .\Invoke-SPInstallerCache.ps1 -Action Restore -Force

.EXAMPLE
    # Full restore of every feed entry, no prompt
    .\Invoke-SPInstallerCache.ps1 -Action RestoreAll -Force

.EXAMPLE
    # One-liner remediation: scan then auto-fix, save report
    .\Invoke-SPInstallerCache.ps1 -Action ScanRestore -Force -ExportReport

.EXAMPLE
    # List all detected SharePoint products with cache status
    .\Invoke-SPInstallerCache.ps1 -Action List

.EXAMPLE
    # Write log to a custom directory
    .\Invoke-SPInstallerCache.ps1 -Action Scan -LogDirectory "D:\Logs\SharePoint"

.EXAMPLE
    # Combine custom log dir with UNC feed and silent restore
    .\Invoke-SPInstallerCache.ps1 -Action ScanRestore -Force -FeedLocation "\\fileserver\SPFeed" -LogDirectory "\\fileserver\Logs"


.NOTES
    Run as Administrator.
    Windows Installer Cache : C:\Windows\Installer
    Registry sources        : HKLM:\SOFTWARE\Classes\Installer\Products
                              HKLM:\SOFTWARE\Classes\Installer\Patches
                              HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [ValidateSet("Export","Scan","Restore","RestoreAll","ScanRestore","List")]
    [string]$Action,

    [Parameter(Mandatory = $false)]
    [string]$FeedLocation = "D:\Binaries\InstallerCache",

    [Parameter(Mandatory = $false)]
    [switch]$Force,

    [Parameter(Mandatory = $false)]
    [switch]$ExportReport,

    [Parameter(Mandatory = $false)]
    [string]$LogDirectory = "C:\Temp"
)

# -----------------------------------------------------------------------------
#  Configuration
# -----------------------------------------------------------------------------
$_scriptName = [IO.Path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Name)
$_logStamp   = Get-Date -Format "yyyyMMdd_HHmmss"
$_logFile    = Join-Path $LogDirectory "${_scriptName}_${_logStamp}.log"

$Config = @{
    InstallerCache  = "C:\Windows\Installer"
    FeedLocation    = $FeedLocation
    LogFile         = $_logFile
    ReportFile      = "$FeedLocation\sp_cache_report.csv"

    ProductRegBase  = "HKLM:\SOFTWARE\Classes\Installer\Products"
    PatchRegBase    = "HKLM:\SOFTWARE\Classes\Installer\Patches"
    UserDataRegBase = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products"

    SharePointKeys  = @(
        "SharePoint",
        "Office Server",
        "Office Web Apps",
        "OSRV",
        "OServer",
        "Project Server",
        "Search Server"
    )
}

# -----------------------------------------------------------------------------
#  Logging
# -----------------------------------------------------------------------------
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("INFO","WARN","ERROR","SUCCESS","HEADER")]
        [string]$Level = "INFO"
    )
    $ts    = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $entry = "[$ts][$Level] $Message"

    $colour = switch ($Level) {
        "INFO"    { "Cyan"    }
        "WARN"    { "Yellow"  }
        "ERROR"   { "Red"     }
        "SUCCESS" { "Green"   }
        "HEADER"  { "Magenta" }
    }
    Write-Host $entry -ForegroundColor $colour

    try {
        $logDir = Split-Path $Config.LogFile
        if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
        $entry | Out-File -FilePath $Config.LogFile -Append -Encoding UTF8
    } catch { }
}

# -----------------------------------------------------------------------------
#  Helper - decode Windows packed GUID to standard GUID
# -----------------------------------------------------------------------------
function ConvertFrom-PackedGuid {
    param([string]$Packed)
    if ($Packed.Length -ne 32) { return $null }
    try {
        $p = $Packed
        $g = "{$($p[7..0] -join '')}-$($p[11..8] -join '')-$($p[15..12] -join '')" +
             "-$($p[17,16,19,18] -join '')-$($p[21,20,23,22,25,24,27,26,29,28,31,30] -join '')"
        return $g.ToUpper()
    } catch { return $null }
}

# -----------------------------------------------------------------------------
#  Registry lookup - returns all SharePoint products and patches with cache paths
# -----------------------------------------------------------------------------
function Get-SharePointProductsFromRegistry {
    $products = [System.Collections.Generic.List[PSCustomObject]]::new()

    # MSI products
    if (Test-Path $Config.ProductRegBase) {
        foreach ($key in Get-ChildItem -Path $Config.ProductRegBase -ErrorAction SilentlyContinue) {
            $props = Get-ItemProperty -Path $key.PSPath -ErrorAction SilentlyContinue
            $name  = $props.ProductName
            if (-not $name) { continue }
            if (-not ($Config.SharePointKeys | Where-Object { $name -like "*$_*" })) { continue }

            $guid     = ConvertFrom-PackedGuid -Packed $key.PSChildName
            $localPkg = $null
            $udPath   = "$($Config.UserDataRegBase)\$($key.PSChildName)\InstallProperties"
            if (Test-Path $udPath) {
                $localPkg = (Get-ItemProperty -Path $udPath -ErrorAction SilentlyContinue).LocalPackage
            }

            $products.Add([PSCustomObject]@{
                PackedGuid   = $key.PSChildName
                ProductGuid  = $guid
                ProductName  = $name
                LocalPackage = $localPkg
                FileExists   = if ($localPkg) { Test-Path $localPkg } else { $false }
                FileSize     = if ($localPkg -and (Test-Path $localPkg)) { (Get-Item $localPkg).Length } else { $null }
                LastModified = if ($localPkg -and (Test-Path $localPkg)) { (Get-Item $localPkg).LastWriteTime } else { $null }
                Type         = if ($localPkg) { [IO.Path]::GetExtension($localPkg).ToUpper().TrimStart('.') } else { "UNKNOWN" }
            })
        }
    }

    # MSP patches
    if (Test-Path $Config.PatchRegBase) {
        foreach ($key in Get-ChildItem -Path $Config.PatchRegBase -ErrorAction SilentlyContinue) {
            $props = Get-ItemProperty -Path $key.PSPath -ErrorAction SilentlyContinue
            $name  = $props.DisplayName
            if (-not $name) { continue }
            if (-not ($Config.SharePointKeys | Where-Object { $name -like "*$_*" })) { continue }

            $guid     = ConvertFrom-PackedGuid -Packed $key.PSChildName
            $localPkg = $props.LocalPackage
            if (-not $localPkg) {
                $mediaPath = "$($key.PSPath)\SourceList\Media"
                if (Test-Path $mediaPath) {
                    $localPkg = (Get-ItemProperty -Path $mediaPath -ErrorAction SilentlyContinue).'1'
                }
            }

            $products.Add([PSCustomObject]@{
                PackedGuid   = $key.PSChildName
                ProductGuid  = $guid
                ProductName  = $name
                LocalPackage = $localPkg
                FileExists   = if ($localPkg) { Test-Path $localPkg } else { $false }
                FileSize     = if ($localPkg -and (Test-Path $localPkg)) { (Get-Item $localPkg).Length } else { $null }
                LastModified = if ($localPkg -and (Test-Path $localPkg)) { (Get-Item $localPkg).LastWriteTime } else { $null }
                Type         = "MSP"
            })
        }
    }

    return $products
}

# -----------------------------------------------------------------------------
#  FUNCTION 1 - Export: Installer Cache to Feed
# -----------------------------------------------------------------------------
function Export-SharePointBinariesToFeed {
    Write-Log "=== EXPORT: Installer Cache -> Feed ===" "HEADER"
    Write-Log "Cache : $($Config.InstallerCache)" "INFO"
    Write-Log "Feed  : $($Config.FeedLocation)"   "INFO"

    if (-not (Test-Path $Config.FeedLocation)) {
        New-Item -ItemType Directory -Path $Config.FeedLocation -Force | Out-Null
        Write-Log "Created feed directory." "INFO"
    }

    $products = Get-SharePointProductsFromRegistry
    if ($products.Count -eq 0) {
        Write-Log "No SharePoint products found in registry." "WARN"
        return
    }

    Write-Log "Found $($products.Count) SharePoint registry entry(s)." "INFO"
    $copied = 0; $skipped = 0; $failed = 0
    $manifest = [System.Collections.Generic.List[PSCustomObject]]::new()

    foreach ($product in $products) {
        if (-not $product.LocalPackage) {
            Write-Log "No cached path for: $($product.ProductName)" "WARN"; $skipped++; continue
        }
        if (-not $product.FileExists) {
            Write-Log "File missing in cache: $($product.LocalPackage) [$($product.ProductName)]" "WARN"; $skipped++; continue
        }

        $safeGuid = if ($product.ProductGuid) { $product.ProductGuid.Trim('{}') } else { $product.PackedGuid }
        $destDir  = Join-Path $Config.FeedLocation "$($product.Type)\$safeGuid"
        $destFile = Join-Path $destDir (Split-Path $product.LocalPackage -Leaf)

        try {
            if (-not (Test-Path $destDir)) { New-Item -ItemType Directory -Path $destDir -Force | Out-Null }
            Copy-Item -Path $product.LocalPackage -Destination $destFile -Force

            $meta = [PSCustomObject]@{
                ProductName       = $product.ProductName
                ProductGuid       = $product.ProductGuid
                PackedGuid        = $product.PackedGuid
                OriginalCachePath = $product.LocalPackage
                FeedFile          = $destFile
                FileSize          = $product.FileSize
                ExportedAt        = (Get-Date -Format "o")
                Type              = $product.Type
            }
            $meta | ConvertTo-Json | Out-File (Join-Path $destDir "metadata.json") -Encoding UTF8 -Force
            $manifest.Add($meta)

            Write-Log "Exported [$($product.Type)] $($product.ProductName)" "SUCCESS"
            Write-Log "         $($product.LocalPackage) -> $destFile" "INFO"
            $copied++
        } catch {
            Write-Log "Failed to export $($product.ProductName): $_" "ERROR"; $failed++
        }
    }

    $manifest | Export-Csv -Path $Config.ReportFile -NoTypeInformation -Encoding UTF8 -Force
    Write-Log "Manifest -> $($Config.ReportFile)" "INFO"
    Write-Log "Export done.  Copied: $copied  |  Skipped: $skipped  |  Failed: $failed" "INFO"
}

# -----------------------------------------------------------------------------
#  FUNCTION 2 - Scan: detect missing binaries in Installer Cache
# -----------------------------------------------------------------------------
function Find-MissingSharePointBinaries {
    param([switch]$ExportReport)

    Write-Log "=== SCAN: Installer Cache Integrity Check ===" "HEADER"

    $products = Get-SharePointProductsFromRegistry
    if ($products.Count -eq 0) {
        Write-Log "No SharePoint products found in registry." "WARN"
        return $null
    }

    Write-Log "Scanning $($products.Count) registry entry(s)..." "INFO"
    $missing = [System.Collections.Generic.List[PSCustomObject]]::new()
    $present = [System.Collections.Generic.List[PSCustomObject]]::new()
    $noPath  = [System.Collections.Generic.List[PSCustomObject]]::new()

    foreach ($p in $products) {
        if (-not $p.LocalPackage) {
            Write-Log "[NO PATH ] $($p.ProductName)" "WARN"; $noPath.Add($p); continue
        }
        if ($p.FileExists) {
            Write-Log "[  OK    ] $($p.ProductName)  ->  $($p.LocalPackage)" "SUCCESS"; $present.Add($p)
        } else {
            Write-Log "[ MISSING] $($p.ProductName)  ->  $($p.LocalPackage)" "ERROR";   $missing.Add($p)
        }
    }

    Write-Host ""
    Write-Host "---------------------------------------------" -ForegroundColor DarkGray
    Write-Host "  Scan Summary" -ForegroundColor Cyan
    Write-Host "---------------------------------------------" -ForegroundColor DarkGray
    Write-Host "  Present  : $($present.Count)" -ForegroundColor Green
    Write-Host "  Missing  : $($missing.Count)" -ForegroundColor $(if ($missing.Count) { "Red" } else { "Green" })
    Write-Host "  No Path  : $($noPath.Count)"  -ForegroundColor Yellow
    Write-Host "---------------------------------------------" -ForegroundColor DarkGray

    if ($missing.Count -gt 0) {
        Write-Host ""
        $missing | Format-Table ProductName, Type, LocalPackage -AutoSize -Wrap
    }

    if ($ExportReport) {
        $rpt = Join-Path (Split-Path $Config.ReportFile) "scan_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        ($present + $missing + $noPath) |
            Select-Object ProductName, ProductGuid, Type, LocalPackage, FileExists,
                @{L="Size_MB"; E={ if ($_.FileSize) { [math]::Round($_.FileSize/1MB,2) } else { $null } }},
                LastModified |
            Export-Csv -Path $rpt -NoTypeInformation -Encoding UTF8
        Write-Log "Scan report -> $rpt" "INFO"
    }

    return @{
        Present  = $present
        Missing  = $missing
        NoPath   = $noPath
        AllClean = ($missing.Count -eq 0)
    }
}

# -----------------------------------------------------------------------------
#  FUNCTION 3 - Restore: Feed to Installer Cache
# -----------------------------------------------------------------------------
function Restore-SharePointBinariesFromFeed {
    param(
        [hashtable]$ScanResult = $null,
        [switch]$Force
    )

    Write-Log "=== RESTORE: Feed -> Installer Cache ===" "HEADER"

    if (-not (Test-Path $Config.FeedLocation)) {
        Write-Log "Feed not found: $($Config.FeedLocation)" "ERROR"; return $false
    }

    $metaFiles = Get-ChildItem -Path $Config.FeedLocation -Recurse -Filter "metadata.json" -ErrorAction SilentlyContinue
    if ($metaFiles.Count -eq 0) {
        Write-Log "No metadata found in feed. Run -Action Export first." "WARN"; return $false
    }

    $allEntries = $metaFiles | ForEach-Object {
        try { Get-Content $_.FullName -Raw | ConvertFrom-Json } catch { $null }
    } | Where-Object { $_ -ne $null }

    $toRestore = if ($ScanResult -and $ScanResult.Missing.Count -gt 0) {
        $missingPaths = $ScanResult.Missing | ForEach-Object { $_.LocalPackage }
        $allEntries   | Where-Object { $_.OriginalCachePath -in $missingPaths }
    } else {
        Write-Log "Evaluating all feed entries against current cache state..." "INFO"
        $allEntries | Where-Object { -not (Test-Path $_.OriginalCachePath) }
    }

    if (@($toRestore).Count -eq 0) {
        Write-Log "Nothing to restore - cache is already complete." "SUCCESS"; return $true
    }

    if (-not $Force) {
        Write-Host ""
        Write-Host "  $(@($toRestore).Count) file(s) will be restored to the Installer Cache:" -ForegroundColor Yellow
        @($toRestore) | ForEach-Object {
            Write-Host "    Product : $($_.ProductName)"       -ForegroundColor White
            Write-Host "    Feed    : $($_.FeedFile)"          -ForegroundColor DarkGray
            Write-Host "    Cache   : $($_.OriginalCachePath)" -ForegroundColor DarkGray
            Write-Host ""
        }
        $confirm = Read-Host "Proceed? (Y/N)"
        if ($confirm -notmatch '^[Yy]') {
            Write-Log "Restore cancelled." "WARN"; return $false
        }
    }

    $success = 0; $failed = 0

    foreach ($entry in @($toRestore)) {
        if (-not (Test-Path $entry.FeedFile)) {
            Write-Log "Feed file missing for '$($entry.ProductName)': $($entry.FeedFile)" "ERROR"; $failed++; continue
        }
        try {
            $cacheDir = Split-Path $entry.OriginalCachePath -Parent
            if (-not (Test-Path $cacheDir)) { New-Item -ItemType Directory -Path $cacheDir -Force | Out-Null }
            Copy-Item -Path $entry.FeedFile -Destination $entry.OriginalCachePath -Force
            Set-Acl -Path $entry.OriginalCachePath -AclObject (Get-Acl $entry.FeedFile) -ErrorAction SilentlyContinue
            Write-Log "Restored: $($entry.ProductName)  ->  $($entry.OriginalCachePath)" "SUCCESS"
            $success++
        } catch {
            Write-Log "Failed to restore '$($entry.ProductName)': $_" "ERROR"; $failed++
        }
    }

    Write-Log "Restore done.  Success: $success  |  Failed: $failed" "INFO"
    return ($failed -eq 0)
}

# -----------------------------------------------------------------------------
#  Entry Point - Parameter dispatcher
# -----------------------------------------------------------------------------
Write-Log "SharePoint Installer Cache Manager  |  Action: $Action" "HEADER"
Write-Log "Feed : $($Config.FeedLocation)" "INFO"

switch ($Action) {

    "Export" {
        Export-SharePointBinariesToFeed
    }

    "Scan" {
        $result = Find-MissingSharePointBinaries -ExportReport:$ExportReport
        if ($result -and -not $result.AllClean) { exit 1 }
    }

    "Restore" {
        Write-Log "Running pre-restore scan..." "INFO"
        $scan = Find-MissingSharePointBinaries
        if ($scan.AllClean) {
            Write-Log "Installer Cache is clean. No restore needed." "SUCCESS"
        } else {
            Restore-SharePointBinariesFromFeed -ScanResult $scan -Force:$Force
        }
    }

    "RestoreAll" {
        Restore-SharePointBinariesFromFeed -Force:$Force
    }

    "ScanRestore" {
        $scan = Find-MissingSharePointBinaries -ExportReport:$ExportReport
        if ($scan.AllClean) {
            Write-Log "Installer Cache is clean. Nothing to restore." "SUCCESS"
        } else {
            Write-Log "$($scan.Missing.Count) missing file(s). Starting restore..." "WARN"
            Restore-SharePointBinariesFromFeed -ScanResult $scan -Force
        }
    }

    "List" {
        $products = Get-SharePointProductsFromRegistry
        if ($products.Count -eq 0) {
            Write-Log "No SharePoint products detected in registry." "WARN"
        } else {
            Write-Host ""
            $products | Sort-Object ProductName | Format-Table `
                ProductName, Type, FileExists,
                @{L="Size (MB)"; E={ if ($_.FileSize) { [math]::Round($_.FileSize/1MB,2) } else { "n/a" } }},
                LastModified -AutoSize -Wrap
        }
    }
}