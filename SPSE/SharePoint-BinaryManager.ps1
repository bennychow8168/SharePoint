#Requires -RunAsAdministrator
<#
.SYNOPSIS
    SharePoint Installer Binary Extraction & Restore Utility

.DESCRIPTION
    Extracts SharePoint install binaries from C:\Windows\Instrallef to D:\Binaries\InstallerBackup.
    Includes detection of missing files and an interactive restore option.

.NOTES
    Run as Administrator
#>

# ─────────────────────────────────────────────
#  Configuration
# ─────────────────────────────────────────────
$Config = @{
    SourcePath      = "C:\Windows\Instrallef"
    DestinationPath = "D:\Binaries\InstallerBackup"
    LogFile         = "D:\Binaries\InstallerBackup\extraction_log.txt"
    # Known SharePoint binary patterns – extend as needed
    BinaryPatterns  = @(
        "*.msi",
        "*.exe",
        "*.cab",
        "*.msp",
        "*.dll"
    )
}

# ─────────────────────────────────────────────
#  Logging Helper
# ─────────────────────────────────────────────
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("INFO","WARN","ERROR","SUCCESS")]
        [string]$Level = "INFO"
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $entry     = "[$timestamp][$Level] $Message"

    # Console colour
    $colour = switch ($Level) {
        "INFO"    { "Cyan"    }
        "WARN"    { "Yellow"  }
        "ERROR"   { "Red"     }
        "SUCCESS" { "Green"   }
    }
    Write-Host $entry -ForegroundColor $colour

    # Append to log file (best-effort)
    try {
        $entry | Out-File -FilePath $Config.LogFile -Append -Encoding UTF8
    } catch { <# silent – log dir may not exist yet #> }
}

# ─────────────────────────────────────────────
#  1. EXTRACT  –  Source → Destination
# ─────────────────────────────────────────────
function Invoke-BinaryExtraction {
    Write-Log "Starting SharePoint binary extraction." "INFO"
    Write-Log "Source      : $($Config.SourcePath)"      "INFO"
    Write-Log "Destination : $($Config.DestinationPath)" "INFO"

    # Validate source
    if (-not (Test-Path $Config.SourcePath)) {
        Write-Log "Source path does not exist: $($Config.SourcePath)" "ERROR"
        return $false
    }

    # Ensure destination exists
    if (-not (Test-Path $Config.DestinationPath)) {
        try {
            New-Item -ItemType Directory -Path $Config.DestinationPath -Force | Out-Null
            Write-Log "Created destination directory: $($Config.DestinationPath)" "INFO"
        } catch {
            Write-Log "Failed to create destination: $_" "ERROR"
            return $false
        }
    }

    # Collect matching files
    $files = @()
    foreach ($pattern in $Config.BinaryPatterns) {
        $files += Get-ChildItem -Path $Config.SourcePath -Filter $pattern -Recurse -ErrorAction SilentlyContinue
    }
    $files = $files | Sort-Object FullName -Unique

    if ($files.Count -eq 0) {
        Write-Log "No SharePoint binaries matched in source path." "WARN"
        return $false
    }

    Write-Log "Found $($files.Count) file(s) to extract." "INFO"

    $successCount = 0
    $failCount    = 0

    foreach ($file in $files) {
        # Mirror relative path under destination
        $relativePath = $file.FullName.Substring($Config.SourcePath.Length).TrimStart('\')
        $destFile     = Join-Path $Config.DestinationPath $relativePath
        $destDir      = Split-Path $destFile -Parent

        try {
            if (-not (Test-Path $destDir)) {
                New-Item -ItemType Directory -Path $destDir -Force | Out-Null
            }
            Copy-Item -Path $file.FullName -Destination $destFile -Force
            Write-Log "Copied: $relativePath" "SUCCESS"
            $successCount++
        } catch {
            Write-Log "Failed to copy '$relativePath': $_" "ERROR"
            $failCount++
        }
    }

    Write-Log "Extraction complete. Success: $successCount | Failed: $failCount" "INFO"
    return ($failCount -eq 0)
}

# ─────────────────────────────────────────────
#  2. DETECT  –  Find missing / corrupt files
# ─────────────────────────────────────────────
function Get-MissingBinaries {
    param(
        [switch]$Verbose
    )

    Write-Log "Scanning for missing or mismatched binaries..." "INFO"

    if (-not (Test-Path $Config.SourcePath)) {
        Write-Log "Source path not accessible: $($Config.SourcePath)" "ERROR"
        return $null
    }

    if (-not (Test-Path $Config.DestinationPath)) {
        Write-Log "Backup destination does not exist – nothing has been extracted yet." "WARN"
        return $null
    }

    # Build source file list
    $sourceFiles = @()
    foreach ($pattern in $Config.BinaryPatterns) {
        $sourceFiles += Get-ChildItem -Path $Config.SourcePath -Filter $pattern -Recurse -ErrorAction SilentlyContinue
    }
    $sourceFiles = $sourceFiles | Sort-Object FullName -Unique

    $missing   = [System.Collections.Generic.List[PSCustomObject]]::new()
    $mismatched = [System.Collections.Generic.List[PSCustomObject]]::new()

    foreach ($src in $sourceFiles) {
        $relativePath = $src.FullName.Substring($Config.SourcePath.Length).TrimStart('\')
        $destFile     = Join-Path $Config.DestinationPath $relativePath

        if (-not (Test-Path $destFile)) {
            $missing.Add([PSCustomObject]@{
                Status       = "MISSING"
                RelativePath = $relativePath
                SourceFile   = $src.FullName
                DestFile     = $destFile
                SourceSize   = $src.Length
                DestSize     = $null
            })
            if ($Verbose) { Write-Log "MISSING  : $relativePath" "WARN" }
        } else {
            $dest = Get-Item $destFile
            if ($src.Length -ne $dest.Length) {
                $mismatched.Add([PSCustomObject]@{
                    Status       = "SIZE_MISMATCH"
                    RelativePath = $relativePath
                    SourceFile   = $src.FullName
                    DestFile     = $destFile
                    SourceSize   = $src.Length
                    DestSize     = $dest.Length
                })
                if ($Verbose) { Write-Log "MISMATCH : $relativePath (src=$($src.Length)B dest=$($dest.Length)B)" "WARN" }
            }
        }
    }

    $result = @{
        Missing    = $missing
        Mismatched = $mismatched
        Total      = $missing.Count + $mismatched.Count
    }

    if ($result.Total -eq 0) {
        Write-Log "All binaries are present and sizes match." "SUCCESS"
    } else {
        Write-Log "Issues found – Missing: $($missing.Count) | Size mismatches: $($mismatched.Count)" "WARN"
    }

    return $result
}

# ─────────────────────────────────────────────
#  3. RESTORE  –  Copy backup → source
# ─────────────────────────────────────────────
function Invoke-BinaryRestore {
    param(
        # Optionally pass the output of Get-MissingBinaries to restore only affected files
        [hashtable]$DetectionResult = $null,
        [switch]$Force
    )

    Write-Log "Starting restore operation." "INFO"

    if (-not (Test-Path $Config.DestinationPath)) {
        Write-Log "Backup path does not exist: $($Config.DestinationPath)" "ERROR"
        return $false
    }

    # Determine which files to restore
    if ($null -ne $DetectionResult -and $DetectionResult.Total -gt 0) {
        $filesToRestore = @($DetectionResult.Missing) + @($DetectionResult.Mismatched)
        Write-Log "Restoring $($filesToRestore.Count) affected file(s) only." "INFO"
    } else {
        # Full restore – all files in backup
        Write-Log "No detection result supplied. Performing FULL restore." "WARN"
        $backupFiles    = Get-ChildItem -Path $Config.DestinationPath -Recurse -File -ErrorAction SilentlyContinue
        $filesToRestore = $backupFiles | ForEach-Object {
            $rel  = $_.FullName.Substring($Config.DestinationPath.Length).TrimStart('\')
            [PSCustomObject]@{
                RelativePath = $rel
                SourceFile   = Join-Path $Config.SourcePath $rel
                DestFile     = $_.FullName
            }
        }
    }

    if ($filesToRestore.Count -eq 0) {
        Write-Log "Nothing to restore." "INFO"
        return $true
    }

    # Confirm unless -Force
    if (-not $Force) {
        Write-Host ""
        Write-Host "The following $($filesToRestore.Count) file(s) will be restored to '$($Config.SourcePath)':" -ForegroundColor Yellow
        $filesToRestore | ForEach-Object { Write-Host "  → $($_.RelativePath)" -ForegroundColor White }
        Write-Host ""
        $confirm = Read-Host "Proceed with restore? (Y/N)"
        if ($confirm -notmatch '^[Yy]') {
            Write-Log "Restore cancelled by user." "WARN"
            return $false
        }
    }

    $successCount = 0
    $failCount    = 0

    foreach ($item in $filesToRestore) {
        $srcFile = if ($item.PSObject.Properties['DestFile']) { $item.DestFile } else { $null }
        $dstFile = if ($item.PSObject.Properties['SourceFile']) { $item.SourceFile } else { $null }

        # For full-restore objects the mapping is already set correctly
        if ($null -eq $srcFile -or $null -eq $dstFile) { continue }

        $dstDir = Split-Path $dstFile -Parent
        try {
            if (-not (Test-Path $dstDir)) {
                New-Item -ItemType Directory -Path $dstDir -Force | Out-Null
            }
            Copy-Item -Path $srcFile -Destination $dstFile -Force
            Write-Log "Restored: $($item.RelativePath)" "SUCCESS"
            $successCount++
        } catch {
            Write-Log "Failed to restore '$($item.RelativePath)': $_" "ERROR"
            $failCount++
        }
    }

    Write-Log "Restore complete. Success: $successCount | Failed: $failCount" "INFO"
    return ($failCount -eq 0)
}

# ─────────────────────────────────────────────
#  Interactive Menu
# ─────────────────────────────────────────────
function Show-Menu {
    while ($true) {
        Write-Host ""
        Write-Host "╔══════════════════════════════════════════════╗" -ForegroundColor Cyan
        Write-Host "║   SharePoint Binary Manager                  ║" -ForegroundColor Cyan
        Write-Host "╠══════════════════════════════════════════════╣" -ForegroundColor Cyan
        Write-Host "║  1. Extract binaries (Source → Backup)       ║" -ForegroundColor White
        Write-Host "║  2. Detect missing / mismatched files        ║" -ForegroundColor White
        Write-Host "║  3. Restore missing files only               ║" -ForegroundColor White
        Write-Host "║  4. Full restore (Backup → Source)           ║" -ForegroundColor White
        Write-Host "║  5. View log                                 ║" -ForegroundColor White
        Write-Host "║  Q. Quit                                     ║" -ForegroundColor White
        Write-Host "╚══════════════════════════════════════════════╝" -ForegroundColor Cyan
        Write-Host ""

        $choice = Read-Host "Select an option"

        switch ($choice.ToUpper()) {
            "1" {
                Invoke-BinaryExtraction
            }
            "2" {
                $result = Get-MissingBinaries -Verbose
                if ($null -ne $result -and $result.Total -gt 0) {
                    Write-Host ""
                    $result.Missing    | Format-Table Status, RelativePath, SourceSize -AutoSize
                    $result.Mismatched | Format-Table Status, RelativePath, SourceSize, DestSize -AutoSize
                }
            }
            "3" {
                $result = Get-MissingBinaries -Verbose
                if ($null -ne $result -and $result.Total -gt 0) {
                    Invoke-BinaryRestore -DetectionResult $result
                } else {
                    Write-Log "No missing files detected. Restore not needed." "SUCCESS"
                }
            }
            "4" {
                Invoke-BinaryRestore -Force:$false
            }
            "5" {
                if (Test-Path $Config.LogFile) {
                    Get-Content $Config.LogFile | Select-Object -Last 50 | Write-Host
                } else {
                    Write-Log "Log file not found." "WARN"
                }
            }
            "Q" {
                Write-Log "Exiting SharePoint Binary Manager." "INFO"
                return
            }
            default {
                Write-Host "Invalid option. Please try again." -ForegroundColor Red
            }
        }
    }
}

# ─────────────────────────────────────────────
#  Entry Point
# ─────────────────────────────────────────────
Show-Menu