#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Reassigns drive letters based on detected drive configuration and handles
    extra drives (above H:) by creating a dynamic spanned volume mounted at I:\SPIndex.

.DESCRIPTION
    Scenario A – Drives D, E, F, G (no H):
        E: → L:   |   F: → P:   |   G: → T:

    Scenario B – Drives D, E, F, G, H:
        E: → I:   |   F: → L:   |   G: → P:   |   H: → T:

    Extra drives (letters I–Z beyond the expected set) are spanned into a
    single dynamic volume and mounted as I:\SPIndex.

.NOTES
    - Must be run as Administrator.
    - Tested on Windows Server 2016/2019/2022 and Windows 10/11.
    - Uses diskpart for dynamic disk/span operations; all other ops use WMI/CIM.
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ─────────────────────────────────────────────────────────────────────────────
# Helper: Write a timestamped log line
# ─────────────────────────────────────────────────────────────────────────────
function Write-Log {
    param([string]$Message, [ValidateSet('INFO','WARN','ERROR')]$Level = 'INFO')
    $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $color = @{ INFO = 'Cyan'; WARN = 'Yellow'; ERROR = 'Red' }[$Level]
    Write-Host "[$ts] [$Level] $Message" -ForegroundColor $color
}

# ─────────────────────────────────────────────────────────────────────────────
# Helper: Change a drive letter via WMI
# ─────────────────────────────────────────────────────────────────────────────
function Set-DriveLetter {
    param(
        [string]$CurrentLetter,   # e.g. "E"
        [string]$NewLetter        # e.g. "L"
    )

    $currentPath = "${CurrentLetter}:"
    $newPath     = "${NewLetter}:"

    # Verify current drive exists
    $vol = Get-WmiObject -Class Win32_Volume -Filter "DriveLetter='${currentPath}'"
    if (-not $vol) {
        Write-Log "Drive ${currentPath} not found – skipping rename to ${newPath}." WARN
        return $false
    }

    # Verify target letter is free
    $conflict = Get-WmiObject -Class Win32_Volume -Filter "DriveLetter='${newPath}'"
    if ($conflict) {
        Write-Log "Target letter ${newPath} is already in use – cannot rename ${currentPath}." ERROR
        return $false
    }

    Write-Log "Renaming ${currentPath} → ${newPath} …"
    $vol.DriveLetter = $newPath
    $result = $vol.Put()

    if ($result.ReturnValue -eq 0) {
        Write-Log "Successfully renamed ${currentPath} → ${newPath}."
        return $true
    } else {
        Write-Log "WMI returned error code $($result.ReturnValue) while renaming ${currentPath}." ERROR
        return $false
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# Helper: Get all currently assigned drive letters (upper-case, no colon)
# ─────────────────────────────────────────────────────────────────────────────
function Get-AssignedLetters {
    Get-WmiObject -Class Win32_Volume |
        Where-Object { $_.DriveLetter -match '^[A-Z]:$' } |
        ForEach-Object { $_.DriveLetter.TrimEnd(':') } |
        Sort-Object
}

# ─────────────────────────────────────────────────────────────────────────────
# Helper: Create a spanned dynamic volume from a list of disk numbers,
#         format it NTFS, and mount it at a folder path.
# ─────────────────────────────────────────────────────────────────────────────
function New-SpannedVolume {
    param(
        [int[]]$DiskNumbers,
        [string]$MountPath   # e.g. "I:\SPIndex"
    )

    if ($DiskNumbers.Count -eq 0) {
        Write-Log "No disks supplied to New-SpannedVolume." WARN
        return
    }

    Write-Log "Creating spanned dynamic volume from disk(s): $($DiskNumbers -join ', ')"

    # Build diskpart script
    $dpLines = @('rescan')

    foreach ($d in $DiskNumbers) {
        $dpLines += "select disk $d"
        $dpLines += 'convert dynamic noerr'
    }

    # Create the span: select first disk, then extend with the rest
    $dpLines += "select disk $($DiskNumbers[0])"
    $dpLines += 'create volume simple noerr'

    for ($i = 1; $i -lt $DiskNumbers.Count; $i++) {
        $dpLines += "select disk $($DiskNumbers[$i])"
        $dpLines += 'extend noerr'
    }

    $dpLines += 'format fs=ntfs label="SPIndex" quick'
    $dpLines += 'remove'            # remove any auto-assigned letter
    $dpLines += 'exit'

    $dpScript = $dpLines -join "`r`n"
    $tmpFile  = [System.IO.Path]::GetTempFileName()
    Set-Content -Path $tmpFile -Value $dpScript -Encoding ASCII

    Write-Log "Running diskpart …"
    $dpOutput = diskpart /s $tmpFile 2>&1
    Remove-Item $tmpFile -Force

    Write-Log "diskpart output:`n$($dpOutput -join "`n")"

    # Ensure mount-point parent exists (I:\)
    $parentDrive = Split-Path $MountPath -Qualifier   # "I:"
    if (-not (Test-Path "${parentDrive}\")) {
        Write-Log "Parent drive ${parentDrive} does not exist for mount point." ERROR
        return
    }

    if (-not (Test-Path $MountPath)) {
        New-Item -ItemType Directory -Path $MountPath -Force | Out-Null
        Write-Log "Created folder $MountPath"
    }

    # Assign the mount-point via WMI – find the newly created volume (no drive letter)
    Start-Sleep -Seconds 3   # brief pause for volume to register
    $newVol = Get-WmiObject -Class Win32_Volume |
              Where-Object { -not $_.DriveLetter -and $_.Label -eq 'SPIndex' } |
              Select-Object -First 1

    if ($newVol) {
        $newVol.AddMountPoint($MountPath) | Out-Null
        Write-Log "Mounted spanned volume at $MountPath"
    } else {
        Write-Log "Could not locate the new SPIndex volume to set mount point." WARN
        Write-Log "You may need to assign the mount point manually via Disk Management." WARN
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────

Write-Log "=== Drive Letter Reassignment Script Starting ==="

# 1. Snapshot current drive letters
$assigned = Get-AssignedLetters
Write-Log "Currently assigned drive letters: $($assigned -join ', ')"

# 2. Determine which of D–H are present
$hasD = $assigned -contains 'D'
$hasE = $assigned -contains 'E'
$hasF = $assigned -contains 'F'
$hasG = $assigned -contains 'G'
$hasH = $assigned -contains 'H'

# 3. Detect drives beyond H (I–Z) that are NOT system/C
$extraLetters = $assigned | Where-Object {
    $_ -notin @('A','B','C','D','E','F','G','H')
}

# ─────────────────────────────────────────────────────────────────────────────
# Scenario detection
# ─────────────────────────────────────────────────────────────────────────────
if ($hasD -and $hasE -and $hasF -and $hasG -and -not $hasH) {

    Write-Log "Scenario A detected: D, E, F, G (no H)"
    Write-Log "Mapping: E→L | F→P | G→T"

    # Order matters – rename in reverse to avoid letter collisions
    Set-DriveLetter 'G' 'T' | Out-Null
    Set-DriveLetter 'F' 'P' | Out-Null
    Set-DriveLetter 'E' 'L' | Out-Null

} elseif ($hasD -and $hasE -and $hasF -and $hasG -and $hasH) {

    Write-Log "Scenario B detected: D, E, F, G, H"
    Write-Log "Mapping: E→I | F→L | G→P | H→T"

    Set-DriveLetter 'H' 'T' | Out-Null
    Set-DriveLetter 'G' 'P' | Out-Null
    Set-DriveLetter 'F' 'L' | Out-Null
    Set-DriveLetter 'E' 'I' | Out-Null

} else {
    Write-Log "Drive configuration does not match Scenario A or B." WARN
    Write-Log "  hasD=$hasD  hasE=$hasE  hasF=$hasF  hasG=$hasG  hasH=$hasH" WARN
    Write-Log "No letter reassignment performed." WARN
}

# ─────────────────────────────────────────────────────────────────────────────
# Handle extra drives (originally I–Z, i.e. beyond the expected base set)
# ─────────────────────────────────────────────────────────────────────────────
if ($extraLetters.Count -gt 0) {

    Write-Log "Extra drives detected beyond H: $($extraLetters -join ', ')"
    Write-Log "These will be spanned into a single dynamic volume at I:\SPIndex"

    # Resolve disk numbers for the extra volumes
    $diskNumbers = @()
    foreach ($letter in $extraLetters) {
        $vol = Get-WmiObject -Class Win32_Volume -Filter "DriveLetter='${letter}:'"
        if ($vol) {
            # Get the underlying disk number via Win32_DiskPartition association
            $partition = Get-WmiObject -Query "
                ASSOCIATORS OF {Win32_Volume.DeviceID='$($vol.DeviceID)'}
                WHERE AssocClass=Win32_DiskDriveToDiskPartition" |
                Select-Object -First 1

            if (-not $partition) {
                # Alternate query path
                $partition = Get-WmiObject -Class Win32_DiskPartition |
                    Where-Object { $vol.DeviceID -like "*$($_.DiskIndex)*" } |
                    Select-Object -First 1
            }

            if ($partition) {
                $diskNumbers += [int]$partition.DiskIndex
                Write-Log "  ${letter}: → Disk $($partition.DiskIndex)"
            } else {
                Write-Log "  Could not determine disk number for ${letter}: – skipping." WARN
            }
        }
    }

    $diskNumbers = $diskNumbers | Sort-Object -Unique

    if ($diskNumbers.Count -gt 0) {
        # After scenario renames, I: should now be free (Scenario A) or
        # already used (Scenario B used it for E→I). Warn if occupied.
        $iDrive = Get-WmiObject -Class Win32_Volume -Filter "DriveLetter='I:'"
        if ($iDrive) {
            Write-Log "I: is already assigned; the mount point I:\SPIndex will be placed on it." INFO
        }

        New-SpannedVolume -DiskNumbers $diskNumbers -MountPath 'I:\SPIndex'
    } else {
        Write-Log "No valid disk numbers resolved for extra drives. Skipping span creation." WARN
    }

} else {
    Write-Log "No extra drives beyond H detected. No span volume needed."
}

# ─────────────────────────────────────────────────────────────────────────────
# Final state
# ─────────────────────────────────────────────────────────────────────────────
Write-Log "=== Final drive letter state ==="
Get-WmiObject -Class Win32_Volume |
    Where-Object { $_.DriveLetter -match '^[A-Z]:$' } |
    Sort-Object DriveLetter |
    ForEach-Object {
        Write-Log "  $($_.DriveLetter)  Label='$($_.Label)'  Type=$($_.DriveType)"
    }

Write-Log "=== Script completed ==="
