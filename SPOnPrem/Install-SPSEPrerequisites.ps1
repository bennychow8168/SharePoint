[CmdletBinding()]
param (
    [Parameter(Mandatory = $true, HelpMessage = "SPSE Build version, e.g. 16.0.17928.20000")] [ValidateNotNullOrEmpty()] [string]$SPSEBuild,
    [Parameter(Mandatory = $true, HelpMessage = "Cumulative Update KB number, e.g. KB5002640")] [ValidatePattern('^KB\d+$')] [string]$CU,
    [Parameter(Mandatory = $false, HelpMessage = "Log directory path. Defaults to C:\Temp")] [string]$LogDirectory = "C:\Temp",
    [Parameter(Mandatory = $false)] [switch]$Force,
    [Parameter(Mandatory = $false)] [ValidateRange(1, 6)] [int]$Stage = 1
)

# ──────────────────────────────────────────────
#  Configuration
# ──────────────────────────────────────────────
$RegistryPath = "HKLM:\SOFTWARE\SharePoint"
$Name         = "StageValue"   # <-- replace with your actual registry value name
$Value        = $null

# ──────────────────────────────────────────────
#  Stage functions  (replace bodies as needed)
# ──────────────────────────────────────────────
function Stage1 { 
    Write-Host "[Stage 1] Running Stage 1..." -ForegroundColor Cyan    
    Set-RegistryValue -Name "BuildVersion" -Value $SPSEBuild
    Set-RegistryValue -Name "KB" -Value $CU
    Set-RegistryValue -Name "StageValue"   -Value "2"
}
function Stage2 { 
    Write-Host "[Stage 2] Running Stage 2..." -ForegroundColor Green
    Set-RegistryValue -Name "StageValue"   -Value "3"
}
function Stage3 { 
    Write-Host "[Stage 3] Running Stage 3..." -ForegroundColor Yellow  
    Set-RegistryValue -Name "StageValue"   -Value "4"
}
function Stage4 { 
    Write-Host "[Stage 4] Running Stage 4..." -ForegroundColor Magenta
    Set-RegistryValue -Name "StageValue"   -Value "5"
}
function Stage5 { 
    Write-Host "[Stage 5] Running Stage 5..." -ForegroundColor Blue    
    Set-RegistryValue -Name "StageValue"   -Value "6"
}
function Stage6 { 
    Write-Host "[Stage 6] Running Stage 6..." -ForegroundColor Red     
}

# ──────────────────────────────────────────────
#  Dispatch helper
# ──────────────────────────────────────────────
function Write-LogEntry {
    param (
        [string]$Message,
        [ValidateSet("INFO","SUCCESS","WARNING","ERROR","SECTION")]
        [string]$Level = "INFO"
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $colors    = @{
        INFO    = "White"
        SUCCESS = "Green"
        WARNING = "Yellow"
        ERROR   = "Red"
        SECTION = "Cyan"
    }
    $prefix = switch ($Level) {
        "INFO"    { "[INFO   ]" }
        "SUCCESS" { "[SUCCESS]" }
        "WARNING" { "[WARNING]" }
        "ERROR"   { "[ERROR  ]" }
        "SECTION" { "[SECTION]" }
    }

    $logLine = "$timestamp $prefix $Message"
    $logLine | Out-File -FilePath $LogFilePath -Append -Encoding UTF8
    Write-Host $logLine -ForegroundColor $colors[$Level]
}
function Set-RegistryValue {
    param (
        [Parameter(Mandatory = $true)] [string]$Name,
        [Parameter(Mandatory = $true)] [string]$Value
    )

    $RegistryPath = "HKLM:\SOFTWARE\ABC\SharePoint"

    # Create the registry key path if it doesn't exist
    if (-not (Test-Path -Path $RegistryPath)) {
        Write-Host "Registry path '$RegistryPath' not found. Creating..."
        New-Item -Path $RegistryPath -Force | Out-Null
        Write-Host "Registry path created."
    }

    # Check if the registry value exists
    $existingValue = Get-ItemProperty -Path $RegistryPath -Name $Name -ErrorAction SilentlyContinue

    if ($null -eq $existingValue) {
        # Value does not exist — create it
        Write-Host "Registry value '$Name' not found. Creating with value '$Value'..."
        New-ItemProperty -Path $RegistryPath -Name $Name -Value $Value -PropertyType String -Force | Out-Null
        Write-Host "Registry value '$Name' created successfully."
    } else {
        # Value already exists
        Write-Host "Registry value '$Name' already exists with value '$($existingValue.$Name)'. No changes made."
    }
}

function Invoke-Stage {
    param([int]$StageNumber)

    switch ($StageNumber) {
        1 { Stage1 }
        2 { Stage2 }
        3 { Stage3 }
        4 { Stage4 }
        5 { Stage5 }
        6 { Stage6 }
        default {
            Write-Warning "Unknown stage number: $StageNumber. Valid values are 1–6."
            exit 1
        }
    }
}


$ScriptName   = [System.IO.Path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Name)
$LogDate = Get-Date -Format "yyyyMMdd-HHmmss"
$LogFileName  = "$ScriptName-$LogDate.log"
$LogFilePath  = Join-Path -Path $LogDirectory -ChildPath $LogFileName

# ──────────────────────────────────────────────
#  -Force -Stage <n>  →  bypass registry, run stage directly
# ──────────────────────────────────────────────
if ($Force) {
    if ($Stage -eq 0) {
        Write-Error "-Force requires -Stage <1-6> to be specified."
        exit 1
    }
    Write-Host "[Force] Skipping registry check. Forcing Stage $Stage." -ForegroundColor DarkYellow
    Invoke-Stage -StageNumber $Stage
    exit 0
}

# ──────────────────────────────────────────────
#  Registry detection
# ──────────────────────────────────────────────
if (-not (Test-Path -Path $RegistryPath)) {
    Write-Error "Registry path not found: $RegistryPath"
    exit 1
}

try {
    $regItem = Get-ItemProperty -Path $RegistryPath -Name $Name -ErrorAction Stop
    $Value   = $regItem.$Name
}
catch {
    Write-Error "Registry value '$Name' not found under '$RegistryPath'."
    exit 1
}

Write-Host "Registry key  : $RegistryPath" -ForegroundColor Gray
Write-Host "Value name    : $Name"          -ForegroundColor Gray
Write-Host "Value detected: $Value"         -ForegroundColor Gray

# ──────────────────────────────────────────────
#  Dispatch based on registry value
# ──────────────────────────────────────────────
if ($Value -ge 1 -and $Value -le 6) {
    Invoke-Stage -StageNumber $Value
}
else {
    Write-Warning "Registry value '$Value' is outside the expected range (1–6). No stage executed."
    exit 1
}
