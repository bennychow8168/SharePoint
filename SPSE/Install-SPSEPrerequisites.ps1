#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Prepares SharePoint Server Subscription Edition (SPSE) prerequisites and installation.

.DESCRIPTION
    This script automates the full setup process for SPSE including:
    - Folder structure creation
    - Environment variable configuration (TEMP/TMP)
    - Page file configuration
    - IIS and Windows Features installation
    - Binary retrieval and extraction
    - Prerequisites installation
    - SharePoint application installation
    - Japanese Language Pack installation
    - Cumulative Update installation

.PARAMETER SPSEBuild
    The SPSE build version string (e.g., "16.0.17928.20000").

.PARAMETER CU
    The Cumulative Update KB number (e.g., "KB5002640").

.PARAMETER LogDirectory
    The directory where the log file will be saved. Defaults to C:\Temp.

.EXAMPLE
    .\Prepare-SPSEPrerequisites.ps1 -SPSEBuild "16.0.17928.20000" -CU "KB5002640"

.EXAMPLE
    .\Prepare-SPSEPrerequisites.ps1 -SPSEBuild "16.0.17928.20000" -CU "KB5002640" -LogDirectory "D:\Logs"

.NOTES
    - Must be run as Administrator.
    - Requires P: and T: drive availability for page file and temp redirection.
    - Internet/network access to http://abd.com is required for binary downloads.
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true, HelpMessage = "SPSE Build version, e.g. 16.0.17928.20000")]
    [ValidateNotNullOrEmpty()]
    [string]$SPSEBuild,

    [Parameter(Mandatory = $true, HelpMessage = "Cumulative Update KB number, e.g. KB5002640")]
    [ValidatePattern('^KB\d+$')]
    [string]$CU,

    [Parameter(Mandatory = $false, HelpMessage = "Log directory path. Defaults to C:\Temp")]
    [string]$LogDirectory = "C:\Temp"
)

###############################################################################
#region --- Logging Functions
###############################################################################

$ScriptName   = [System.IO.Path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Name)
$LogTimestamp = Get-Date -Format "yyyyMMdd-HHmmss"
$LogFileName  = "$ScriptName-$LogTimestamp.log"
$LogFilePath  = Join-Path -Path $LogDirectory -ChildPath $LogFileName

function Initialize-Log {
    if (-not (Test-Path -Path $LogDirectory)) {
        New-Item -ItemType Directory -Path $LogDirectory -Force | Out-Null
    }
    $header = @"
================================================================================
  SPSE Prerequisites Setup Log
  Script   : $ScriptName
  Build    : $SPSEBuild
  CU       : $CU
  Started  : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
  Log File : $LogFilePath
================================================================================
"@
    $header | Out-File -FilePath $LogFilePath -Encoding UTF8
    Write-Host $header -ForegroundColor Cyan
}

function Write-Log {
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

function Write-Section {
    param([string]$Title)
    $separator = "=" * 80
    $line      = "`n$separator`n  $Title`n$separator"
    $line | Out-File -FilePath $LogFilePath -Append -Encoding UTF8
    Write-Host $line -ForegroundColor Cyan
}

#endregion

###############################################################################
#region --- Helper Functions
###############################################################################

function Test-DriveAvailability {
    param([string]$DriveLetter)
    $letter = $DriveLetter.TrimEnd(':\/')
    return (Get-PSDrive -Name $letter -ErrorAction SilentlyContinue) -ne $null
}

function New-DirectoryIfNotExist {
    param([string]$Path)
    if (-not (Test-Path -Path $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
        Write-Log "Created directory: $Path" -Level SUCCESS
    } else {
        Write-Log "Directory already exists: $Path" -Level INFO
    }
}

function Invoke-FileDownload {
    param (
        [string]$Url,
        [string]$Destination
    )
    Write-Log "Downloading: $Url" -Level INFO
    Write-Log "Destination: $Destination" -Level INFO
    try {
        $webClient = New-Object System.Net.WebClient
        $webClient.DownloadFile($Url, $Destination)
        Write-Log "Download complete: $([System.IO.Path]::GetFileName($Destination))" -Level SUCCESS
    } catch {
        Write-Log "Failed to download $Url -- $_" -Level ERROR
        throw
    }
}

function Expand-ZipArchive {
    param (
        [string]$ZipPath,
        [string]$Destination
    )
    Write-Log "Extracting: $ZipPath => $Destination" -Level INFO
    try {
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        [System.IO.Compression.ZipFile]::ExtractToDirectory($ZipPath, $Destination)
        Write-Log "Extraction complete." -Level SUCCESS
    } catch {
        Write-Log "Extraction failed for $ZipPath -- $_" -Level ERROR
        throw
    }
}

function Invoke-ExternalProcess {
    param (
        [string]$FilePath,
        [string]$ArgumentList,
        [string]$Description,
        [int[]]$SuccessExitCodes = @(0)
    )
    Write-Log "Executing [$Description]: $FilePath $ArgumentList" -Level INFO
    try {
        $psi                        = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName               = $FilePath
        $psi.Arguments              = $ArgumentList
        $psi.UseShellExecute        = $false
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError  = $true

        $process  = [System.Diagnostics.Process]::Start($psi)
        $stdout   = $process.StandardOutput.ReadToEnd()
        $stderr   = $process.StandardError.ReadToEnd()
        $process.WaitForExit()
        $exitCode = $process.ExitCode

        if ($stdout) { Write-Log "STDOUT: $stdout" -Level INFO }
        if ($stderr) { Write-Log "STDERR: $stderr" -Level WARNING }

        if ($exitCode -in $SuccessExitCodes) {
            Write-Log "[$Description] completed with exit code $exitCode." -Level SUCCESS
        } else {
            Write-Log "[$Description] returned unexpected exit code $exitCode." -Level ERROR
            throw "Process '$FilePath' exited with code $exitCode."
        }
    } catch {
        Write-Log "Failed to execute [$Description] -- $_" -Level ERROR
        throw
    }
}

#endregion

###############################################################################
#region --- Step Functions
###############################################################################

# --- Step 1: Create Folder Structure -------------------------------------------
function Initialize-FolderStructure {
    Write-Section "STEP 1: Create Folder Structure"

    $folders = @(
        "D:\Binaries",
        "D:\Jobs",
        "D:\Scripts",
        "D:\Binaries\SP\SE\LanguagePacks",
        "D:\Binaries\SP\SE\SharePoint",
        "D:\Binaries\SP\SE\Updates",
        "D:\Binaries\SP\Automation",
        "D:\Binaries\SP\PrerequisitesInstaller"
    )

    foreach ($folder in $folders) {
        New-DirectoryIfNotExist -Path $folder
    }
}

# --- Step 2: Configure TEMP / TMP on T: Drive ----------------------------------
function Set-TempEnvironment {
    Write-Section "STEP 2: Configure TEMP/TMP to T:\Temp"

    $tDrive = "T"

    if (-not (Test-DriveAvailability -DriveLetter $tDrive)) {
        Write-Log "T: drive is NOT available. TEMP/TMP will not be redirected." -Level ERROR
        throw "T: drive is required for TEMP/TMP redirection but is not available."
    }

    Write-Log "T: drive detected." -Level SUCCESS
    $tempPath = "T:\Temp"
    New-DirectoryIfNotExist -Path $tempPath

    Write-Log "Setting Machine-level TEMP and TMP to $tempPath" -Level INFO
    [System.Environment]::SetEnvironmentVariable("TEMP", $tempPath, [System.EnvironmentVariableTarget]::Machine)
    [System.Environment]::SetEnvironmentVariable("TMP",  $tempPath, [System.EnvironmentVariableTarget]::Machine)

    Write-Log "Setting User-level TEMP and TMP to $tempPath" -Level INFO
    [System.Environment]::SetEnvironmentVariable("TEMP", $tempPath, [System.EnvironmentVariableTarget]::User)
    [System.Environment]::SetEnvironmentVariable("TMP",  $tempPath, [System.EnvironmentVariableTarget]::User)

    # Also update the current process environment
    $env:TEMP = $tempPath
    $env:TMP  = $tempPath

    Write-Log "TEMP/TMP successfully configured to $tempPath" -Level SUCCESS
}

# --- Step 3: Configure Page File on P: Drive -----------------------------------
function Set-PageFile {
    Write-Section "STEP 3: Configure Page File on P:\"

    $pDrive = "P"

    if (-not (Test-DriveAvailability -DriveLetter $pDrive)) {
        Write-Log "P: drive is NOT available. Page file cannot be configured." -Level ERROR
        throw "P: drive is required for page file configuration but is not available."
    }

    Write-Log "P: drive detected." -Level SUCCESS

    # Calculate page file size: 1.5 x physical RAM in MB
    $ramMB        = [math]::Round((Get-CimInstance Win32_ComputerSystem).TotalPhysicalMemory / 1MB)
    $pageFileMB   = [math]::Round($ramMB * 1.5)
    $pageFilePath = "P:\pagefile.sys"

    Write-Log "Physical RAM   : $ramMB MB"         -Level INFO
    Write-Log "Page File Size : $pageFileMB MB (1.5x RAM)" -Level INFO
    Write-Log "Page File Path : $pageFilePath"     -Level INFO

    # Disable automatic page file management
    $cs = Get-CimInstance -ClassName Win32_ComputerSystem
    if ($cs.AutomaticManagedPagefile) {
        Set-CimInstance -InputObject $cs -Property @{ AutomaticManagedPagefile = $false }
        Write-Log "Disabled automatic managed page file." -Level INFO
    }

    # Remove any existing page files on all drives
    $existingPF = Get-CimInstance -ClassName Win32_PageFileSetting
    if ($existingPF) {
        foreach ($pf in $existingPF) {
            Write-Log "Removing existing page file: $($pf.Name)" -Level INFO
            Remove-CimInstance -InputObject $pf
        }
    }

    # Create new page file on P:
    New-CimInstance -ClassName Win32_PageFileSetting -Property @{
        Name        = $pageFilePath
        InitialSize = $pageFileMB
        MaximumSize = $pageFileMB
    } | Out-Null

    Write-Log "Page file configured: $pageFilePath -- Initial/Max: $pageFileMB MB" -Level SUCCESS
    Write-Log "A system restart is required for the page file change to take effect." -Level WARNING
}

# --- Step 4: Install IIS and Windows Features ----------------------------------
function Install-WindowsFeatures {
    Write-Section "STEP 4: Install Required IIS and Windows Features"

    $features = @(
        # Core Web Server
        "Web-Server",
        "Web-WebServer",

        # Common HTTP Features
        "Web-Common-Http",
        "Web-Default-Doc",
        "Web-Dir-Browsing",
        "Web-Http-Errors",
        "Web-Static-Content",
        "Web-Http-Redirect",

        # Health & Diagnostics
        "Web-Health",
        "Web-Http-Logging",
        "Web-Log-Libraries",
        "Web-Request-Monitor",
        "Web-Http-Tracing",

        # Performance
        "Web-Performance",
        "Web-Stat-Compression",
        "Web-Dyn-Compression",

        # Security
        "Web-Security",
        "Web-Filtering",
        "Web-Basic-Auth",
        "Web-Windows-Auth",

        # Application Development
        "Web-App-Dev",
        "Web-Net-Ext",
        "Web-Net-Ext45",
        "Web-Asp-Net",
        "Web-Asp-Net45",
        "Web-ISAPI-Ext",
        "Web-ISAPI-Filter",

        # Management Tools
        "Web-Mgmt-Tools",
        "Web-Mgmt-Console",
        "Web-Mgmt-Compat",
        "Web-Metabase",
        "Web-Lgcy-Scripting",
        "Web-WMI",
        "Web-Scripting-Tools",

        # .NET Framework
        "NET-Framework-Features",
        "NET-Framework-Core",
        "NET-Framework-45-Features",
        "NET-Framework-45-Core",
        "NET-Framework-45-ASPNET",
        "NET-WCF-Services45",
        "NET-WCF-HTTP-Activation45",
        "NET-WCF-TCP-Activation45",

        # Windows Process Activation Service
        "WAS",
        "WAS-Process-Model",
        "WAS-NET-Environment",
        "WAS-Config-APIs",

        # Other required features
        "Windows-Identity-Foundation",
        "Server-Media-Foundation",
        "RSAT-DNS-Server",
        "RSAT-AD-PowerShell"
    )

    Write-Log "Installing $($features.Count) Windows features..." -Level INFO

    try {
        $result = Install-WindowsFeature -Name $features -IncludeManagementTools -ErrorAction Stop

        if ($result.Success) {
            Write-Log "Windows features installed successfully." -Level SUCCESS
            if ($result.RestartNeeded -eq "Yes") {
                Write-Log "A system restart is needed after feature installation." -Level WARNING
            }
        } else {
            Write-Log "Some Windows features may not have installed correctly." -Level WARNING
        }
    } catch {
        Write-Log "Failed to install Windows features -- $_" -Level ERROR
        throw
    }
}

# --- Step 5: Download and Extract SPSE Binaries --------------------------------
function Get-SPSEBinaries {
    Write-Section "STEP 5: Download and Extract SPSE Binaries"

    $baseUrl       = "http://abd.com/sp/se"
    $updateExeName = "uber-subscription-$CU-fullfile-x64-glb.exe"

    $downloads = @(
        @{
            Url         = "$baseUrl/se.zip"
            Destination = "D:\Binaries\SP\SE\SharePoint\se.zip"
            ExtractTo   = "D:\Binaries\SP\SE\SharePoint"
            IsZip       = $true
            Label       = "SPSE SharePoint Binaries"
        },
        @{
            Url         = "$baseUrl/jp.zip"
            Destination = "D:\Binaries\SP\SE\LanguagePacks\jp.zip"
            ExtractTo   = "D:\Binaries\SP\SE\LanguagePacks"
            IsZip       = $true
            Label       = "SPSE Japanese Language Pack"
        },
        @{
            Url         = "$baseUrl/$updateExeName"
            Destination = "D:\Binaries\SP\SE\Updates\$updateExeName"
            ExtractTo   = $null
            IsZip       = $false
            Label       = "SPSE Cumulative Update ($CU)"
        },
        @{
            Url         = "$baseUrl/PrerequisitesInstaller.zip"
            Destination = "D:\Binaries\SP\PrerequisitesInstaller\PrerequisitesInstaller.zip"
            ExtractTo   = "D:\Binaries\SP\PrerequisitesInstaller"
            IsZip       = $true
            Label       = "SPSE Prerequisites Installer"
        }
    )

    foreach ($item in $downloads) {
        Write-Log "--- $($item.Label) ---" -Level INFO
        Invoke-FileDownload -Url $item.Url -Destination $item.Destination

        if ($item.IsZip) {
            Expand-ZipArchive -ZipPath $item.Destination -Destination $item.ExtractTo
            Remove-Item -Path $item.Destination -Force -ErrorAction SilentlyContinue
            Write-Log "Removed archive: $($item.Destination)" -Level INFO
        }
    }
}

# --- Step 6: Install SPSE Prerequisites ----------------------------------------
function Install-SPSEPrerequisites {
    Write-Section "STEP 6: Install SharePoint Prerequisites"

    $prereqDir = "D:\Binaries\SP\PrerequisitesInstaller"
    $prereqExe = Join-Path -Path $prereqDir -ChildPath "prerequisiteinstaller.exe"

    if (-not (Test-Path -Path $prereqExe)) {
        Write-Log "prerequisiteinstaller.exe not found at: $prereqExe" -Level ERROR
        throw "Prerequisites installer executable not found."
    }

    $arguments = @(
        "/unattended",
        "/SQLNCli:`"$prereqDir\sqlncli.msi`"",
        "/Sync:`"$prereqDir\Synchronization.msi`"",
        "/AppFabric:`"$prereqDir\WindowsServerAppFabricSetup_x64.exe`"",
        "/IDFX11:`"$prereqDir\MicrosoftIdentityExtensions-64.msi`"",
        "/MSIPCClient:`"$prereqDir\setup_msipc_x64.exe`"",
        "/WCFDataServices56:`"$prereqDir\WcfDataServices.exe`"",
        "/DotNetFx:`"$prereqDir\dotnet-sdk.exe`"",
        "/MSVCRT141:`"$prereqDir\vc_redist.x64.exe`"",
        "/ODBC:`"$prereqDir\msodbcsql.msi`""
    ) -join " "

    Invoke-ExternalProcess `
        -FilePath         $prereqExe `
        -ArgumentList     $arguments `
        -Description      "SharePoint Prerequisites Installer" `
        -SuccessExitCodes @(0, 1001, 3010)
}

# --- Step 7: Install SharePoint Server Application -----------------------------
function Install-SPSEApplication {
    Write-Section "STEP 7: Install SharePoint Server Subscription Edition"

    $setupExe  = "D:\Binaries\SP\SE\SharePoint\setup.exe"
    $configXml = "D:\Binaries\SP\SE\SharePoint\config.xml"

    if (-not (Test-Path -Path $setupExe)) {
        Write-Log "setup.exe not found at: $setupExe" -Level ERROR
        throw "SharePoint setup.exe not found."
    }

    if (-not (Test-Path -Path $configXml)) {
        Write-Log "config.xml not found at: $configXml" -Level ERROR
        throw "SharePoint config.xml not found."
    }

    $arguments = "/config `"$configXml`""

    Invoke-ExternalProcess `
        -FilePath         $setupExe `
        -ArgumentList     $arguments `
        -Description      "SharePoint Server SE Application Installation" `
        -SuccessExitCodes @(0, 3010)
}

# --- Step 8: Install Japanese Language Pack ------------------------------------
function Install-SPSELanguagePack {
    Write-Section "STEP 8: Install SPSE Japanese Language Pack"

    $lpDir   = "D:\Binaries\SP\SE\LanguagePacks"
    $lpSetup = Join-Path -Path $lpDir -ChildPath "setup.exe"

    if (-not (Test-Path -Path $lpSetup)) {
        $lpSetup = Get-ChildItem -Path $lpDir -Filter "*.exe" -Recurse |
                   Where-Object { $_.Name -match "setup|languagepack|lpksetup" -or $_.Name -like "*ja-jp*" } |
                   Select-Object -First 1 -ExpandProperty FullName
    }

    if (-not $lpSetup -or -not (Test-Path -Path $lpSetup)) {
        Write-Log "Language Pack installer not found in: $lpDir" -Level ERROR
        throw "Language Pack executable not found."
    }

    Write-Log "Using Language Pack installer: $lpSetup" -Level INFO

    $lpConfigXml = Join-Path -Path $lpDir -ChildPath "config.xml"
    $arguments   = if (Test-Path $lpConfigXml) { "/config `"$lpConfigXml`"" } else { "/unattend" }

    Invoke-ExternalProcess `
        -FilePath         $lpSetup `
        -ArgumentList     $arguments `
        -Description      "SPSE Japanese Language Pack Installation" `
        -SuccessExitCodes @(0, 3010)
}

# --- Step 9: Install Cumulative Update -----------------------------------------
function Install-CumulativeUpdate {
    Write-Section "STEP 9: Install Cumulative Update ($CU)"

    $updateExeName = "uber-subscription-$CU-fullfile-x64-glb.exe"
    $updateExePath = "D:\Binaries\SP\SE\Updates\$updateExeName"

    if (-not (Test-Path -Path $updateExePath)) {
        Write-Log "CU executable not found at: $updateExePath" -Level ERROR
        throw "Cumulative Update executable not found."
    }

    $arguments = "/passive /norestart /log:`"$LogDirectory\CU-$CU-install.log`""

    Invoke-ExternalProcess `
        -FilePath         $updateExePath `
        -ArgumentList     $arguments `
        -Description      "SPSE Cumulative Update $CU" `
        -SuccessExitCodes @(0, 3010)
}

#endregion

###############################################################################
#region --- Main Execution
###############################################################################

Initialize-Log

Write-Log "Script parameters:" -Level INFO
Write-Log "  SPSEBuild    : $SPSEBuild"    -Level INFO
Write-Log "  CU           : $CU"           -Level INFO
Write-Log "  LogDirectory : $LogDirectory" -Level INFO

$script:OverallSuccess = $true

$steps = @(
    @{ Name = "Initialize-FolderStructure";    Fn = { Initialize-FolderStructure } },
    @{ Name = "Set-TempEnvironment";           Fn = { Set-TempEnvironment } },
    @{ Name = "Set-PageFile";                  Fn = { Set-PageFile } },
    @{ Name = "Install-WindowsFeatures";       Fn = { Install-WindowsFeatures } },
    @{ Name = "Get-SPSEBinaries";              Fn = { Get-SPSEBinaries } },
    @{ Name = "Install-SPSEPrerequisites";     Fn = { Install-SPSEPrerequisites } },
    @{ Name = "Install-SPSEApplication";       Fn = { Install-SPSEApplication } },
    @{ Name = "Install-SPSELanguagePack";      Fn = { Install-SPSELanguagePack } },
    @{ Name = "Install-CumulativeUpdate";      Fn = { Install-CumulativeUpdate } }
)

foreach ($step in $steps) {
    try {
        & $step.Fn
    } catch {
        Write-Log "Step '$($step.Name)' FAILED: $_" -Level ERROR
        $script:OverallSuccess = $false
        $choice = Read-Host "Step '$($step.Name)' encountered an error. Continue with next step? (Y/N)"
        if ($choice -notmatch '^[Yy]') {
            Write-Log "User chose to abort the script." -Level WARNING
            break
        }
    }
}

Write-Section "SPSE Prerequisites Setup -- Summary"

if ($script:OverallSuccess) {
    Write-Log "All steps completed successfully." -Level SUCCESS
} else {
    Write-Log "One or more steps failed. Review the log for details: $LogFilePath" -Level WARNING
}

Write-Log "Log file saved to: $LogFilePath" -Level INFO
Write-Log "NOTE: A system restart may be required to complete the configuration." -Level WARNING

#endregion