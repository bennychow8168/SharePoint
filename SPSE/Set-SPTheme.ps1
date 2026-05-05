# =============================================================================
# Set-SPTheme.ps1
# Applies a custom modern theme to SharePoint Server Subscription Edition
# using the thememanager REST API with Windows Authentication.
#
# Usage:
#   .\Set-SPTheme.ps1 -SiteUrl "http://your-sharepoint-site" -ThemeName "Corporate Theme"
#
# Requirements:
#   - Run as a Site Collection Administrator on the target site
#   - Windows Authentication (NTLM/Kerberos) must be enabled on the web application
# =============================================================================

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [string]$SiteUrl,

    [Parameter(Mandatory = $false)]
    [string]$ThemeName = "Corporate Theme",

    # Optional: pass explicit credentials, otherwise uses current Windows session
    [Parameter(Mandatory = $false)]
    [System.Management.Automation.PSCredential]$Credential
)

# =============================================================================
# THEME PALETTE — Edit these hex values to match your brand colors.
# Tip: Generate the full palette at https://aka.ms/themedesigner
#      then paste the JSON values here.
# =============================================================================
$themePalette = [ordered]@{
    themePrimary         = "#1E479A"
    themeLighterAlt      = "#f3f6fb"
    themeLighter         = "#d0daef"
    themeLight           = "#aabce0"
    themeTertiary        = "#6483c2"
    themeSecondary       = "#2f57a5"
    themeDarkAlt         = "#1a3f8a"
    themeDark            = "#163574"
    themeDarker          = "#102756"
    neutralLighterAlt    = "#f8f8f8"
    neutralLighter       = "#f4f4f4"
    neutralLight         = "#eaeaea"
    neutralQuaternaryAlt = "#dadada"
    neutralQuaternary    = "#d0d0d0"
    neutralTertiaryAlt   = "#c8c8c8"
    neutralTertiary      = "#bab8b7"
    neutralSecondary     = "#a3a2a0"
    neutralPrimaryAlt    = "#8d8b8a"
    neutralPrimary       = "#323130"
    neutralDark          = "#605e5d"
    black                = "#494847"
    white                = "#ffffff"
}

# =============================================================================
# HELPER — Build a WebSession with Windows Auth or explicit credentials
# =============================================================================
function New-SPSession {
    param (
        [string]$Url,
        [System.Management.Automation.PSCredential]$Credential
    )

    $session = New-Object Microsoft.PowerShell.Commands.WebRequestSession

    if ($Credential) {
        $session.Credentials = $Credential.GetNetworkCredential()
        Write-Verbose "Using supplied credentials: $($Credential.UserName)"
    }
    else {
        $session.UseDefaultCredentials = $true
        Write-Verbose "Using current Windows session credentials"
    }

    return $session
}

# =============================================================================
# STEP 1 — Fetch Form Digest (equivalent to _spPageContextInfo.formDigestValue)
# =============================================================================
function Get-SPFormDigest {
    param (
        [string]$SiteUrl,
        [Microsoft.PowerShell.Commands.WebRequestSession]$Session
    )

    $contextUrl = "$SiteUrl/_api/contextinfo"

    try {
        Write-Host "  Fetching form digest from: $contextUrl" -ForegroundColor Gray

        $response = Invoke-WebRequest `
            -Uri         $contextUrl `
            -Method      POST `
            -WebSession  $Session `
            -Headers     @{ "Accept" = "application/json;odata=verbose" } `
            -ContentType "application/json;charset=utf-8" `
            -UseBasicParsing

        $json = $response.Content | ConvertFrom-Json
        $digest = $json.d.GetContextWebInformation.FormDigestValue

        if (-not $digest) {
            throw "FormDigestValue was empty. Check that the URL is correct and you have access."
        }

        Write-Host "  Form digest obtained successfully." -ForegroundColor Gray
        return $digest
    }
    catch {
        throw "Failed to get form digest: $_"
    }
}

# =============================================================================
# STEP 2 — Apply the theme via /_api/thememanager/ApplyTheme
# =============================================================================
function Invoke-ApplyTheme {
    param (
        [string]$SiteUrl,
        [string]$ThemeName,
        [hashtable]$Palette,
        [string]$FormDigest,
        [Microsoft.PowerShell.Commands.WebRequestSession]$Session
    )

    $applyUrl = "$SiteUrl/_api/thememanager/ApplyTheme"

    # Build the nested JSON structure the API expects:
    # { name: "...", themeJson: "<escaped JSON string of palette>" }
    $paletteWrapper = @{ palette = $Palette }
    $paletteJson    = $paletteWrapper | ConvertTo-Json -Depth 5 -Compress
    $body           = @{ name = $ThemeName; themeJson = $paletteJson } | ConvertTo-Json -Depth 3

    try {
        Write-Host "  Posting theme to: $applyUrl" -ForegroundColor Gray

        $response = Invoke-WebRequest `
            -Uri         $applyUrl `
            -Method      POST `
            -WebSession  $Session `
            -Headers     @{
                "Accept"       = "application/json; odata.metadata=minimal"
                "X-RequestDigest" = $FormDigest
                "ODATA-VERSION"   = "4.0"
            } `
            -ContentType "application/json;charset=utf-8" `
            -Body        $body `
            -UseBasicParsing

        return $response
    }
    catch {
        throw "Failed to apply theme: $_"
    }
}

# =============================================================================
# MAIN
# =============================================================================

Write-Host ""
Write-Host "=======================================" -ForegroundColor Cyan
Write-Host "  SharePoint On-Premises Theme Applier " -ForegroundColor Cyan
Write-Host "=======================================" -ForegroundColor Cyan
Write-Host "  Site    : $SiteUrl"
Write-Host "  Theme   : $ThemeName"
Write-Host ""

# Normalize URL (strip trailing slash)
$SiteUrl = $SiteUrl.TrimEnd("/")

try {
    # Build session
    Write-Host "[1/3] Building authentication session..." -ForegroundColor Yellow
    $session = New-SPSession -Url $SiteUrl -Credential $Credential

    # Get form digest
    Write-Host "[2/3] Obtaining form digest..." -ForegroundColor Yellow
    $digest = Get-SPFormDigest -SiteUrl $SiteUrl -Session $session

    # Apply theme
    Write-Host "[3/3] Applying theme '$ThemeName'..." -ForegroundColor Yellow
    $result = Invoke-ApplyTheme `
        -SiteUrl    $SiteUrl `
        -ThemeName  $ThemeName `
        -Palette    $themePalette `
        -FormDigest $digest `
        -Session    $session

    if ($result.StatusCode -eq 200) {
        Write-Host ""
        Write-Host "✔  Theme '$ThemeName' applied successfully!" -ForegroundColor Green
        Write-Host "   Verify at: $SiteUrl/_layouts/15/designgallery.aspx" -ForegroundColor Gray
        Write-Host "   Or go to: Site Settings → Change the look" -ForegroundColor Gray
    }
    else {
        Write-Warning "Unexpected response code: $($result.StatusCode)"
        Write-Host $result.Content
    }
}
catch {
    Write-Host ""
    Write-Host "✘  Error: $_" -ForegroundColor Red
    Write-Host ""
    Write-Host "Troubleshooting tips:" -ForegroundColor Yellow
    Write-Host "  1. Ensure you are running this as a Site Collection Administrator."
    Write-Host "  2. Confirm Windows Authentication is enabled on the web application."
    Write-Host "  3. Try adding -Credential (Get-Credential) if the session auth fails."
    Write-Host "  4. Check that the SiteUrl is reachable from this machine."
    exit 1
}
