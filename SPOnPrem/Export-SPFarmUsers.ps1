# =============================================================================
# Export-SPFarmUsers.ps1
# Description : Exports full user information from a SharePoint SE On-Premises
#               farm to a CSV file, including logon name, email, and domain.
# Requirements: SharePoint Server SE (Subscription Edition) Management Shell
#               Must be run as a Farm Administrator
# =============================================================================

# ── Configuration ─────────────────────────────────────────────────────────────
$OutputPath = "C:\Reports\SharePoint_Users_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"

# Create output directory if it doesn't exist
$OutputDir = Split-Path $OutputPath -Parent
if (-not (Test-Path $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
}

# ── Load SharePoint Snap-in (if not already loaded) ───────────────────────────
if ((Get-PSSnapin -Name "Microsoft.SharePoint.PowerShell" -ErrorAction SilentlyContinue) -eq $null) {
    Add-PSSnapin Microsoft.SharePoint.PowerShell
}

# ── Helper: Parse domain and logon from LoginName ─────────────────────────────
function Parse-LoginInfo {
    param ([string]$LoginName)

    $domain    = ""
    $logon     = $LoginName
    $authType  = "Unknown"

    if ($LoginName -match "^i:0#\.w\|(.+)\\(.+)$") {
        # Claims Windows: i:0#.w|DOMAIN\user
        $authType = "Claims-Windows"
        $domain   = $Matches[1]
        $logon    = $Matches[2]
    }
    elseif ($LoginName -match "^(.+)\\(.+)$") {
        # Classic Windows: DOMAIN\user
        $authType = "Classic-Windows"
        $domain   = $Matches[1]
        $logon    = $Matches[2]
    }
    elseif ($LoginName -match "^i:0#\.f\|membership\|(.+)$") {
        # Forms-based authentication
        $authType = "FBA"
        $logon    = $Matches[1]
        $domain   = "FBA"
    }
    elseif ($LoginName -match "^i:0e\.t\|saml\|(.+)$") {
        # SAML / Trusted Identity Provider
        $authType = "SAML"
        $logon    = $Matches[1]
        $domain   = "SAML"
    }

    return [PSCustomObject]@{
        ParsedLogon  = $logon
        ParsedDomain = $domain
        AuthType     = $authType
    }
}

# ── Main Collection ────────────────────────────────────────────────────────────
Write-Host "`n[$(Get-Date -Format 'HH:mm:ss')] Starting SharePoint user export..." -ForegroundColor Cyan

$allUsers = [System.Collections.Generic.List[PSCustomObject]]::new()
$seenKeys = [System.Collections.Generic.HashSet[string]]::new()

$webApps = Get-SPWebApplication
Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Found $($webApps.Count) Web Application(s)." -ForegroundColor Yellow

foreach ($webApp in $webApps) {
    Write-Host "`n  Web Application : $($webApp.Url)" -ForegroundColor Magenta

    foreach ($site in $webApp.Sites) {
        Write-Host "    Site Collection: $($site.Url)" -ForegroundColor DarkCyan

        foreach ($web in $site.AllWebs) {
            try {
                foreach ($user in $web.AllUsers) {

                    # Deduplicate by LoginName + SiteCollection
                    $key = "$($site.Url)|$($user.LoginName)"
                    if (-not $seenKeys.Add($key)) { continue }

                    $parsed = Parse-LoginInfo -LoginName $user.LoginName

                    $allUsers.Add([PSCustomObject]@{
                        # ── Identity ───────────────────────────────────────
                        LoginName        = $user.LoginName
                        UserLogon        = $parsed.ParsedLogon
                        Domain           = $parsed.ParsedDomain
                        AuthType         = $parsed.AuthType

                        # ── Profile ────────────────────────────────────────
                        DisplayName      = $user.Name
                        Email            = $user.Email
                        Notes            = $user.Notes

                        # ── Flags ──────────────────────────────────────────
                        IsSiteAdmin      = $user.IsSiteAdmin
                        IsDomainGroup    = $user.IsDomainGroup
                        IsHiddenInUI     = $user.IsHiddenInUI

                        # ── Groups membership (semicolon-separated) ────────
                        Groups           = ($user.Groups | Select-Object -ExpandProperty Name) -join "; "

                        # ── Location ───────────────────────────────────────
                        WebApplication   = $webApp.Url
                        SiteCollection   = $site.Url
                        Web              = $web.Url

                        # ── Timestamp ──────────────────────────────────────
                        ExportedAt       = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
                    })
                }
            }
            catch {
                Write-Warning "      Could not read users from $($web.Url): $_"
            }
            finally {
                $web.Dispose()
            }
        }

        $site.Dispose()
    }
}

# ── Export ─────────────────────────────────────────────────────────────────────
Write-Host "`n[$(Get-Date -Format 'HH:mm:ss')] Total unique user records collected: $($allUsers.Count)" -ForegroundColor Green

$allUsers |
    Sort-Object Domain, UserLogon |
    Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8

Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Report saved to: $OutputPath" -ForegroundColor Green

# ── Quick summary ──────────────────────────────────────────────────────────────
Write-Host "`n── Summary by Domain ──────────────────────────────────────────" -ForegroundColor Cyan
$allUsers |
    Group-Object Domain |
    Sort-Object Count -Descending |
    Format-Table Name, Count -AutoSize

Write-Host "── Summary by Auth Type ───────────────────────────────────────" -ForegroundColor Cyan
$allUsers |
    Group-Object AuthType |
    Sort-Object Count -Descending |
    Format-Table Name, Count -AutoSize

Write-Host "`nDone.`n" -ForegroundColor Green
