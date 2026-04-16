#Requires -Version 5.1
<#
.SYNOPSIS
    Exports all SharePoint 2019 on-premises farm users (login + email) to a CSV file.

.DESCRIPTION
    This script enumerates every Web Application, Site Collection, and Sub-Web
    in the SharePoint 2019 farm. It collects unique users via two strategies:

        1. PRIMARY  – SharePoint REST API  /_api/web/siteusers
                      (works without the server-side DLL; runs from any machine
                       with network access to SharePoint)

        2. FALLBACK – SharePoint Server Object Model
                      (requires running the script on a SharePoint server inside
                       the SharePoint Management Shell, or with the PSSnapin loaded)

    The REST API strategy is tried first for each site collection.
    If the call fails (e.g. permission error, network issue), the script falls
    back to the object-model automatically.

    Output: A CSV file with columns  LoginName, Email, DisplayName, Source

.PARAMETER OutputFile
    Full path to the output CSV file.
    Default: .\SP2019_FarmUsers_<timestamp>.csv

.PARAMETER UseObjectModelOnly
    Skip the REST API and use only the SharePoint Object Model.
    Requires running on a SharePoint server.

.PARAMETER Credential
    PSCredential used for REST API calls.
    If omitted, the script uses the current Windows identity (default auth).

.PARAMETER WebApplicationUrl
    Optionally scope the export to a single Web Application URL.
    If omitted, all Web Applications in the farm are processed.

.EXAMPLE
    # Run from a SharePoint server (uses Object Model + REST API)
    .\Export-SP2019FarmUsers.ps1

.EXAMPLE
    # Run from any machine targeting a specific web app, using explicit credentials
    .\Export-SP2019FarmUsers.ps1 -WebApplicationUrl "https://sharepoint.contoso.com" `
        -Credential (Get-Credential) -OutputFile "C:\Temp\users.csv"

.NOTES
    • Requires Farm Administrator or at least Site Collection Administrator rights.
    • Must be executed as a user who can enumerate all site collections.
    • Duplicate users (same LoginName) are deduplicated across the whole farm.
#>

[CmdletBinding()]
param(
    [string]$OutputFile,

    [switch]$UseObjectModelOnly,

    [System.Management.Automation.PSCredential]
    [System.Management.Automation.Credential()]
    $Credential,

    [string]$WebApplicationUrl
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ── Helpers ─────────────────────────────────────────────────────────────────

function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $color = switch ($Level) {
        'WARN'  { 'Yellow' }
        'ERROR' { 'Red'    }
        default { 'Cyan'   }
    }
    Write-Host "[$ts][$Level] $Message" -ForegroundColor $color
}

# ── Output path ──────────────────────────────────────────────────────────────

if (-not $OutputFile) {
    $ts         = Get-Date -Format 'yyyyMMdd_HHmmss'
    $OutputFile = Join-Path $PWD "SP2019_FarmUsers_$ts.csv"
}

$OutputDir = Split-Path $OutputFile -Parent
if ($OutputDir -and -not (Test-Path $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir | Out-Null
}

# ── Shared collection (LoginName → user object) ──────────────────────────────

$FarmUsers = [System.Collections.Generic.Dictionary[string,PSCustomObject]]::new(
    [System.StringComparer]::OrdinalIgnoreCase
)

function Add-User {
    param([string]$LoginName, [string]$Email, [string]$DisplayName, [string]$Source)
    $key = $LoginName.Trim()
    if ($key -and -not $FarmUsers.ContainsKey($key)) {
        $FarmUsers[$key] = [PSCustomObject]@{
            LoginName   = $LoginName.Trim()
            Email       = $Email.Trim()
            DisplayName = $DisplayName.Trim()
            Source      = $Source
        }
    }
}

# ════════════════════════════════════════════════════════════════════════════
#  STRATEGY 1 – REST API
# ════════════════════════════════════════════════════════════════════════════

function Get-SiteUsersViaREST {
    <#
    .SYNOPSIS  Calls /_api/web/siteusers and /_api/web/allproperties for a site.
    #>
    param([string]$SiteUrl)

    $headers = @{ 'Accept' = 'application/json;odata=verbose' }

    $iwrParams = @{
        Uri             = "$($SiteUrl.TrimEnd('/'))/_api/web/siteusers"
        Headers         = $headers
        UseBasicParsing = $true
        ErrorAction     = 'Stop'
    }

    if ($Credential) {
        $iwrParams['Credential'] = $Credential
    } else {
        $iwrParams['UseDefaultCredentials'] = $true
    }

    $response = Invoke-WebRequest @iwrParams
    $json     = $response.Content | ConvertFrom-Json

    foreach ($user in $json.d.results) {
        # Skip system accounts, groups, and empty logins
        if ($user.LoginName -match 'SHAREPOINT\\system|NT AUTHORITY|c:0\(.s\|true\)' ) { continue }
        if ($user.PrincipalType -ne 1) { continue }   # 1 = User, 4 = SecurityGroup, 8 = SharePointGroup

        Add-User -LoginName   $user.LoginName `
                 -Email       ($user.Email        ?? '') `
                 -DisplayName ($user.Title        ?? '') `
                 -Source      "REST:$SiteUrl"
    }
}

# ════════════════════════════════════════════════════════════════════════════
#  STRATEGY 2 – SharePoint Server Object Model
# ════════════════════════════════════════════════════════════════════════════

$ObjectModelAvailable = $false

function Initialize-ObjectModel {
    if (-not (Get-PSSnapin -Name 'Microsoft.SharePoint.PowerShell' -ErrorAction SilentlyContinue)) {
        try {
            Add-PSSnapin 'Microsoft.SharePoint.PowerShell' -ErrorAction Stop
            Write-Log 'SharePoint PowerShell snap-in loaded.'
            $script:ObjectModelAvailable = $true
        } catch {
            Write-Log 'SharePoint PSSnapin not available on this machine.' 'WARN'
            $script:ObjectModelAvailable = $false
        }
    } else {
        $script:ObjectModelAvailable = $true
    }
}

function Get-SiteUsersViaObjectModel {
    param([Microsoft.SharePoint.SPSite]$Site)

    foreach ($web in $Site.AllWebs) {
        try {
            foreach ($user in $web.AllUsers) {
                if ($user.LoginName -match 'SHAREPOINT\\system|NT AUTHORITY|c:0\(.s\|true\)') { continue }
                if ($user.IsDomainGroup) { continue }

                Add-User -LoginName   $user.LoginName `
                         -Email       ($user.Email        ?? '') `
                         -DisplayName ($user.DisplayName  ?? '') `
                         -Source      "ObjectModel:$($web.Url)"
            }
        } catch {
            Write-Log "  Could not enumerate users in $($web.Url): $_" 'WARN'
        } finally {
            $web.Dispose()
        }
    }
}

# ════════════════════════════════════════════════════════════════════════════
#  MAIN – Enumerate the Farm
# ════════════════════════════════════════════════════════════════════════════

Write-Log '══════════════════════════════════════════════'
Write-Log ' SharePoint 2019 Farm User Export – Starting'
Write-Log '══════════════════════════════════════════════'

Initialize-ObjectModel

# ── Decide how to get the list of site collections ───────────────────────────

$SiteUrls = [System.Collections.Generic.List[string]]::new()

if ($ObjectModelAvailable) {
    Write-Log 'Using Object Model to enumerate site collections...'

    $webApps = if ($WebApplicationUrl) {
        @(Get-SPWebApplication $WebApplicationUrl)
    } else {
        Get-SPWebApplication
    }

    foreach ($wa in $webApps) {
        Write-Log "  Web Application: $($wa.Url)"
        foreach ($sc in $wa.Sites) {
            $SiteUrls.Add($sc.Url)
            $sc.Dispose()
        }
    }
} else {
    # No Object Model – user must supply at least one URL
    if (-not $WebApplicationUrl) {
        Write-Log ('Object Model unavailable and -WebApplicationUrl not specified. ' +
                   'Please supply -WebApplicationUrl to enumerate sites via REST.') 'ERROR'
        exit 1
    }

    Write-Log 'No Object Model – enumerating site collections via REST _api/search...'

    # Use Search REST API to discover all site collections under the web app
    $searchUrl = "$($WebApplicationUrl.TrimEnd('/'))/_api/search/query?" +
                 "querytext='contentclass:STS_Site'&selectproperties='Path'&rowlimit=500&startrow=0"

    $iwrParams = @{
        Uri             = $searchUrl
        Headers         = @{ 'Accept' = 'application/json;odata=verbose' }
        UseBasicParsing = $true
        ErrorAction     = 'Stop'
    }
    if ($Credential)      { $iwrParams['Credential']          = $Credential }
    else                  { $iwrParams['UseDefaultCredentials'] = $true      }

    try {
        $resp  = Invoke-WebRequest @iwrParams
        $json  = $resp.Content | ConvertFrom-Json
        $rows  = $json.d.query.PrimaryQueryResult.RelevantResults.Table.Rows.results
        foreach ($row in $rows) {
            $path = ($row.Cells.results | Where-Object { $_.Key -eq 'Path' }).Value
            if ($path) { $SiteUrls.Add($path) }
        }
        Write-Log "  Discovered $($SiteUrls.Count) site collection(s) via Search."
    } catch {
        Write-Log "  Search REST enumeration failed: $_" 'WARN'
        Write-Log "  Falling back to root site only: $WebApplicationUrl" 'WARN'
        $SiteUrls.Add($WebApplicationUrl)
    }
}

Write-Log "Total site collections to process: $($SiteUrls.Count)"

# ── Process each site collection ─────────────────────────────────────────────

$processed = 0
$errors    = 0

foreach ($url in $SiteUrls) {
    $processed++
    Write-Log "[$processed/$($SiteUrls.Count)] Processing: $url"

    $restOk = $false

    if (-not $UseObjectModelOnly) {
        # Try REST API first
        try {
            Get-SiteUsersViaREST -SiteUrl $url
            $restOk = $true
            Write-Log "  REST API OK – unique users so far: $($FarmUsers.Count)"
        } catch {
            Write-Log "  REST API failed for $url : $($_.Exception.Message)" 'WARN'
        }
    }

    # Fall back to Object Model if REST failed or was skipped
    if (-not $restOk -and $ObjectModelAvailable) {
        try {
            $site = Get-SPSite $url -ErrorAction Stop
            Get-SiteUsersViaObjectModel -Site $site
            $site.Dispose()
            Write-Log "  Object Model OK – unique users so far: $($FarmUsers.Count)"
        } catch {
            Write-Log "  Object Model also failed for $url : $($_.Exception.Message)" 'ERROR'
            $errors++
        }
    }
}

# ════════════════════════════════════════════════════════════════════════════
#  Export to CSV
# ════════════════════════════════════════════════════════════════════════════

Write-Log '──────────────────────────────────────────────'
Write-Log "Total unique users collected: $($FarmUsers.Count)"
Write-Log "Sites with errors           : $errors"
Write-Log "Exporting to               : $OutputFile"

$FarmUsers.Values |
    Sort-Object LoginName |
    Select-Object LoginName, Email, DisplayName, Source |
    Export-Csv -Path $OutputFile -NoTypeInformation -Encoding UTF8

Write-Log '══════════════════════════════════════════════'
Write-Log " Export complete → $OutputFile"
Write-Log '══════════════════════════════════════════════'