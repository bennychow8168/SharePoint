#Requires -Version 5.1
<#
.SYNOPSIS
    Generates per-site-collection user access reports for a SharePoint 2019 on-premises farm.

.DESCRIPTION
    Uses the SharePoint REST API to enumerate every site collection and all its sub-sites,
    then collects user/group permission assignments.

    One CSV is produced per site collection, saved inside -OutputFolder.
    File names are sanitised from the site collection URL.

    Report columns:
        SiteUrl | UserID | Email | DisplayName | PermissionLevel | SourceGroup | WebTitle

    System and built-in service accounts are always excluded.

.PARAMETER WebAppUrl
    Root URL of the SharePoint Web Application (e.g. https://sharepoint.contoso.com).

.PARAMETER OutputFolder
    Folder where per-site-collection CSV files will be written.
    Defaults to .\SP_AccessReport_<timestamp>\ (created automatically).

.PARAMETER Credential
    PSCredential for SharePoint authentication.
    If omitted, Windows integrated (default) credentials are used.

.PARAMETER IncludeGroupMembers
    Expand SharePoint group membership so each member gets its own row.

.EXAMPLE
    .\Get-SPUserAccessReport.ps1 -WebAppUrl "https://sharepoint.contoso.com"

.EXAMPLE
    $cred = Get-Credential
    .\Get-SPUserAccessReport.ps1 -WebAppUrl "https://sharepoint.contoso.com" `
        -Credential $cred -IncludeGroupMembers -OutputFolder "C:\Reports\SP"
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory)] [string]$WebAppUrl,
    [string]$OutputFolder = ".\SP_AccessReport_$(Get-Date -Format 'yyyyMMdd_HHmmss')",
    [PSCredential]$Credential,
    [switch]$IncludeGroupMembers
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Sub-sites are ALWAYS included; this script always drills down recursively.

#region -- System account filter ----------------------------------------------
$SystemAccountPatterns = @(
    'sharepoint\system',
    'nt authority\authenticated users',
    'nt authority\local service',
    'nt authority\network service',
    'nt service\',
    'sharepoint app',
    'everyone',
    'all users',
    'c:0(.s|.t)#.windows nte',
    'spsearch',
    'spfarm',
    'app@sharepoint',
    'spocsvc',
    'spocrawl'
)

function Test-IsSystemAccount {
    param([string]$LoginName, [string]$Email)
    $combined = "$($LoginName.ToLower()) $($Email.ToLower())"
    foreach ($p in $SystemAccountPatterns) {
        if ($combined -like "*$p*") { return $true }
    }
    return $false
}
#endregion

#region -- REST helpers -------------------------------------------------------
function New-RestHeaders {
    @{ 'Accept' = 'application/json;odata=verbose'; 'Content-Type' = 'application/json;odata=verbose' }
}

function Invoke-SpRestGet {
    param([string]$Url, [hashtable]$Headers, [PSCredential]$Cred)
    $splat = @{ Uri = $Url; Method = 'GET'; Headers = $Headers; UseBasicParsing = $true }
    if ($Cred) { $splat['Credential'] = $Cred } else { $splat['UseDefaultCredentials'] = $true }
    try {
        $resp = Invoke-WebRequest @splat
        return ($resp.Content | ConvertFrom-Json).d
    }
    catch {
        $code = $null
        if ($_.Exception.Response) { $code = $_.Exception.Response.StatusCode.value__ }
        Write-Warning "  [HTTP $code] $Url`n  $_"
        return $null
    }
}
#endregion

#region -- Site collection discovery -----------------------------------------
function Get-SiteCollections {
    param([string]$WebApp, [hashtable]$Headers, [PSCredential]$Cred)
    Write-Host "`nDiscovering site collections in: $WebApp" -ForegroundColor Cyan

    $url = "$WebApp/_api/search/query" +
           "?querytext='contentclass:STS_Site'" +
           "&selectproperties='SPSiteUrl,Title'" +
           "&rowlimit=500&trimduplicates=false"

    $result = Invoke-SpRestGet -Url $url -Headers $Headers -Cred $Cred
    if (-not $result) { return @() }

    $rows = $result.query.PrimaryQueryResult.RelevantResults.Table.Rows.results
    foreach ($row in $rows) {
        $cells = $row.Cells.results
        [PSCustomObject]@{
            Url   = ($cells | Where-Object Key -eq 'SPSiteUrl').Value
            Title = ($cells | Where-Object Key -eq 'Title').Value
        }
    }
}
#endregion

#region -- Sub-site enumeration (BFS, always on) -----------------------------
function Get-AllWebs {
    param([string]$SiteUrl, [hashtable]$Headers, [PSCredential]$Cred)

    # Start with the root web; include Title for the report
    $rootData = Invoke-SpRestGet -Url "$SiteUrl/_api/web?`$select=Url,Title" -Headers $Headers -Cred $Cred
    $webs = [System.Collections.Generic.List[PSCustomObject]]::new()
    if ($rootData) {
        $webs.Add([PSCustomObject]@{ Url = $rootData.Url; Title = $rootData.Title })
    } else {
        $webs.Add([PSCustomObject]@{ Url = $SiteUrl; Title = $SiteUrl })
    }

    $queue = [System.Collections.Generic.Queue[string]]::new()
    $queue.Enqueue($SiteUrl)

    while ($queue.Count -gt 0) {
        $current = $queue.Dequeue()
        $subData = Invoke-SpRestGet -Url "$current/_api/web/webs?`$select=Url,Title" -Headers $Headers -Cred $Cred
        if (-not $subData) { continue }
        foreach ($sub in $subData.results) {
            $webs.Add([PSCustomObject]@{ Url = $sub.Url; Title = $sub.Title })
            $queue.Enqueue($sub.Url)
        }
    }
    return $webs
}
#endregion

#region -- Role assignments for a single web ---------------------------------
function Get-WebRoleAssignments {
    param([string]$WebUrl, [hashtable]$Headers, [PSCredential]$Cred)
    $url = "$WebUrl/_api/web/roleassignments" +
           "?`$expand=Member,RoleDefinitionBindings" +
           "&`$select=Member/Id,Member/LoginName,Member/Email,Member/Title,Member/PrincipalType," +
                     "RoleDefinitionBindings/Name"
    return Invoke-SpRestGet -Url $url -Headers $Headers -Cred $Cred
}
#endregion

#region -- SharePoint group member expansion ---------------------------------
function Get-GroupMembers {
    param([string]$WebUrl, [int]$GroupId, [hashtable]$Headers, [PSCredential]$Cred)
    $data = Invoke-SpRestGet -Url "$WebUrl/_api/web/sitegroups/getbyid($GroupId)/users?`$select=LoginName,Email,Title" `
                             -Headers $Headers -Cred $Cred
    if (-not $data) { return @() }
    return $data.results
}
#endregion

#region -- Sanitise a URL into a safe filename --------------------------------
function ConvertTo-SafeFileName {
    param([string]$Url)
    # Strip scheme, replace non-alphanumeric chars with underscores, trim length
    $safe = $Url -replace '^https?://', '' -replace '[^a-zA-Z0-9\-]', '_'
    $safe = $safe.Trim('_')
    if ($safe.Length -gt 100) { $safe = $safe.Substring(0, 100) }
    return $safe
}
#endregion

#region -- Main ---------------------------------------------------------------
$headers = New-RestHeaders

# Ensure output folder exists
if (-not (Test-Path $OutputFolder)) {
    New-Item -ItemType Directory -Path $OutputFolder | Out-Null
    Write-Host "Created output folder: $OutputFolder" -ForegroundColor DarkGray
}

$sites = Get-SiteCollections -WebApp $WebAppUrl -Headers $headers -Cred $Credential

if ($sites.Count -eq 0) {
    Write-Warning "No site collections found. Verify the Web Application URL and that the account has Full Read access."
    exit 1
}

Write-Host "Found $($sites.Count) site collection(s).`n" -ForegroundColor Green

$PT_USER  = 1   # Individual user
$PT_ADGRP = 4   # AD Security Group
$PT_SPGRP = 8   # SharePoint Group

$grandTotal = 0

foreach ($site in $sites) {
    Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkGray
    Write-Host "Site Collection : $($site.Url)" -ForegroundColor Yellow

    $webs = Get-AllWebs -SiteUrl $site.Url -Headers $headers -Cred $Credential
    Write-Host "Webs found      : $($webs.Count) (root + sub-sites)" -ForegroundColor DarkCyan

    $siteRows = [System.Collections.Generic.List[PSCustomObject]]::new()

    foreach ($web in $webs) {
        Write-Verbose "  Processing web: $($web.Url)"

        $raData = Get-WebRoleAssignments -WebUrl $web.Url -Headers $headers -Cred $Credential
        if (-not $raData) { continue }

        foreach ($ra in $raData.results) {
            $member        = $ra.Member
            $roles         = ($ra.RoleDefinitionBindings.results |
                              Select-Object -ExpandProperty Name) -join '; '
            $principalType = [int]$member.PrincipalType

            switch ($principalType) {

                $PT_USER {
                    if (Test-IsSystemAccount $member.LoginName $member.Email) { break }
                    $siteRows.Add([PSCustomObject]@{
                        SiteCollectionUrl = $site.Url
                        WebUrl            = $web.Url
                        WebTitle          = $web.Title
                        UserID            = $member.LoginName
                        Email             = $member.Email
                        DisplayName       = $member.Title
                        PermissionLevel   = $roles
                        SourceGroup       = ''
                    })
                    break
                }

                $PT_SPGRP {
                    if ($IncludeGroupMembers) {
                        # Resolve numeric group ID
                        $gid = 0
                        if ($member.Id) { $gid = [int]$member.Id }
                        elseif ($member.__metadata.id -match "getbyid\((\d+)\)") { $gid = [int]$Matches[1] }
                        if ($gid -eq 0) { break }

                        $members = Get-GroupMembers -WebUrl $web.Url -GroupId $gid -Headers $headers -Cred $Credential
                        foreach ($u in $members) {
                            if (Test-IsSystemAccount $u.LoginName $u.Email) { continue }
                            $siteRows.Add([PSCustomObject]@{
                                SiteCollectionUrl = $site.Url
                                WebUrl            = $web.Url
                                WebTitle          = $web.Title
                                UserID            = $u.LoginName
                                Email             = $u.Email
                                DisplayName       = $u.Title
                                PermissionLevel   = $roles
                                SourceGroup       = $member.Title
                            })
                        }
                    } else {
                        # Record the group as a single entry
                        $siteRows.Add([PSCustomObject]@{
                            SiteCollectionUrl = $site.Url
                            WebUrl            = $web.Url
                            WebTitle          = $web.Title
                            UserID            = $member.LoginName
                            Email             = ''
                            DisplayName       = "$($member.Title) [SP Group]"
                            PermissionLevel   = $roles
                            SourceGroup       = ''
                        })
                    }
                    break
                }

                $PT_ADGRP {
                    if (Test-IsSystemAccount $member.LoginName $member.Email) { break }
                    $siteRows.Add([PSCustomObject]@{
                        SiteCollectionUrl = $site.Url
                        WebUrl            = $web.Url
                        WebTitle          = $web.Title
                        UserID            = $member.LoginName
                        Email             = $member.Email
                        DisplayName       = "$($member.Title) [AD Security Group]"
                        PermissionLevel   = $roles
                        SourceGroup       = ''
                    })
                    break
                }
            }
        }
    }

    # -- Write per-site-collection CSV ---------------------------------------
    $safeName = ConvertTo-SafeFileName -Url $site.Url
    $csvPath  = Join-Path $OutputFolder "$safeName.csv"

    if ($siteRows.Count -gt 0) {
        $siteRows |
            Sort-Object WebUrl, UserID |
            Select-Object SiteCollectionUrl, WebUrl, WebTitle,
                          UserID, Email, DisplayName,
                          PermissionLevel, SourceGroup |
            Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8

        Write-Host "Records written : $($siteRows.Count)  ->  $csvPath" -ForegroundColor Green
    } else {
        Write-Host "Records written : 0 (no non-system users found)" -ForegroundColor DarkYellow
    }

    $grandTotal += $siteRows.Count
}

Write-Host "`n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkGray
Write-Host "All done. Total records across all sites: $grandTotal" -ForegroundColor Cyan
Write-Host "Output folder: $(Resolve-Path $OutputFolder)" -ForegroundColor Cyan
#endregion
