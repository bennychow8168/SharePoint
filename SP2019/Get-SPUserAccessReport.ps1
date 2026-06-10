#Requires -Version 5.1
<#
.SYNOPSIS
    Generates per-site-collection user access reports for a SharePoint 2019 on-premises farm.

.DESCRIPTION
    Uses the SharePoint REST API to enumerate every site collection and all its sub-sites,
    then collects user/group permission assignments.

    One CSV is produced per site collection, saved inside -OutputFolder.
    Sub-sites are always included (full recursive traversal).

    Report columns:
        SiteCollectionUrl | WebUrl | WebTitle | UserID | Email | DisplayName | PermissionLevel | SourceGroup

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
    [Parameter(Mandatory)]
    [string]$WebAppUrl,

    [string]$OutputFolder = ".\SP_AccessReport_$(Get-Date -Format 'yyyyMMdd_HHmmss')",

    [PSCredential]$Credential,

    [switch]$IncludeGroupMembers
)

# NOTE: Do NOT use Set-StrictMode here — it causes ".Count not found" errors
# on single-item JSON-deserialized objects that are PSCustomObject, not arrays.
$ErrorActionPreference = 'Continue'

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

#region -- REST helper --------------------------------------------------------
function Invoke-SpRestGet {
    param(
        [string]$Url,
        [PSCredential]$Cred
    )
    $headers = @{
        'Accept'       = 'application/json;odata=verbose'
        'Content-Type' = 'application/json;odata=verbose'
    }
    $splat = @{
        Uri             = $Url
        Method          = 'GET'
        Headers         = $headers
        UseBasicParsing = $true
    }
    if ($Cred) {
        $splat['Credential'] = $Cred
    } else {
        $splat['UseDefaultCredentials'] = $true
    }
    try {
        $resp = Invoke-WebRequest @splat
        $parsed = $resp.Content | ConvertFrom-Json
        return $parsed.d
    }
    catch {
        $code = $null
        try { $code = $_.Exception.Response.StatusCode.value__ } catch {}
        Write-Warning "[HTTP $code] Failed: $Url`n  $($_.Exception.Message)"
        return $null
    }
}
#endregion

#region -- Safe array helper --------------------------------------------------
# Always returns a real PowerShell array, never $null, never a bare object
function ConvertTo-Array {
    param($InputObject)
    if ($null -eq $InputObject) { return @() }
    # Already an array or list
    if ($InputObject -is [System.Array] -or
        $InputObject -is [System.Collections.IList]) {
        return @($InputObject)
    }
    # Single object — wrap it
    return @($InputObject)
}
#endregion

#region -- Site collection discovery -----------------------------------------
function Get-SiteCollections {
    param([string]$WebApp, [PSCredential]$Cred)

    Write-Host "`nDiscovering site collections in: $WebApp" -ForegroundColor Cyan

    $url = "$WebApp/_api/search/query" +
           "?querytext='contentclass:STS_Site'" +
           "&selectproperties='SPSiteUrl,Title'" +
           "&rowlimit=500&trimduplicates=false"

    $d = Invoke-SpRestGet -Url $url -Cred $Cred
    if ($null -eq $d) { return @() }

    $rows = ConvertTo-Array $d.query.PrimaryQueryResult.RelevantResults.Table.Rows.results
    if ($rows.Count -eq 0) { return @() }

    $output = New-Object System.Collections.ArrayList
    foreach ($row in $rows) {
        $cells   = ConvertTo-Array $row.Cells.results
        $urlVal  = ($cells | Where-Object { $_.Key -eq 'SPSiteUrl' } | Select-Object -First 1).Value
        $titleVal= ($cells | Where-Object { $_.Key -eq 'Title'     } | Select-Object -First 1).Value
        if ($urlVal) {
            [void]$output.Add([PSCustomObject]@{ Url = $urlVal; Title = $titleVal })
        }
    }
    return @($output)
}
#endregion

#region -- Sub-site enumeration (BFS) ----------------------------------------
function Get-AllWebs {
    param([string]$SiteUrl, [PSCredential]$Cred)

    $output = New-Object System.Collections.ArrayList
    $queue  = New-Object System.Collections.Generic.Queue[string]

    # Root web
    $rootD = Invoke-SpRestGet -Url "$SiteUrl/_api/web?`$select=Url,Title" -Cred $Cred
    if ($rootD) {
        [void]$output.Add([PSCustomObject]@{ Url = [string]$rootD.Url; Title = [string]$rootD.Title })
        $queue.Enqueue([string]$rootD.Url)
    } else {
        [void]$output.Add([PSCustomObject]@{ Url = $SiteUrl; Title = $SiteUrl })
        $queue.Enqueue($SiteUrl)
    }

    while ($queue.Count -gt 0) {
        $current = $queue.Dequeue()
        $subD    = Invoke-SpRestGet -Url "$current/_api/web/webs?`$select=Url,Title" -Cred $Cred
        if ($null -eq $subD) { continue }

        $subs = ConvertTo-Array $subD.results
        foreach ($sub in $subs) {
            [void]$output.Add([PSCustomObject]@{ Url = [string]$sub.Url; Title = [string]$sub.Title })
            $queue.Enqueue([string]$sub.Url)
        }
    }

    return @($output)
}
#endregion

#region -- Role assignments ---------------------------------------------------
function Get-WebRoleAssignments {
    param([string]$WebUrl, [PSCredential]$Cred)

    $url = "$WebUrl/_api/web/roleassignments" +
           "?`$expand=Member,RoleDefinitionBindings" +
           "&`$select=Member/Id,Member/LoginName,Member/Email,Member/Title,Member/PrincipalType," +
                     "RoleDefinitionBindings/Name"

    $d = Invoke-SpRestGet -Url $url -Cred $Cred
    if ($null -eq $d) { return @() }
    return ConvertTo-Array $d.results
}
#endregion

#region -- SharePoint group member expansion ---------------------------------
function Get-GroupMembers {
    param([string]$WebUrl, [int]$GroupId, [PSCredential]$Cred)

    $d = Invoke-SpRestGet `
        -Url "$WebUrl/_api/web/sitegroups/getbyid($GroupId)/users?`$select=LoginName,Email,Title" `
        -Cred $Cred

    if ($null -eq $d) { return @() }
    return ConvertTo-Array $d.results
}
#endregion

#region -- Safe file name helper ---------------------------------------------
function ConvertTo-SafeFileName {
    param([string]$Url)
    $safe = $Url -replace '^https?://', '' -replace '[^a-zA-Z0-9\-]', '_'
    $safe = $safe.Trim('_')
    if ($safe.Length -gt 100) { $safe = $safe.Substring(0, 100) }
    return $safe
}
#endregion

#region -- Main ---------------------------------------------------------------

# Ensure output folder exists
if (-not (Test-Path $OutputFolder)) {
    New-Item -ItemType Directory -Path $OutputFolder | Out-Null
    Write-Host "Created output folder: $OutputFolder" -ForegroundColor DarkGray
}

$sites = Get-SiteCollections -WebApp $WebAppUrl -Cred $Credential

if ($null -eq $sites -or $sites.Count -eq 0) {
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

    $webs = Get-AllWebs -SiteUrl $site.Url -Cred $Credential
    Write-Host "Webs found      : $($webs.Count) (root + sub-sites)" -ForegroundColor DarkCyan

    $siteRows = New-Object System.Collections.ArrayList

    foreach ($web in $webs) {
        Write-Verbose "  Processing web: $($web.Url)"

        $assignments = Get-WebRoleAssignments -WebUrl $web.Url -Cred $Credential
        if ($null -eq $assignments -or $assignments.Count -eq 0) { continue }

        foreach ($ra in $assignments) {
            $member        = $ra.Member
            $principalType = [int]$member.PrincipalType

            $bindingResults = ConvertTo-Array $ra.RoleDefinitionBindings.results
            $roles = ($bindingResults | ForEach-Object { $_.Name } | Where-Object { $_ }) -join '; '

            switch ($principalType) {

                $PT_USER {
                    $login = [string]$member.LoginName
                    $email = [string]$member.Email
                    if (Test-IsSystemAccount -LoginName $login -Email $email) { break }
                    [void]$siteRows.Add([PSCustomObject]@{
                        SiteCollectionUrl = $site.Url
                        WebUrl            = $web.Url
                        WebTitle          = $web.Title
                        UserID            = $login
                        Email             = $email
                        DisplayName       = [string]$member.Title
                        PermissionLevel   = $roles
                        SourceGroup       = ''
                    })
                    break
                }

                $PT_SPGRP {
                    if ($IncludeGroupMembers) {
                        $gid = 0
                        if ($member.Id) { $gid = [int]$member.Id }
                        elseif ($member.__metadata.id -match "getbyid\((\d+)\)") { $gid = [int]$Matches[1] }
                        if ($gid -eq 0) { break }

                        $members = Get-GroupMembers -WebUrl $web.Url -GroupId $gid -Cred $Credential
                        foreach ($u in $members) {
                            $uLogin = [string]$u.LoginName
                            $uEmail = [string]$u.Email
                            if (Test-IsSystemAccount -LoginName $uLogin -Email $uEmail) { continue }
                            [void]$siteRows.Add([PSCustomObject]@{
                                SiteCollectionUrl = $site.Url
                                WebUrl            = $web.Url
                                WebTitle          = $web.Title
                                UserID            = $uLogin
                                Email             = $uEmail
                                DisplayName       = [string]$u.Title
                                PermissionLevel   = $roles
                                SourceGroup       = [string]$member.Title
                            })
                        }
                    } else {
                        [void]$siteRows.Add([PSCustomObject]@{
                            SiteCollectionUrl = $site.Url
                            WebUrl            = $web.Url
                            WebTitle          = $web.Title
                            UserID            = [string]$member.LoginName
                            Email             = ''
                            DisplayName       = "$($member.Title) [SP Group]"
                            PermissionLevel   = $roles
                            SourceGroup       = ''
                        })
                    }
                    break
                }

                $PT_ADGRP {
                    $login = [string]$member.LoginName
                    $email = [string]$member.Email
                    if (Test-IsSystemAccount -LoginName $login -Email $email) { break }
                    [void]$siteRows.Add([PSCustomObject]@{
                        SiteCollectionUrl = $site.Url
                        WebUrl            = $web.Url
                        WebTitle          = $web.Title
                        UserID            = $login
                        Email             = $email
                        DisplayName       = "$($member.Title) [AD Security Group]"
                        PermissionLevel   = $roles
                        SourceGroup       = ''
                    })
                    break
                }
            }
        }
    }

    # Write per-site-collection CSV
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
Write-Host "All done. Total records: $grandTotal" -ForegroundColor Cyan
Write-Host "Output folder: $(Resolve-Path $OutputFolder)" -ForegroundColor Cyan
#endregion
