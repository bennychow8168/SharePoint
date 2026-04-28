<#
.SYNOPSIS
    Exports user lists from a SharePoint 2019 Web Application using REST API.
    Creates one CSV file per site/subsite.
#>

param (
    [Parameter(Mandatory=$true)]
    [string]$WebApplicationUrl,

    [Parameter(Mandatory=$true)]
    [string]$OutputFolder,

    [switch]$IncludeSubsites
)

# 1. Load SharePoint Environment
if (!(Get-PSSnapin -Name Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue)) {
    Add-PSSnapin Microsoft.SharePoint.PowerShell
}

# 2. Ensure Output Directory
if (!(Test-Path $OutputFolder)) {
    New-Item -ItemType Directory -Path $OutputFolder | Out-Null
}

# 3. REST Helper Function (Handles Paging & Auth)
function Invoke-SPRest {
    param(
        [Parameter(Mandatory)][string]$Url
    )
    $headers = @{ "Accept" = "application/json;odata=verbose" }
    $allResults = [System.Collections.Generic.List[object]]::new()
    $nextUrl = $Url

    do {
        try {
            $response = Invoke-RestMethod -Uri $nextUrl -Headers $headers -Method GET -UseDefaultCredentials -ErrorAction Stop
            $data = if ($response.d.results) { $response.d.results } else { $response.d }
            
            if ($data -is [System.Collections.IEnumerable] -and $data -isnot [string]) {
                $allResults.AddRange([object[]]$data)
            } else { $allResults.Add($data) }

            $nextUrl = $response.d.__next
        }
        catch {
            Write-Warning "  [REST ERROR] $nextUrl : $_"
            return $null
        }
    } while ($nextUrl)
    return $allResults
}

# 4. User Resolution Function
function Get-SiteUserPermissions 
{
    param (
        [string]$WebUrl,
        [bool]$IsSiteRoot
    )

    $userMap = @{}

    # Helper to add/update user data
    $AddUser = {
        param($u, $role)
        if (-not $userMap.ContainsKey($u.LoginName)) {
            $userMap[$u.LoginName] = @{
                UserName = $u.Title; LoginID = $u.LoginName; Email = $u.Email; 
                Roles = [System.Collections.Generic.HashSet[string]]::new()
            }
        }
        if ($role) { [void]$userMap[$u.LoginName].Roles.Add($role) }
    }

    # A. Check Site Admins (Root Web only)
    if ($IsSiteRoot) {
        $admins = Invoke-SPRest -Url "$WebUrl/_api/web/siteusers?`$filter=IsSiteAdmin eq true"
        foreach ($u in $admins) { & $AddUser $u "Admin" }
    }

    # B. Check Standard Groups (Owner/Member/Visitor)
    $groups = @{ 
        "Owner" = "associatedownergroup"; 
        "Member" = "associatedmembergroup"; 
        "Visitor" = "associatedvisitorgroup" 
    }
    foreach ($key in $groups.Keys) {
        $users = Invoke-SPRest -Url "$WebUrl/_api/web/$($groups[$key])/users"
        foreach ($u in $users) { & $AddUser $u $key }
    }

    # C. Check Direct/Specific/Limited Access
    $raUrl = "$WebUrl/_api/web/roleassignments?`$expand=Member/Users,RoleDefinitionBindings"
    $ras = Invoke-SPRest -Url $raUrl
    foreach ($ra in $ras) {
        $roleLabel = "Specific Access"
        foreach ($rd in $ra.RoleDefinitionBindings.results) {
            if ($rd.Name -eq "Limited Access") { 
                $roleLabel = "Limit Access"; break 
            }
        }

        if ($ra.Member.PrincipalType -eq 1) { & $AddUser $ra.Member $roleLabel }
        else {
            foreach ($gu in $ra.Member.Users.results) { & $AddUser $gu $roleLabel }
        }
    }

    # Format for Export
    return ForEach-Object ($key in $userMap.Keys) {
        $u = $userMap[$key]
        if ($u.LoginID -match "sharepoint\\system|NT AUTHORITY") { continue }
        [PSCustomObject]@{
            UserName = $u.UserName
            LoginID  = $u.LoginID
            Email    = $u.Email
            Access   = ($u.Roles | Sort-Object) -join "; "
        }
    }
}

# 5. Main Execution Loop
try {
    $siteCollections = Get-SPWebApplication $WebApplicationUrl | Get-SPSite -Limit ALL
    
    foreach ($site in $siteCollections) {
        $webs = @($site.RootWeb)
        if ($IncludeSubsites) { $webs += $site.AllWebs | Where-Object { $_.Url -ne $site.RootWeb.Url } }

        foreach ($web in $webs) {
            Write-Host "Processing: $($web.Url)" -ForegroundColor Cyan
            
            $isRoot = ($web.Url -eq $site.Url)
            $userData = Get-SiteUserPermissions -WebUrl $web.Url -IsSiteRoot $isRoot

            if ($userData) {
                # Filename based on sanitized URL
                $safeName = $web.Url -replace "https?://", "" -replace "[/\\:*?`"<>|]", "_"
                $filePath = Join-Path $OutputFolder "$safeName.csv"
                
                $userData | Export-Csv -Path $filePath -NoTypeInformation -Encoding UTF8
                Write-Host "  -> Exported to $safeName.csv" -ForegroundColor Green
            }
            $web.Dispose()
        }
        $site.Dispose()
    }
}
catch {
    Write-Error "Failed to process Web Application: $_"
}

Write-Host "`nProcess Complete. Files saved to: $OutputFolder" -ForegroundColor Yellow