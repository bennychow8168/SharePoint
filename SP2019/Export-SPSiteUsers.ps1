param(
    [Parameter(Mandatory = $true)]
    [string]$WebApplicationUrl,

    [Parameter(Mandatory = $false)]
    [switch]$IncludeSubsites,

    [Parameter(Mandatory = $true)]
    [string]$OutputDirectory,

    [Parameter(Mandatory = $false)]
    [System.Management.Automation.PSCredential]$Credential
)

Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue

$ErrorActionPreference = "Stop"

if (-not (Test-Path -LiteralPath $OutputDirectory)) {
    New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
}

function Invoke-SPRestGet {
    param(
        [Parameter(Mandatory = $true)][string]$Url,
        [Parameter(Mandatory = $false)][System.Management.Automation.PSCredential]$Credential
    )

    $headers = @{
        "Accept" = "application/json;odata=verbose"
        "Content-Type" = "application/json;odata=verbose"
    }

    if ($Credential) {
        return Invoke-RestMethod -Uri $Url -Method Get -Headers $headers -Credential $Credential
    }
    else {
        return Invoke-RestMethod -Uri $Url -Method Get -Headers $headers -UseDefaultCredentials
    }
}

function Get-AllRestItems {
    param(
        [Parameter(Mandatory = $true)][string]$Url,
        [Parameter(Mandatory = $false)][System.Management.Automation.PSCredential]$Credential
    )

    $items = @()
    $nextUrl = $Url

    while ($nextUrl) {
        $response = Invoke-SPRestGet -Url $nextUrl -Credential $Credential

        if ($response.d.results) {
            $items += $response.d.results
            $nextUrl = $response.d.__next
        }
        elseif ($response.d) {
            $items += $response.d
            $nextUrl = $null
        }
        else {
            $nextUrl = $null
        }
    }

    return $items
}

function Test-IsSystemAccount {
    param(
        [Parameter(Mandatory = $true)]$UserObject
    )

    $login = [string]$UserObject.LoginName
    $title = [string]$UserObject.Title
    $email = [string]$UserObject.Email

    if ([string]::IsNullOrWhiteSpace($login) -and [string]::IsNullOrWhiteSpace($title)) { return $true }

    $patterns = @(
        "^SHAREPOINT\\system$",
        "^NT AUTHORITY\\",
        "^IUSR",
        "^IWAM",
        "^svc[_\-]",
        "^sp_",
        "^app@sharepoint$",
        "^i:0#.w\|",
        "^c:0\..+",
        "System Account"
    )

    foreach ($pattern in $patterns) {
        if ($login -match $pattern -or $title -match $pattern -or $email -match $pattern) {
            return $true
        }
    }

    return $false
}

function ConvertTo-SafeFileName {
    param(
        [Parameter(Mandatory = $true)][string]$Name
    )

    $invalidChars = [System.IO.Path]::GetInvalidFileNameChars()
    $safeName = $Name

    foreach ($char in $invalidChars) {
        $safeName = $safeName.Replace($char, "_")
    }

    $safeName = $safeName -replace "\s+", "_"
    return $safeName
}

function Get-WebUsersAndPermissions {
    param(
        [Parameter(Mandatory = $true)][string]$WebUrl,
        [Parameter(Mandatory = $false)][System.Management.Automation.PSCredential]$Credential
    )

    $siteUsersUrl = "$WebUrl/_api/web/siteusers?`$select=Id,Title,LoginName,Email,PrincipalType,IsSiteAdmin"
    $roleAssignmentsUrl = "$WebUrl/_api/web/roleassignments?`$expand=Member,RoleDefinitionBindings"

    $siteUsers = Get-AllRestItems -Url $siteUsersUrl -Credential $Credential
    $roleAssignments = Get-AllRestItems -Url $roleAssignmentsUrl -Credential $Credential

    $permissionMap = @{}

    foreach ($ra in $roleAssignments) {
        if (-not $ra.Member) { continue }

        $memberLogin = [string]$ra.Member.LoginName
        if ([string]::IsNullOrWhiteSpace($memberLogin)) { continue }

        $roleNames = @()

        if ($ra.RoleDefinitionBindings -and $ra.RoleDefinitionBindings.results) {
            $roleNames = $ra.RoleDefinitionBindings.results |
                Where-Object { $_.Name -and $_.Name -ne "Limited Access" } |
                Select-Object -ExpandProperty Name
        }

        if (-not $roleNames -or $roleNames.Count -eq 0) { continue }

        if (-not $permissionMap.ContainsKey($memberLogin)) {
            $permissionMap[$memberLogin] = New-Object System.Collections.Generic.List[string]
        }

        foreach ($roleName in $roleNames) {
            if (-not $permissionMap[$memberLogin].Contains($roleName)) {
                [void]$permissionMap[$memberLogin].Add($roleName)
            }
        }
    }

    $results = foreach ($user in $siteUsers) {
        if ($user.PrincipalType -ne 1) { continue }   # 1 = User
        if (Test-IsSystemAccount -UserObject $user) { continue }

        $permissionLevel = ""
        if ($permissionMap.ContainsKey([string]$user.LoginName)) {
            $permissionLevel = ($permissionMap[[string]$user.LoginName] | Sort-Object) -join "; "
        }

        [PSCustomObject]@{
            WebUrl           = $WebUrl
            UserName         = [string]$user.Title
            LoginID          = [string]$user.LoginName
            Email            = [string]$user.Email
            PermissionLevel  = $permissionLevel
            IsSiteAdmin      = [string]$user.IsSiteAdmin
        }
    }

    return $results
}

function Get-SubWebUrls {
    param(
        [Parameter(Mandatory = $true)][string]$WebUrl,
        [Parameter(Mandatory = $false)][System.Management.Automation.PSCredential]$Credential
    )

    $subwebs = @()
    $websUrl = "$WebUrl/_api/web/webs?`$select=Title,Url"

    $childWebs = Get-AllRestItems -Url $websUrl -Credential $Credential

    foreach ($child in $childWebs) {
        $subwebs += [string]$child.Url
        $subwebs += Get-SubWebUrls -WebUrl ([string]$child.Url) -Credential $Credential
    }

    return $subwebs
}

$webApp = Get-SPWebApplication $WebApplicationUrl
$siteCollections = $webApp.Sites

foreach ($site in $siteCollections) {
    try {
        $rootWebUrl = $site.RootWeb.Url
        Write-Host "Processing site collection: $rootWebUrl" -ForegroundColor Cyan

        $targetWebs = @($rootWebUrl)

        if ($IncludeSubsites) {
            $targetWebs += Get-SubWebUrls -WebUrl $rootWebUrl -Credential $Credential
        }

        $allRows = foreach ($webUrl in ($targetWebs | Select-Object -Unique)) {
            Write-Host "  Reading: $webUrl" -ForegroundColor Yellow
            Get-WebUsersAndPermissions -WebUrl $webUrl -Credential $Credential
        }

        $siteCollectionName = ConvertTo-SafeFileName -Name $site.Url
        $outputFile = Join-Path $OutputDirectory "$siteCollectionName.csv"

        $allRows |
            Sort-Object WebUrl, UserName, LoginID |
            Export-Csv -Path $outputFile -NoTypeInformation -Encoding UTF8

        Write-Host "  Exported: $outputFile" -ForegroundColor Green
    }
    catch {
        Write-Warning "Failed site collection [$($site.Url)]: $($_.Exception.Message)"
    }
    finally {
        if ($site.RootWeb -ne $null) { $site.RootWeb.Dispose() }
        $site.Dispose()
    }
}