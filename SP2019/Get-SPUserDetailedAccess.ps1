Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue

param(
    [Parameter(Mandatory=$true)]
    [string]$UserLogin,   # Example: CONTOSO\jdoe or claims login

    [string]$WebApplicationUrl,
    [string]$SiteCollectionUrl,
    [string]$OutputPath = "C:\Temp\SP2019-UserAccess-Detailed.csv",

    [switch]$ScanItems,
    [switch]$IncludeHiddenLists
)

$results = New-Object System.Collections.Generic.List[Object]

function Add-Result {
    param(
        [string]$ScopeType,
        [string]$WebApplication,
        [string]$SiteCollection,
        [string]$WebUrl,
        [string]$ObjectTitle,
        [string]$ObjectUrl,
        [string]$AccessType,
        [string]$GrantedThrough,
        [string]$PermissionLevels,
        [bool]$HasUniquePermissions
    )

    $results.Add([PSCustomObject]@{
        ScopeType            = $ScopeType
        WebApplication       = $WebApplication
        SiteCollection       = $SiteCollection
        WebUrl               = $WebUrl
        ObjectTitle          = $ObjectTitle
        ObjectUrl            = $ObjectUrl
        AccessType           = $AccessType
        GrantedThrough       = $GrantedThrough
        PermissionLevels     = $PermissionLevels
        HasUniquePermissions = $HasUniquePermissions
    })
}

function Test-PrincipalMatch {
    param(
        [Microsoft.SharePoint.SPPrincipal]$Principal,
        [Microsoft.SharePoint.SPUser]$TargetUser
    )

    if ($null -eq $Principal -or $null -eq $TargetUser) { return $false }

    if ($Principal -is [Microsoft.SharePoint.SPUser]) {
        return ($Principal.LoginName -eq $TargetUser.LoginName)
    }

    if ($Principal -is [Microsoft.SharePoint.SPGroup]) {
        foreach ($u in $Principal.Users) {
            if ($u.LoginName -eq $TargetUser.LoginName) {
                return $true
            }
        }
    }

    if ($Principal.LoginName -eq $TargetUser.LoginName -or $Principal.Name -eq $TargetUser.LoginName) {
        return $true
    }

    return $false
}

function Get-PermissionLevels {
    param($RoleAssignment)
    return (($RoleAssignment.RoleDefinitionBindings | Select-Object -ExpandProperty Name) -join "; ")
}

function CheckSecurableObject {
    param(
        [string]$ScopeType,
        [object]$Object,
        [string]$Title,
        [string]$Url,
        [Microsoft.SharePoint.SPUser]$TargetUser,
        [string]$WebApplicationUrl,
        [string]$SiteCollectionUrl,
        [string]$WebUrl,
        [bool]$HasUniquePermissions
    )

    $matched = $false

    try {
        foreach ($ra in $Object.RoleAssignments) {
            $permLevels = Get-PermissionLevels -RoleAssignment $ra
            $principal = $ra.Member

            if ($principal -is [Microsoft.SharePoint.SPUser]) {
                if ($principal.LoginName -eq $TargetUser.LoginName) {
                    Add-Result `
                        -ScopeType $ScopeType `
                        -WebApplication $WebApplicationUrl `
                        -SiteCollection $SiteCollectionUrl `
                        -WebUrl $WebUrl `
                        -ObjectTitle $Title `
                        -ObjectUrl $Url `
                        -AccessType "Direct" `
                        -GrantedThrough $principal.LoginName `
                        -PermissionLevels $permLevels `
                        -HasUniquePermissions $HasUniquePermissions
                    $matched = $true
                }
            }
            elseif ($principal -is [Microsoft.SharePoint.SPGroup]) {
                if (Test-PrincipalMatch -Principal $principal -TargetUser $TargetUser) {
                    Add-Result `
                        -ScopeType $ScopeType `
                        -WebApplication $WebApplicationUrl `
                        -SiteCollection $SiteCollectionUrl `
                        -WebUrl $WebUrl `
                        -ObjectTitle $Title `
                        -ObjectUrl $Url `
                        -AccessType "SharePoint Group" `
                        -GrantedThrough $principal.Name `
                        -PermissionLevels $permLevels `
                        -HasUniquePermissions $HasUniquePermissions
                    $matched = $true
                }
            }
            else {
                if (Test-PrincipalMatch -Principal $principal -TargetUser $TargetUser) {
                    Add-Result `
                        -ScopeType $ScopeType `
                        -WebApplication $WebApplicationUrl `
                        -SiteCollection $SiteCollectionUrl `
                        -WebUrl $WebUrl `
                        -ObjectTitle $Title `
                        -ObjectUrl $Url `
                        -AccessType "Principal Match" `
                        -GrantedThrough $principal.Name `
                        -PermissionLevels $permLevels `
                        -HasUniquePermissions $HasUniquePermissions
                    $matched = $true
                }
            }
        }
    }
    catch {
    }

    return $matched
}

function Get-ItemDisplayUrl {
    param(
        [Microsoft.SharePoint.SPWeb]$Web,
        [Microsoft.SharePoint.SPListItem]$Item
    )

    try {
        if ($Item.Url) {
            if ($Item.Url.StartsWith("http")) { return $Item.Url }
            return ($Web.Url.TrimEnd("/") + "/" + $Item.Url.TrimStart("/"))
        }
    }
    catch {
    }

    return $null
}

# Determine site collection scope
$sitesToProcess = @()

if ($SiteCollectionUrl) {
    $sitesToProcess = @(Get-SPSite $SiteCollectionUrl)
}
elseif ($WebApplicationUrl) {
    $sitesToProcess = Get-SPSite -WebApplication $WebApplicationUrl -Limit All
}
else {
    $sitesToProcess = Get-SPSite -Limit All
}

foreach ($site in $sitesToProcess) {
    try {
        $waUrl = $site.WebApplication.Url

        foreach ($web in $site.AllWebs) {
            try {
                $targetUser = $null

                try {
                    $targetUser = Get-SPUser -Web $web -Identity $UserLogin -ErrorAction Stop
                }
                catch {
                    $targetUser = $null
                }

                if (-not $targetUser) {
                    continue
                }

                # Web-level permissions
                $webMatched = CheckSecurableObject `
                    -ScopeType "Web" `
                    -Object $web `
                    -Title $web.Title `
                    -Url $web.Url `
                    -TargetUser $targetUser `
                    -WebApplicationUrl $waUrl `
                    -SiteCollectionUrl $site.Url `
                    -WebUrl $web.Url `
                    -HasUniquePermissions $web.HasUniqueRoleAssignments

                if (-not $webMatched) {
                    try {
                        if ($web.DoesUserHavePermissions($targetUser.LoginName, [Microsoft.SharePoint.SPBasePermissions]::Open)) {
                            Add-Result `
                                -ScopeType "Web" `
                                -WebApplication $waUrl `
                                -SiteCollection $site.Url `
                                -WebUrl $web.Url `
                                -ObjectTitle $web.Title `
                                -ObjectUrl $web.Url `
                                -AccessType "Inherited / Effective Access" `
                                -GrantedThrough "Parent scope or AD/security group" `
                                -PermissionLevels "Open" `
                                -HasUniquePermissions $web.HasUniqueRoleAssignments
                        }
                    }
                    catch {
                    }
                }

                # List / library permissions
                foreach ($list in $web.Lists) {
                    try {
                        if (-not $IncludeHiddenLists -and $list.Hidden) { continue }

                        $listUrl = $null
                        try {
                            $listUrl = $web.Url.TrimEnd("/") + "/" + $list.RootFolder.Url.TrimStart("/")
                        }
                        catch {
                            $listUrl = $web.Url
                        }

                        $listMatched = $false

                        if ($list.HasUniqueRoleAssignments) {
                            $listMatched = CheckSecurableObject `
                                -ScopeType "List/Library" `
                                -Object $list `
                                -Title $list.Title `
                                -Url $listUrl `
                                -TargetUser $targetUser `
                                -WebApplicationUrl $waUrl `
                                -SiteCollectionUrl $site.Url `
                                -WebUrl $web.Url `
                                -HasUniquePermissions $true
                        }
                        else {
                            try {
                                if ($web.DoesUserHavePermissions($targetUser.LoginName, [Microsoft.SharePoint.SPBasePermissions]::ViewListItems)) {
                                    Add-Result `
                                        -ScopeType "List/Library" `
                                        -WebApplication $waUrl `
                                        -SiteCollection $site.Url `
                                        -WebUrl $web.Url `
                                        -ObjectTitle $list.Title `
                                        -ObjectUrl $listUrl `
                                        -AccessType "Inherited / Effective Access" `
                                        -GrantedThrough "Inherited from web" `
                                        -PermissionLevels "ViewListItems" `
                                        -HasUniquePermissions $false
                                    $listMatched = $true
                                }
                            }
                            catch {
                            }
                        }

                        # Optional item/folder scanning
                        if ($ScanItems) {
                            foreach ($item in $list.Items) {
                                try {
                                    if (-not $item.HasUniqueRoleAssignments) { continue }

                                    $itemTitle = $null
                                    try {
                                        $itemTitle = $item["Title"]
                                    }
                                    catch {
                                    }

                                    if ([string]::IsNullOrWhiteSpace($itemTitle)) {
                                        try { $itemTitle = $item.Name } catch {}
                                    }

                                    if ([string]::IsNullOrWhiteSpace($itemTitle)) {
                                        $itemTitle = "Item ID $($item.ID)"
                                    }

                                    $itemUrl = Get-ItemDisplayUrl -Web $web -Item $item

                                    [void](CheckSecurableObject `
                                        -ScopeType "Item/Folder" `
                                        -Object $item `
                                        -Title $itemTitle `
                                        -Url $itemUrl `
                                        -TargetUser $targetUser `
                                        -WebApplicationUrl $waUrl `
                                        -SiteCollectionUrl $site.Url `
                                        -WebUrl $web.Url `
                                        -HasUniquePermissions $true)
                                }
                                catch {
                                }
                            }
                        }
                    }
                    catch {
                    }
                }
            }
            finally {
                $web.Dispose()
            }
        }
    }
    finally {
        $site.Dispose()
    }
}

$results |
    Sort-Object WebApplication, SiteCollection, WebUrl, ScopeType, ObjectTitle |
    Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8

Write-Host "Done. Report saved to $OutputPath" -ForegroundColor Green