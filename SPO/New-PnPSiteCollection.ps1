param(
    # Publishing or Team (no M365 Group)
    [Parameter(Mandatory=$true)] [ValidateSet("Prod","Production","QA","RND")] [string]$Environment,
    [Parameter(Mandatory=$false)] [ValidateSet("Communication","Teams")] [string]$SiteType = "Communication",
    [Parameter(Mandatory=$true)] [string]$SiteUrl,  # e.g. https://tenant.sharepoint.com/sites/Intranet
    [Parameter(Mandatory=$true)] [string]$Title,
    [Parameter(Mandatory=$true)] [string]$InitialOwner,
    [Parameter(Mandatory=$false)] [ValidateSet("English","Japanese")] [string]$DefaultLocale = "English",   # Site locale, default value is English
    [Parameter(Mandatory=$false)] [ValidateSet(1024,1024000)] [int]$SiteQuota = "51200",    # Site quota by MB (default value is 51200MB)
    [Parameter(Mandatory=$false)] [switch]$GrantEveryoneVisitors = $False,
    [Parameter(Mandatory=$false)] [ValidateSet("Intranet","StandAloneSite","Project","Division")] [string]$SiteCatalogValue = "StandAloneSite",
    [Parameter(Mandatory=$false)] [string]$SiteDivisionValue,
    [Parameter(Mandatory=$false)] [string]$OwnerDivision,
    [Parameter(Mandatory=$false)] [string]$TimeZone = 24,
    [Parameter(Mandatory=$false)] [switch]$EnableHubSite = $False,
    [Parameter(Mandatory=$false)] [string]$HubeSiteAssocation = [string]::IsNullOrEmpty
)

# =========================
# Config - admin & template
# =========================

# Map logical types to templates (Modern)
# Publishing root site
$publishingTemplate   = 'SITEPAGEPUBLISHING#0'        # Commuinication site [web:22]
# Team site without group
$teamNoGroupTemplate  = 'STS#3'                       # Modern teams site [web:22]

switch ($SiteType) {
    "Communication"   { $templateToUse = $publishingTemplate }
    "Teams"  { $templateToUse = $teamNoGroupTemplate }
}

# Locale IDs [common SPO values]
$lcidEnglish  = 1033
$lcidJapanese = 1041

if ($DefaultLocale -eq "English") {
    $defaultLcid = $lcidEnglish
    $secondaryLcid = $lcidJapanese
}
else {
    $defaultLcid = $lcidJapanese
    $secondaryLcid = $lcidEnglish
}

# TimeZoneId – adjust to your region
$timeZoneId = 13    # (UTC+08:00) Beijing, Chongqing, Hong Kong, Urumqi, etc. [web:19]

# Theme name
$themeName = "CompanyColor"           # 3. Pre-created tenant theme [web:21]

# Roles to auto-assign as owners/site admins
$pinSharePointAdminRole = "SharePoint Administrator"
$pinGlobalAdminRole     = "Global Administrator"

# =============================
# Helper: ensure SPO connection
# =============================

function Ensure-SPOConnection {
    param(
        [Parameter(Mandatory=$true)][string]$ExpectedTenantAdminUrl
    )

    $currentContext = Get-SPOContext  -ErrorAction SilentlyContinue  # not a real cmdlet; emulate via Get-SPOTenant
    $needConnect = $true

    try {
        $tenant = Get-SPOTenant -ErrorAction Stop   # requires Microsoft.Online.SharePoint.PowerShell [web:28]
        if ($tenant.AdminCenterUrl -eq $ExpectedTenantAdminUrl) {
            $needConnect = $false
        }
    } catch {
        $needConnect = $true
    }
    if ($needConnect) {
        Connect-SPOService -Url $ExpectedTenantAdminUrl
    }
}

function Ensure-PNPAdminConnection {
    param(
        [Parameter(Mandatory=$true)][string]$ExpectedTenantAdminUrl
    )

    # Ensure PnP connection to admin center
    $connection = Get-PnPConnection -ErrorAction SilentlyContinue
    if (-not $connection -or $connection.Url -ne $ExpectedTenantAdminUrl) {
        Write-Host "Connecting to SharePoint Online admin center..."
        #Connect-PnPOnline -Url $adminUrl -Interactive
        Connect-PnPOnline -Url $ExpectedTenantAdminUrl -UseWebLogin
        $connection = Get-PnPConnection
        if ($connection.Url -ne $ExpectedTenantAdminUrl) {
            throw "Connected PnP context is not the expected admin URL: $ExpectedTenantAdminUrl"
        }
    }
}

# ==============================
# Helper: ensure PnP connection
# ==============================

function Ensure-PnPConnection {
    param(
        [Parameter(Mandatory=$true)][string]$TargetUrl
    )

    $conn = Get-PnPConnection -ErrorAction SilentlyContinue
    if (-not $conn -or $conn.Url -ne $TargetUrl) {
        #Connect-PnPOnline -Url $TargetUrl -Interactive  # PnP.PowerShell 1.12.0 supports -Interactive [web:16]
        Connect-PnPOnline -Url $TargetUrl -UseWebLogin  # PnP.PowerShell 1.12.0 supports -UseWebLogin [web:16]
    }
}

# =========================
# Main logic
# =========================

# Derive tenant admin URL from site URL
# e.g. https://tenant-admin.sharepoint.com
$uri = [Uri]$SiteUrl
$adminHost = $uri.Host.Replace(".sharepoint.com",".sharepoint.com") -replace "^([^\.]+)",'$1-admin'
$tenantAdminUrl = "https://$adminHost"

# 0. Ensure modules loaded
Import-Module Microsoft.Online.SharePoint.PowerShell -ErrorAction Stop
Import-Module PnP.PowerShell -RequiredVersion 1.12.0 -ErrorAction Stop  # ensure required version [web:16]

# 1. Ensure connected to correct tenant (SPO service)
#Ensure-SPOConnection -ExpectedTenantAdminUrl $tenantAdminUrl
Ensure-PNPAdminConnection -ExpectedTenantAdminUrl $tenantAdminUrl

# 2. Create site collection (classic) using New-SPOSite [web:22]
# Owner will be updated later with PIN roles
#$initialOwner = "temp.owner@yourtenant.onmicrosoft.com"  # substitute a service account if required

# Identifying site template
if ($SiteType -eq "Communication"){
    $SiteTemplate = "CommunicationSite"
} elseif ($SiteType -eq "Teams") {
    $SiteTemplate = "TeamSiteWithoutMicrosoft365Group"
} else {
    Write-Host "SiteType only accept [Communication] or [Teams]."
    Write-Host "Existing."
}
if ($DefaultLocale -eq "English") {
    $lcid = 1033
} elseif ($DefaultLocale -eq "Japanese") {
    $lcid = 1041
} else {
    Write-Host "DefaultLocale only accept [English] or [Japanese]."
    Write-Host "Existing."
}

# PnP.PowerShell method
if (-not ($existingSite = Get-PnPTenantSite -Identity $SiteUrl -ErrorAction SilentlyContinue)) {
    New-PnPSite -Type $SiteTemplate -Title $SiteTitle -Url $SiteUrl -Lcid $lcid -Owner $InitialOwner -Description "Created by provisioning script" -Wait
}
<#-- SharePoint.Online.Management.Shell method
if (-not (Get-SPOSite -Identity $SiteUrl -ErrorAction SilentlyContinue)) {
    [int]$StorageQuotaWarningLevel = $SiteQuota * 0.95
    New-SPOSite `
        -Url $SiteUrl -Title $Title -Owner $initialOwner -Template $templateToUse -LocaleId $defaultLcid -TimeZoneId $TimeZone -StorageQuota $SiteQuota -StorageQuotaWarningLevel $StorageQuotaWarningLevel
    Set-SPOWebTheme -Theme $themeName -Web $SiteUrl
}
#>

# Connect PnP to the new site
Ensure-PnPConnection -TargetUrl $SiteUrl

Add-PnPSiteCollectionAdmin -Owners @("SP Admin", "GA")

# Set site elements
[int]$StorageQuotaWarningLevel = $SiteQuota * 0.95
Set-PnPSite -StorageMaximumLevel $SiteQuota -StorageWarningLevel $StorageQuotaWarningLevel


# 4. Regional settings & languages (default + alternate) [web:19][web:25]
$web = Get-PnPWeb -Includes RegionalSettings,SupportedUILanguages

# Set default locale
$web.RegionalSettings.LocaleId = $defaultLcid

# Add secondary UI language if needed
$existingLangs = $web.SupportedUILanguages | ForEach-Object { $_.LCID }
if ($secondaryLcid -notin $existingLangs) {
    $web.SupportedUILanguages.Add($secondaryLcid) | Out-Null
}
$web.Update()
Invoke-PnPQuery

# Apply theme "CompanyColor"
Set-PnPWebTheme -Theme $themeName

if ($SiteCatalogValue -eq "Intranet") {
    # Grant "Everyone except external users" as Visitors (optional) [web:24][web:30]
    if ($GrantEveryoneVisitors) {
        # Ensure Everyone claim visible at tenant level:
        # Set-SPOTenant -ShowEveryoneClaim $true  # run once at tenant if needed [web:24]
        #$visitorsGroupName = (Get-PnPWeb).Title + " Visitors"
        $visitorsGroupName = Get-PnPGroup -AssociatedVisitorGroup
        $visitorsGroup = Get-PnPGroup -Identity $visitorsGroupName -ErrorAction SilentlyContinue
        if ($visitorsGroup) {
            # Add Everyone except external users claim
            $everyoneClaim = "c:0-.f|rolemanager|spo-grid-all-users/{tenantid}"  # adjust if you use a known claim pattern
            Add-PnPGroupMember -Identity $visitorsGroup -LoginName $everyoneClaim
        }
        Add-PnPSiteCollectionAdmin -PrimarySiteCollectionAdmin '<Admin GUID>'
    }
    # Create “Publisher” group with Edit permission & access to Usage/Reports [web:20][web:32]
    $publisherRoleDefinition = "Publisher"
    $publisherGroupName = "Publisher"
    # Get existing Designer role definition (name is localized; adjust if needed)
    $designerRole = Get-PnPRoleDefinition -Identity "Design"

    # Create new permission level cloned from Designer, and ensure ViewUsageData is included
    Add-PnPRoleDefinition -RoleName $publisherRoleDefinition -Description "Designer permissions plus ViewUsageData." -Clone $designerRole -Include ViewUsageData
    $publisherGroup = Get-PnPGroup -Identity $publisherGroupName -ErrorAction SilentlyContinue

    if (-not $publisherGroup) {
        New-PnPGroup -Title $publisherGroupName -Description "Publisher group with Edit permission and access to reports."
        $publisherGroup = Get-PnPGroup -Identity $publisherGroupName
    }

    # Ensure Publisher permission level
    Set-PnPGroupPermissions -Identity $publisherGroupName -AddRole $publisherRoleDefinition
}




# Grant access to site usage / reports:
# Typically included with 'Edit' on site; if using custom permission level for reports, add here too.
# Example: Set-PnPGroupPermissions -Identity $publisherGroupName -AddRole "View Usage Data"

# 8. Set default owners to PIN roles (SharePoint Admin & Global Admin)
#    Assumes you have mapped these roles to security groups / AAD groups that can be resolved in SPO

$ownerGroupName = (Get-PnPWeb).Title + " Owners"
$ownerGroup = Get-PnPGroup -Identity $ownerGroupName -ErrorAction SilentlyContinue

if ($ownerGroup) {
    # Replace or add role-based groups as owners
    Add-PnPGroupMember -Identity $ownerGroup -LoginName $pinSharePointAdminRole
    Add-PnPGroupMember -Identity $ownerGroup -LoginName $pinGlobalAdminRole
}

# Also make them site collection admins
$adminLogins = @(
    $pinSharePointAdminRole,
    $pinGlobalAdminRole
)

foreach ($login in $adminLogins) {
    Add-PnPSiteCollectionAdmin -Owners $login -ErrorAction SilentlyContinue
}

Remove-PnPUser -Identity $InitialOwner

Write-Host "Site created and configured: $SiteUrl"
