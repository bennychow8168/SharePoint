param(
    [Parameter(Mandatory)] $SiteURL,
    [Parameter(Mandatory)] $SiteTitle,
    [Parameter(Mandatory)] $SiteOwner,
    [Parameter(Mandatory)] $SiteQuota = "51200",
    [Parameter(Mandatory)] $Template = "STS#3", #Modern Team Site
    [Parameter(Mandatory)] $Timezone = "24"
)


If ($SiteURL -like "*.sharepoint.com/sites/*") {
    $tenant = ($site.Split('/')[2]).Split('.')[0]
    $AdminCenterURL = "https://" + $tenant + "-admin.sharepoint.com"
    $AdminCenterURL
} Else {
    Write-Host "URL error"
}

Try
{
    #Connect to Tenant Admin
    Connect-PnPOnline -URL $AdminCenterURL -UseWebLogin
     
    #Check if site exists already
    $SiteURL = Get-PnPTenantSite | Where-Object {$_.Url -eq $SiteURL}
 
    If ($SiteURL -eq $null) {
        #sharepoint online pnp powershell create a new team site collection
        New-PnPTenantSite -Url $SiteURL -Owner $SiteOwner -Title $SiteTitle -Template $Template -TimeZone $TimeZone -StorageQuota $SiteQuota -StorageQuotaWarningLevel $SiteQuota -RemoveDeletedSite
        write-host "Site Collection $($SiteURL) Created Successfully!" -foregroundcolor Green

        # Appplying theme
        # Connect to SharePoint Online Site
        $SiteConn = Connect-PnPOnline -Url $SiteURL -UseWebLogin -ReturnConnection
 
        # Get all the Webs - Exclude App Sites
        $Webs = Get-PnPSubWeb -Recurse -IncludeRootWeb -Connection $SiteConn | Where-Object {$_.WebTemplate -ne "App"}
 
        # Call the function to set site theme for site collection
        $Webs | ForEach-Object { Set-PnPSiteTheme -Web $_ -ThemeName "NomuraColor" }

        # Insatll SPFx into site
        $AppName = "Modern Script Editor web part by Puzzlepart"
 
        #Get the App from App Catalog
        $App = Get-PnPApp -Scope Tenant | Where-Object {$_.Title -eq $AppName}
        
        #Install App to the Site
        Install-PnPApp -Identity $App.Id

        # Activate the SPFx Application Customizer to the Site through Custom Action
        $customCSSUrl = "https://crescent.sharepoint.com/sites/Marketing/Style%20Library/custom.css"
 
        #Add Custom Action
        Add-PnPCustomAction -Title "Inject CSS Application Extension" -Name "InjectCssApplicationCustomizer" -Location "ClientSideExtension.ApplicationCustomizer" -ClientSideComponentId "5a1fcffd-dfeb-4844-b478-1feb4325a5a7" -ClientSideComponentProperties "{""cssurl"":""$customCSSUrl""}"
    }
    else  {
        write-host "Site $($SiteURL) exists already!" -foregroundcolor Yellow
    }
}
catch {
    write-host "Error: $($_.Exception.Message)" -foregroundcolor Red
}