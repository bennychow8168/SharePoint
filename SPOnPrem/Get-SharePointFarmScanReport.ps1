#######################################################
#Add Add-PSSnapin Microsoft.SharePoint.PowerShell
Set-ExecutionPolicy "Unrestricted"
#Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction Stop
Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue
#######################################################
function SPFramScanReport()
  { 

		Try
		{
			Write-Host "SharePoint Farm Scan Report" -ForegroundColor cyan
            Write-Host "--------------------------------------------------------------------"
            $ex = "Exceed the Supported Limit"
            $wa = "Within the Limit"
            #SharePoint Web Application
            $SPWebApp = Get-SPWebApplication 
            $SPWebAppcount = $SPWebApp.count
            #SharePoint Web Application Pool
            $SPWebAppPool = Get-SPWebApplication | Select-Object ApplicationPool -Unique
            #SharePoint Service Application Pool
            $SPSrvWebAppPool = Get-SPServiceApplicationPool  | Select-Object -Unique
            $SPAppPoolCount = $SPSrvWebAppPool.count + $SPWebAppPool.count
            #SharePoint Service Applications
            $SPSrvApp = Get-SPServiceApplication | Select-Object id,name
            $SPSrvAppcount= $SPSrvApp.count
            # Content Databases
            $ContentDB = Get-SPContentDatabase | Select-Object name,WebApplication,@{Name="Contenet Database Size (GB)"; Expression={[math]::Round($_.disksizerequired/1024MB,2)}},@{Name="Within the Limit"; Expression={if([math]::Round($_.disksizerequired/1024MB,2) -ge 200){ "No"} else {"Yes"}}}
            $ContentDBcount = $ContentDB.count
            # Site Collections
            $SiteCollections = Get-SPSite | Select-Object url,contentdatabase,webapplication,@{Name="Site Collection Size (MB)";Expression={[math]::Round($_.usage.storage/(1MB),2)}},@{Name="Within the Limit"; Expression={if([math]::Round($_.usage.storage/1MB,2) -ge 100*1024){ "No"} else {"Yes"}}} | Format-List
            $SiteCollectionscount = (Get-SPSite).count
            # SubSites
            $Subsites = Get-SPSite | Get-SPWeb -Limit All
            $Subsitescount= $Subsites.count
            
            ##################################################################################
            #How Many Web Applications in farm 
            Write-Host "SharePoint Web Applications per Farm" -ForegroundColor cyan
            Write-Host "The supported limit for the SharePoint web application per farm is 20 web applications."
            switch($SPWebAppcount)
            {
            {$_ -ge 20} {Write-Host "Total Number of SharePoint Web Application:" $_ "|"$ex -ForegroundColor red }
            {$_ -lt 20} {Write-Host "Total Number of SharePoint Web Application:" $_ "|"$wa -ForegroundColor yellow}
            }
            $SPWebApp
            Write-Host "For more details, please check https://spgeeks.devoworx.com/sharepoint-2019-limitations" -ForegroundColor cyan
            Write-Host "--------------------------------------------------------------------"


            ##################################################################################
            #How Many SharePoint Web Application Pools for the web server Per farm 
            Write-Host "SharePoint Web Application Pools for the web server Per farm" -ForegroundColor cyan
            Write-Host "SharePoint Application Pool Limits for Web Server per Farm is 10 Application Pools (This limit depends mainly on the Server Hardware specifications)."
            switch($SPAppPoolCount)
            {
            {$_ -ge 10} {Write-Host "Total Number of SharePoint Application Pool:" $_ "|"$ex -ForegroundColor red }
            {$_ -lt 10} {Write-Host "Total Number of SharePoint Application Pool:" $_ "|"$wa -ForegroundColor yellow}
            }
            $SPWebAppPool
            $SPSrvWebAppPool | Select-Object name
            Write-Host "For more details, please check https://spgeeks.devoworx.com/sharepoint-2019-service-accounts-best-practice/" -ForegroundColor cyan
            Write-Host "--------------------------------------------------------------------"

            
            ##################################################################################
            #How Many SharePoint Service Applications Running on farm
            Write-Host "Running SharePoint Service Applications" -ForegroundColor cyan 
            Write-Host "Total Number of Running SharePoint Service Applications:" $SPSrvApp.count -ForegroundColor Green
            $SPSrvApp 
            Write-Host "--------------------------------------------------------------------"

            
            ##################################################################################
            #How Many SharePoint Content Database Per farm 
            Write-Host "SharePoint Content Database Per farm" -ForegroundColor cyan
            Write-Host "The supported limit for SharePoint Content Databases per farm is 500 content databases."
            switch($ContentDBcount)
            {
            {$_ -ge 500} {Write-Host "Total Number of SharePoint Content Database:" $_ "|"$ex -ForegroundColor red }
            {$_ -lt 500} {Write-Host "Total Number of SharePoint Content Database:" $_ "|"$wa -ForegroundColor yellow}
            }
            $ContentDB
            Write-Host "For more details, please check https://spgeeks.devoworx.com/sql-server-best-practices-sharepoint/" -ForegroundColor cyan
            Write-Host "--------------------------------------------------------------------"


            ##################################################################################
            #How Many SharePoint Content Database Per Web Application
            Write-Host "SharePoint Content Database Per Web Application" -ForegroundColor cyan
            foreach($WebApp in $SPWebApp){
            Write-Host "The Total Number of Content Database per Web Application" $WebApp.Url "is" (Get-SPContentDatabase -WebApplication $WebApp).count -ForegroundColor green
            Get-SPContentDatabase -WebApplication $WebApp | Select-Object name,@{Name="Contenet Database Size (GB)"; Expression={[math]::Round($_.disksizerequired/1024MB,2)}},@{Name="Within the Limit"; Expression={if([math]::Round($_.disksizerequired/1024MB,2) -ge 200){ "No"} else {"Yes"}}} | format-list
            Write-Host "---------------------------------"
            }

            ##################################################################################
            #How Many SharePoint Site Collections Per farm 
            Write-Host "SharePoint Site Collections Per farm" -ForegroundColor cyan
            Write-Host "The supported limit for SharePoint Site Collections per farm is 750,000 site collections."
            switch($SiteCollectionscount)
            {
            {$_ -ge 750000} {Write-Host "Total Number of SharePoint Site Collections:" $_ "|"$ex -ForegroundColor red }
            {$_ -lt 750000} {Write-Host "Total Number of SharePoint Site Collections:" $_ "|"$wa -ForegroundColor yellow}
            }
            $SiteCollections
            Write-Host "For more details, please check https://spgeeks.devoworx.com/sharepoint-2019-limitations" -ForegroundColor cyan
            Write-Host "--------------------------------------------------------------------"


            ##################################################################################
            #How Many SharePoint Site Collection Per Web Application
            Write-Host "SharePoint Site Collection Per Web Application" -ForegroundColor cyan
            foreach($WebApp in $SPWebApp){
            Write-Host "The Total Number of site collection per Web Application" $WebApp.Url "is" (Get-SPSite -WebApplication $WebApp).count -ForegroundColor green
            Get-SPSite -WebApplication $WebApp | Select-Object url,contentdatabase,@{Name="Site Collection Size (MB)";Expression={[math]::Round($_.usage.storage/(1MB),2)}},@{Name="Within the Limit"; Expression={if([math]::Round($_.usage.storage/1MB,2) -ge 100*1024){ "No"} else {"Yes"}}} | format-list
            Write-Host "---------------------------------"
            }

            ##################################################################################
            #How Many SharePoint Site Collection Per Content Database
            Write-Host "SharePoint Site Collection Per Content Database" -ForegroundColor cyan
            foreach($cDB in Get-SPContentDatabase){
            Write-Host "The Total Number of site collection per content database" $cDB.name "is" (Get-SPSite -ContentDatabase $cDB).count -ForegroundColor green
            Get-SPSite -ContentDatabase $cDB | Select-Object url,@{Name="Site Collection Size (MB)";Expression={[math]::Round($_.usage.storage/(1MB),2)}},@{Name="Within the Limit"; Expression={if([math]::Round($_.usage.storage/1MB,2) -ge 100*1024){ "No"} else {"Yes"}}} | format-list | format-list
            Write-Host "---------------------------------"
            }


            ##################################################################################
            #How Many SharePoint SubSites Per farm 
            Write-Host "SharePoint SubSites Per farm" -ForegroundColor cyan
            Write-Host "The supported limit for subsites per farm is 250,000 subsites"
            switch($Subsitescount)
            {
            {$_ -ge 250000} {Write-Host "Total Number of SharePoint SubSites:" $_ "|"$ex -ForegroundColor red }
            {$_ -lt 250000} {Write-Host "Total Number of SharePoint SubSites:" $_ "|"$wa -ForegroundColor yellow}
            }
            $Subsites
            Write-Host "For more details, please check https://spgeeks.devoworx.com/sharepoint-2019-limitations" -ForegroundColor cyan
            Write-Host "--------------------------------------------------------------------"

            ##################################################################################
            #How Many SharePoint SubSites per Site Collection
            Write-Host "SharePoint SubSites per Site Collection" -ForegroundColor cyan
            foreach($SC in Get-SPSite){
            Write-Host "The Total Number of SubSites per Site Collection" $SC.Url "is" (Get-SPWeb -Site $SC -Limit All).count -ForegroundColor green
            Get-SPWeb -Site $SC -Limit All | Select-Object url | format-list
            Write-Host "---------------------------------"
            }

            ##################################################################################
            #Farm Summary 
            Write-Host "Farm Report Summary" -ForegroundColor Green
            Write-Host "Total Number of SharePoint Web Application Per Farm:" $SPWebAppcount
            Write-Host "Total Number of SharePoint Application Pool Per Farm:" $SPAppPoolCount
            Write-Host "Total Number of SharePoint Service Application Per Farm:" $SPSrvAppcount
            Write-Host "Total Number of SharePoint Content Database Per Farm:" $ContentDBcount
            Write-Host "Total Number of SharePoint Site Collection Per Farm:" $SiteCollectionscount
            Write-Host "Total Number of SharePoint SubSites Per Farm:" $Subsitescount
            Write-Host "For more details, please check https://spgeeks.devoworx.com/sharepoint-farm-scan-report/" -ForegroundColor cyan
            Write-Host "--------------------------------------------------------------------"
		}
		Catch
		{
			Write-Host $_.Exception.Message -ForegroundColor Red
		}
  }

#Run the SharePoint Farm Scan Report
SPFramScanReport