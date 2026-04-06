Add-PSSnapin Microsoft.SharePoint.PowerShell

#	List all site collections in a farm:
Get-SPSite -Limit All | Select-Object -ExpandProperty AllWebs | Select-Object -ExpandProperty Lists | Select-Object {$_.ParentWeb.Url}, Title
   
#	Report on a certain file on a SharePoint site:
Get-SPWeb http://sharepoint/sites/enterprise | Select-Object -ExpandProperty Lists | Where-Object { $_.GetType().Name -eq "SPDocumentLibrary" -and -not $_.Hidden } | Select-Object -ExpandProperty Items | Where-Object { $_.Name -like "*FileName*" } | Select-Object Name, {$_.File.Length}, url
   
#	Report on all files created by a certain Active Directory user:
Get-SPWeb http://sharepoint/sites/enterprise | Select-Object -ExpandProperty Lists |  Where-Object { $_.GetType().Name -eq "SPDocumentLibrary" -and -not $_.Hidden } | Select-Object -ExpandProperty Items | Where-Object { $_["Created By"] -like "*UserName*" } | Select-Object Name, url, {$_["Created By"]}

#	Report on all files with a specified extension:
Get-SPWeb http://sharepoint/sites/enterprise | Select-Object -ExpandProperty Lists | Where-Object { $_.GetType().Name -eq "SPDocumentLibrary" -and -not $_.Hidden } | Select-Object -ExpandProperty Items | Where-Object { $_.Name -Like "*.rtf" } | Select-Object Name, @{Name="URL"; Expression={$_.ParentList.ParentWeb.Url + "/" + $_.Url}}

#	Report on the number of files that are hosted your sites and their total size:
Get-SPWeb http://sharepoint/sites/enterprise | Select-Object -ExpandProperty Lists | Where-Object { $_.GetType().Name -eq "SPDocumentLibrary" -and -not $_.Hidden } | Select-Object -ExpandProperty Items | Group-Object {$_.ParentList.ParentWeb.Url + "/" + $_.ParentList.Title} | Select-Object Name, count, @{Name='Total'; Expression={$_.Group | ForEach-Object -Begin {$total=0;} -Process {$total+=[int]$_.File.Length} -End {$total} }} | Format-Table -AutoSize

