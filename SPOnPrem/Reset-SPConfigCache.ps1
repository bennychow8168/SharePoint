# Add Snapin
# Start Loading SharePoint Snap-in
$snapin = (Get-PSSnapin -Name Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue)
IF ($null -ne $snapin) {
	write-host -f Green “SharePoint Snap-in is loaded… No Action taken”
}
ELSE {
	write-host -f Yellow “SharePoint Snap-in not found… Loading now”
	Add-PSSnapin Microsoft.SharePoint.PowerShell
	write-host -f Green “SharePoint Snap-in is now loaded”
}
# END Loading SharePoint Snapin

##################

Stop-Service SPTimerV4
$folders = Get-ChildItem C:\ProgramData\Microsoft\SharePoint\Config
foreach ($folder in $folders)
{
	$items = Get-ChildItem $folder.FullName -Recurse
	foreach ($item in $items)
	{
		if ($item.Name.ToLower() -eq “cache.ini”)
		{
			$cachefolder = $folder.FullName
		}
	}
}
$cacheIn = Get-ChildItem $cachefolder -Recurse
foreach ($cachefolderitem in $cacheIn)
{
	if ($cachefolderitem -like “*.xml”)
	{
		$cachefolderitem.Delete()
	}
}

$a = Get-Content  $cachefolder\cache.ini
$a  = 1
Set-Content $a -Path $cachefolder\cache.ini

read-host “press ENTER”
start-Service SPTimerV4