Add-PSSnapin Microsoft.SharePoint.PowerShell

#	Run the following script on your SharePoint server, specifying the SharePoint site URL ($SPSiteURL) and the file path for export to csv ($ExportFile):

[void][System.Reflection.Assembly]::LoadWithPartialName
("Microsoft.SharePoint")
$SPSiteUrl = "http://sharepoint/sites/ent"
$SPSite = New-Object Microsoft.SharePoint.SPSite($SPSiteUrl);
$ExportFile = "C:\root\Permissions.csv" 
"Web Title,Web URL,List Title,User or Group,Role,Inherited" | out-file $ExportFile 
foreach ($WebPath in $SPSite.AllWebs) {
  if ($WebPath.HasUniqueRoleAssignments) {
    $SPRoles = $WebPath.RoleAssignments;
    foreach ($SPRole in $SPRoles) {
      foreach ($SPRoleDefinition in $SPRole.RoleDefinitionBindings) {
        $WebPath.Title + "," + $WebPath.Url + "," + "N/A" + "," + $SPRole.Member.Name + "," + $SPRoleDefinition.Name + "," + $WebPath.HasUniqueRoleAssignments | out-file $ExportFile -append
      }
    }
  }
  foreach ($List in $WebPath.Lists) {
    if ($List.HasUniqueRoleAssignments) {
      $SPRoles = $List.RoleAssignments;
      foreach ($SPRole in $SPRoles) {
        foreach ($SPRoleDefinition in $SPRole.RoleDefinitionBindings) {
          $WebPath.Title + "," + $WebPath.Url + "," + $List.Title + "," + $SPRole.Member.Name + "," + $SPRoleDefinition.Name | out-file $ExportFile -append
        }
      }
    }
  }
}
$SPSite.Dispose();
