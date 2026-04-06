$farm = Get-SPFarm

foreach($solution in $farm.Solutions){

   $solution = $farm.Solutions[$solution.Name]

   Write-Host $solution.Name

}

You can download all Custom WSPs with a simple addition in above PowerShell script:

$farm = Get-SPFarm

$folderPath = “C:solutions”

foreach($solution in $farm.Solutions){

   $solution = $farm.Solutions[$solution.Name]

   $file = $solution.SolutionFile

   #This will save the WSP

   $file.SaveAs($folderPath + ‘’ + $solution.Name)

} 