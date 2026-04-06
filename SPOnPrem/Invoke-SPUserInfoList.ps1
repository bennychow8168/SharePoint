Param
(
    [Parameter(Mandatory=$False)] [Switch] $Export = $False,
    # SharePoint 2013 site URL
    [Parameter(Mandatory=$False)] [string] $SourceSiteUrl,
    [Parameter(Mandatory=$False)] [Switch] $Import = $False,
    # SharePoint 2019 site URL
    [Parameter(Mandatory=$False)] [string] $TargetSiteUrl,
    [Parameter(Mandatory=$True)] [string] $CsvPath        
)

# Load SharePoint snap-in if not loaded
if (-not (Get-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue)) {
    Add-PSSnapin Microsoft.SharePoint.PowerShell
}

If ($Export -eq $true -and $SourceSiteUrl -ne $null -and $Import -eq $false -and $TargetSiteUrl -eq $null) {
    # 1. Export User Information List from SharePoint 2013
    Write-Host "Exporting User Information List from $sourceSiteUrl ..."
    $web2013 = Get-SPWeb $sourceSiteUrl
    $userInfoList = $web2013.SiteUserInfoList

    $userDataCollection = @()
    foreach ($item in $userInfoList.Items) {
        $exportItem = New-Object PSObject
        $xml = [xml]$item.xml

        $exportItem | Add-Member -MemberType NoteProperty -Name "UserID" -Value $item.ID
        $exportItem | Add-Member -MemberType NoteProperty -Name "Title" -Value $xml.row.ows_Title
        $exportItem | Add-Member -MemberType NoteProperty -Name "Login" -Value $xml.row.ows_Name
        $exportItem | Add-Member -MemberType NoteProperty -Name "Email" -Value $xml.row.ows_EMail
        $exportItem | Add-Member -MemberType NoteProperty -Name "Department" -Value $xml.row.ows_Department
        $exportItem | Add-Member -MemberType NoteProperty -Name "JobTitle" -Value $xml.row.ows_JobTitle

        $userDataCollection += $exportItem
    }
    $userDataCollection | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding utf8
    Write-Host "Export completed to [$CsvPath]"
    $web2013.Dispose()
}

ElseIf ($Import -eq $true -and $TargetSiteUrl -ne $null -and $Export -eq $false -and $SourceSiteUrl -eq $null) {
    # 2. Connect to SharePoint 2019 site and create "Dummy" list
    Write-Host "Creating 'Dummy' list on $targetSiteUrl ..."
    $web2019 = Get-SPWeb $targetSiteUrl

    # Check if list exists, if yes, remove or rename as needed
    $listName = "Dummy"
    $existingList = $web2019.Lists.TryGetList($listName)
    if ($existingList -ne $null) {
        Write-Host "'Dummy' list already exists. Removing it..."
        $web2019.Lists.Delete($listName)
        $web2019.Update()
    }

    # Create a generic custom list
    $web2019.Lists.Add($listName, "Imported User Information", [Microsoft.SharePoint.SPListTemplateType]::GenericList) | Out-Null
    $dummyList = $web2019.Lists[$listName]

    # Add columns to the list matching exported CSV fields
    $fieldsToAdd = @(
        @{InternalName="Login"; DisplayName="Login"; Type="Text"},
        @{InternalName="Email"; DisplayName="Email"; Type="Text"},
        @{InternalName="Department"; DisplayName="Department"; Type="Text"},
        @{InternalName="JobTitle"; DisplayName="JobTitle"; Type="Text"},
        @{InternalName="DataFrom"; DisplayName="DataFrom"; Type="Text"}
    )

    foreach ($field in $fieldsToAdd) {
        if (-not $dummyList.Fields.ContainsField($field.DisplayName)) {
            $dummyList.Fields.Add($field.DisplayName, [Microsoft.SharePoint.SPFieldType]::$field.Type, $false)
        }
    }
    $dummyList.Update()

    # 3. Import CSV data into "Dummy" list and set DataFrom = "2013"
    Write-Host "Importing data into 'Dummy' list ..."
    $csvData = Import-Csv -Path $exportCsvPath

    foreach ($row in $csvData) {
        $item = $dummyList.Items.Add()
        $item["UserID"] = $row.UserID
        $item["Title"] = $row.Title
        $item["Login"] = $row.Login
        $item["Email"] = $row.Email
        $item["Department"] = $row.Department
        $item["JobTitle"] = $row.JobTitle
        $item["DataFrom"] = "2013"  # Existing items set to 2013
        $item.Update()
    }

    # 4. Set default value of "DataFrom" column to "2019" for new items
    # Note: SharePoint does not automatically update existing items with default value, but new items will have it

    $dataFromField = $dummyList.Fields.GetField("DataFrom")
    $dataFromField.DefaultValue = "2019"
    $dataFromField.Update()
    $dummyList.Update()

    Write-Host "Import and configuration completed."

    $web2019.Dispose()
}

