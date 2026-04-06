Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue

# Clear the screen
Clear-Host

# Stop on errors
$ErrorActionPreference = "Stop"

############################
##
## FUNCTIONS
##
############################

# Exports the permission groups and their permission levels to CSV
Function exportPermissionGroup($fileNameToUse, $doExport)
{
    Clear-Host

    # Check if the data need to be written to CVS
    if($doExport) {
        # Ask where to export the file + add the file name
        $exportPath = Select-Folder
        $exportPath +=  ("\" + $fileNameToUse + ".csv")

        # If the file exists, delete it
        deleteExisting -filePathToCheck $exportPath

        # Add the headers
        writeToCSV -textToWrite ("Group Name:;Roles:") -filePath $exportPath
    }

    # Loop through the groups
    foreach($group in $web.Groups) {               
        # Go through the roles
        $roleList = ""
        foreach($role in $group.Roles) {
            $roleList += (";" + $role.Name)
        }

        # Inform
        Write-Output ("> " + $group.Name)
        Write-Output ("  |-> " + $roleList.replace(";","`t"))
        Write-Output ""

        # Check if the data need to be written to CVS
        if($doExport) {
            # Add the group name and the roles to the file
            writeToCSV -textToWrite ($group.Name + $roleList) -filePath $exportPath
        }
    }

    # Check if the data need to be written to CVS
    Write-Output ""
    if($doExport) {
        Write-Warning ("The file has been stored at: " + $exportPath)
    }
    Write-Output ""
}

# Exports the users and their permission levels to CSV
Function exportWebUsers($fileNameToUse, $doExport)
{
    Clear-Host

    # Create a variable to hold the users
    $userList = @()

    # Check if the data need to be written to CVS
    if ($doExport) {
        # Ask where to export the file + add the file name
        $exportPath = Select-Folder
        $exportPath +=  ("\" + $fileNameToUse + ".csv")

        # If the file exists, delete it
        deleteExisting -filePathToCheck $exportPath

        # Add the headers
        writeToCSV -textToWrite ("User:;Display Name:;E-mail:;SharePoint ID:;Site Admin:;Permission Origin:") -filePath $exportPath
    }

    # Loop through the groups
    foreach ($group in $web.Groups) {               
        # Loop through the users
        foreach($user in $group.Users) {
            # Get the user's information
            $userInfo = [String]::Format("{0};{1};{2};{3};{4};SP Group", $user.UserLogin, $user.DisplayName, $user.Email, $user.ID, $user.IsSiteAdmin)

            # Check if the user is not already in the list
            if(!($userList -contains $userInfo)) {
                # Add the user to the list
                $userList += @($userInfo)
            }
        }
    }

    # Get the direct permissions on the site
    # Loop through the groups
    foreach($user in $web.Users) {
        # Get the user's information
        $userInfo = [String]::Format("{0};{1};{2};{3};{4};Direct", $user.UserLogin, $user.DisplayName, $user.Email, $user.ID, $user.IsSiteAdmin)

        # Check if the user is not already in the list
        if(!($userList -contains $userInfo)) {
            # Add the user to the list
            $userList += @($userInfo)
        }
    }

    # Go through the list of users
    foreach ($storedUser in $userList) {       
        # Inform
        Write-Output ("> " + $storedUser.replace(";","`t"))

        # Check if the data need to be written to CVS
        if ($doExport) {
        # Add the users to the file
        writeToCSV -textToWrite $storedUser -filePath $exportPath
        }
    }

    # Check if the data need to be written to CVS
    Write-Output ""
    if ($doExport) {
        Write-Warning ("The file has been stored at: " + $exportPath)
    }
    Write-Output ""
}

# Exports the groups and the users of the web
Function exportGroupAndUsers($fileNameToUse, $doExport)
{
    Clear-Host

    # Check if the data need to be written to CVS
    if ($doExport) {
        # Ask where to export the file + add the file name
        $exportPath = Select-Folder
        $exportPath += ("\" + $fileNameToUse + ".csv")

        # If the file exists, delete it
        deleteExisting -filePathToCheck $exportPath

        # Add the headers
        writeToCSV -textToWrite ("Group Name:;User:;Display Name:;E-mail:;SharePoint ID:;Site Admin:") -filePath $exportPath
    }

    # Loop through the groups
    foreach ($group in $web.Groups) {
        # Check if there are users
        if ($group.Users.Count -gt 0) {
            # Loop through the users
            foreach ($user in $group.Users) {
                # Create the text
                $textToWrite = [String]::Format("{0};{1};{2};{3};{4};{5}", $group.Name,$user.UserLogin, $user.DisplayName, $user.Email, $user.ID, $user.IsSiteAdmin)

                # Inform
                Write-Output ($textToWrite.replace(";","`t"))

                # Check if the data need to be written to CVS
                if ($doExport) {
                    writeToCSV -textToWrite $textToWrite -filePath $exportPath
                }
            }
        }
        else # No users
        {
            # Create the text
            $textToWrite = [String]::Format("{0};[Empty]", $group.Name)

            # Inform
            Write-Output ($textToWrite.replace(";","`t"))

            # Check if the data need to be written to CVS
            if ($doExport) {
                writeToCSV -textToWrite $textToWrite -filePath $exportPath
            }
        }
    }


    # Check if the data need to be written to CVS
    Write-Output ""
    if ($doExport) {
        Write-Warning ("The file has been stored at: " + $exportPath)
    }
    Write-Output ""
}

# Export single user information
Function exportSingleUser($fileNameToUse, $doExport)
{
    Clear-Host

    # Check if the data need to be written to CVS
    if ($doExport) {
        # Ask where to export the file + add the file name
        $exportPath = Select-Folder
        $exportPath +=  ("\" + $fileNameToUse + ".csv")

        # If the file exists, delete it
        deleteExisting -filePathToCheck $exportPath
    }

    # Variable to hold the user information
    $userName = $null

    while($null -eq $userName) {
        # Ask for the user's login name
        $userName = Read-Host -Prompt "User to export information from - [domain\username]"

        # Variables to hold the group and user information
        $getUserGroups = $null
        $getUserInfo = $null

        try {
            # Try to retrieve the user information
            $getUserInfo = Get-SPUser -Web $web -Identity $userName

            # Check if there is data for this user in this web
            $getUserGroups = @($web.Groups | Where-Object {$_.Users | Where-Object {$_.UserLogin -eq $userName}})

            # Check if this user has information for this site
            if ($getUserGroups.Count -eq 0) {
                # Trigger the catch block
                throw
            }
        }
        catch {
            # Show message
            #Clear-Host
            Write-Output ""
            Write-Warning "The user could not be found or is not member of any groups in this web"
            Write-Output ""

            # Reset the username
            $userName = $null
        }
    }

    # Variable to store user info + fill it
    $userData = [String]::Format("User:;{0}`r`nDisplay Name:;{1}`r`nE-mail:;{2}`r`nSharePoint ID:;{3}`r`nSite Admin:;{4}`r`n`r`nGroups:", $getUserInfo.UserLogin, $getUserInfo.Name, $getUserInfo.Email, $getUserInfo.ID, $getUserInfo.IsSiteAdmin)
    Write-Output ($userData.replace(";","`t"))

    # Check if the data need to be written to CVS
    if ($doExport) {
        writeToCSV -textToWrite $userData -filePath $exportPath
    }

    # Loop through the groups
    foreach ($group in $getUserGroups) {
        # Variable to store the data
        $groupInfo = [String]::Format(";{0}", $group.Name)
        Write-Output $groupInfo.replace(";","`t")

        # Check if the data need to be written to CVS
        if ($doExport) {
            writeToCSV -textToWrite $groupInfo -filePath $exportPath
        }
    }

    # Check if the data need to be written to CVS
    Write-Output ""
    if ($doExport) {
        Write-Warning ("The file has been stored at: " + $exportPath)
    }
    Write-Output ""
}

# Appends the text to the file in ASCII
Function writeToCSV($textToWrite, $filePath)
{
    # Write as ASCII
    ($textToWrite) | Out-File $filePath -Append -Encoding ASCII
}

# Checks if the path exists, if it does, the file gets deleted
Function deleteExisting($filePathToCheck)
{
    # Check if the path exists
    if ((Test-Path -LiteralPath $filePathToCheck) -eq $True) {
        # The file exists, so delete it
        Remove-Item -LiteralPath $filePathToCheck
    }
}

# Opens a folder browser dialog
Function Select-Folder($title='Select a folder', $path = 0) 
{  
    #   Create a new shell object
    $object = New-Object -comObject Shell.Application   

    #   Open the browser dialog
    $folder = $object.BrowseForFolder(0, $title, 0, $path)  

    #   Return the selected path
    return $folder.self.Path
}

############################
############################

# Ask the user what web to connect to
$webUrl = (Read-Host -Prompt "URL of the web to connect to").trim()

# Try to connect to the web
$web = Get-SPWeb $webUrl

# Ask the user which action to perform
Clear-Host
$selectedVal = -1
while($selectedVal -eq -1)
{
    Write-Output ""
    Write-Output "Actions:"
    Write-Output ""

    Write-Output "1 . Export permission groups on this web to CSV"
    Write-Output "2.  Export users that have permissions on this web to CSV"
    Write-Output "3.  Export permission groups including users on this web CSV"
    Write-Output "4.  Export permissions a single user on this web to CSV"
    Write-Output ""
    Write-Output "-1. Display permission groups on this web"
    Write-Output "-2. Display users that have permissions on this web"
    Write-Output "-3. Display permission groups including users"
    Write-Output "-4. Display permissions of a single user on this web"

    Write-Output ""
    $selectedVal = Read-Host -Prompt "Which action would you like to perform?"

    switch ($selectedVal) {
        ####
        # EXPORT
        ####

        # 1 was selected
        1 { exportPermissionGroup -fileNameToUse ($web.Title + "_Group_Export") -doExport $true }

        # 2 was selected
        2 { exportWebUsers -fileNameToUse ($web.Title + "_User_Export") -doExport $true }

        # 3 was selected
        3 { exportGroupAndUsers -fileNameToUse ($web.Title + "_Group_And_User_Export") -doExport $true }

        # 4 was selected
        4 { exportSingleUser -fileNameToUse ($web.Title + "_Single_User_Export") -doExport $true }

        ####
        # NO EXPORT
        ####

        # -1 was selected
        -1 { exportPermissionGroup -fileNameToUse $null -doExport $false }

        # -2 was selected
        -2 { exportWebUsers -fileNameToUse $null -doExport $false }

        # -3 was selected
        -3 { exportGroupAndUsers -fileNameToUse $null -doExport $false }

        # -4 was selected
        -4 { exportSingleUser  -fileNameToUse $null -doExport $false }

        default {
            # Invalid selection
            Clear-Host
            Write-Output ("")
            Write-Warning "The selection you made is invalid, please try again..."
            Write-Output ("")

            # Reset the selected value
            $selectedVal = -1
        }
    }
}