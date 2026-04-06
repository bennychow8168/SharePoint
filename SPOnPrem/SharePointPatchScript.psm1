<#
    Updated for SharePoint Server 2016 and 2019 by Trevor Seward (https://thesharepointfarm.com). Original
    script provided by Russ Maxwell, available for SharePoint 2013 from
    https://blog.russmax.com/why-sharepoint-2013-cumulative-update-takes-5-hours-to-install/.

    This script supports both SharePoint 2013, SharePoint Server 2016, and SharePoint Server 2019. SharePoint Server 2016/2019 supports both the
    sts*.exe and wssloc*.exe in the same directory.
 
    License: MIT (https://github.com/Nauplius/SharePoint-Patch-Script/blob/master/LICENSE)
#>

Add-PSSnapin Microsoft.SharePoint.PowerShell -EA Continue

<#
    .SYNOPSIS
        Get-SPPatchInfo retrieves information about a knowledge base article or build number.
    .DESCRIPTION
        Get-SPPatch retreives the patch metadata from https://sharepointupdates.com/patches from the supplied knowledge base article number
        or SharePoint patch build number. Additional information can be found at https://github.com/Nauplius.
    .PARAMETER Build
        The build number to retrieve metadata for.
    .PARAMETER KnowledgeBaseArticle
        The knowledge base article to retrieve metadata for.
    .EXAMPLE
        Get-SPPatchInfo -Build 16.0.10354.20001

        Retrieves the patch information for build number 16.0.10354.20001.
    .EXAMPLE
        Get-SPPatchInfo -KnowledgeBaseArticle 4484224

        Retrieves the patch information for the knowledge base article 4484224.
    .NOTES
        Author: Trevor Seward
        Date: 01/16/2020
    .LINK
        https://thesharepointfarm.com
    .LINK
        https://github.com/Nauplius
    .LINK
        https://sharepointupdates.com
#>

Function Get-SPPatchInfo {
    [CmdletBinding()]
    param (
        [string]
        [Parameter(Mandatory = $true, ParameterSetName = "Build")]
        $Build,
        [string]
        [Parameter(Mandatory = $true, ParameterSetName = "KnowledgeBaseArticle")]
        $KnowledgeBaseArticle
    )

    $service = 'https://sharepointupdates.com/api/Articles/'

    switch ($PSCmdlet.ParameterSetName) {
        "Build" {
            if ($Build -notmatch "[0-9]{2}.[0-9]{1}.[0-9]{4,5}.[0-9]{4,5}") {
                throw 'Invalid build number.'
            }

            try {
                Invoke-WebRequest "$($service)/$Build" | ConvertFrom-Json
            }
            catch {
                throw 'Unable to find build.'
            }
        }
        "KnowledgeBaseArticle" {
            $KnowledgeBaseArticle = [regex]::Replace($KnowledgeBaseArticle, "[^0-9]", "")

            try {
                Invoke-WebRequest "$($service)/$KnowledgeBaseArticle" | ConvertFrom-Json
            }
            catch {
                throw 'Unable to find article.'
            }  
        }
    }
}

<#
    .SYNOPSIS
        Get-SPPatch
    .DESCRIPTION
        Get-SPPatch retreives the patch metadata from https://sharepointupdates.com/patches and then downloads the patch file(s) to the specified directory. Additional information
        can be found at https://github.com/Nauplius.
    .PARAMETER Path
        The folder where the patch file(s) will be downloaded to. Must be an existing folder to which the account performing the download has write permissions.
    .PARAMETER KBs
        The knowledge base article(s) to look up the patch file(s), separated by a comma. Only specify the KB article number.
    .PARAMETER ShowDownloadProgress
        Set to false by default, displaying download progress may significantly reduce download performance.
    .EXAMPLE
        Get-SPPatch -Path 'C:\Temp' -KBs 4484176,4484177

        Downloads the knowledge base articles to C:\Temp.
    .EXAMPLE
        Get-SPPatch -Path 'C:\Temp' -KBs 4484176,4484177 -ShowDownloadProgress $true

        Downloads the knowledge base articles to C:\Temp and displays the download progress.
    .NOTES
        When downloading SharePoint Server 2013 or Project Server 2013 Cumulative Updates, do not specify more than one Cumulative Update a time.
        Cumulative Updates for these products have two cab files, 'ubersrv_1.cab' and 'ubersrv_2.cab' which are not unqiuely named
        between different knowledge base articles.

        Files will be automatically overwritten when downloading a file of the same name.

        Author: Trevor Seward
        Date: 01/16/2020
    .LINK
        https://thesharepointfarm.com
    .LINK
        https://github.com/Nauplius
    .LINK
        https://sharepointupdates.com
#>

Function Get-SPPatch {
    [CmdletBinding()]
    param (
        [string]
        [Parameter(Mandatory = $true)]
        $Path,
        [string[]]
        [Parameter(Mandatory = $true)]
        $KBs,
        [bool]
        [Parameter(Mandatory = $false)]
        $ShowDownloadProgress = $false
    )

    $service = 'https://sharepointupdates.com/api/Articles/'

    if ($ShowDownloadProgress) {
        $ProgressPreference = 'Continue'
        Write-Host -ForegroundColor Yellow 'Download progress will be displayed which may significantly reduce download performance.'
    }
    else {
        $ProgressPreference = 'SilentlyContinue'
        Write-Host -ForegroundColor Yellow 'Download progress will not be displayed to improve download performance.'
    }   

    if (-not (Test-Path -Path $Path)) {
        throw "Path $($Path) does not exist."
    }
    
    foreach ($KB in $KBs) {
        $KB = [regex]::Replace($KB, "[^0-9]", "")

        try {
            $json = Invoke-WebRequest "$($service)/$KB" | ConvertFrom-Json
            if ($json.PatchUrl1.Length -ne 0) {
                $file = Split-Path $json.PatchUrl1 -Leaf
                Invoke-WebRequest $json.PatchUrl1 -OutFile "$($Path)\$($file)"
                Unblock-File -Path "$($Path)\$($file)" -Confirm:$false
            }
            else {
                throw 'Patch file not found.'
            }
            
            if ($json.PatchUrl2.Length -ne 0) {
                $file = Split-Path $json.PatchUrl2 -Leaf
                Invoke-WebRequest $json.PatchUrl2 -OutFile "$($Path)\$($file)"
                Unblock-File -Path "$($Path)\$($file)" -Confirm:$false
            }

            if ($json.PatchUrl3.Length -ne 0) {
                $file = Split-Path $json.PatchUrl3 -Leaf
                Invoke-WebRequest $json.PatchUrl3 -OutFile "$($Path)\$($file)"
                Unblock-File -Path "$($Path)\$($file)" -Confirm:$false
            }           
        }
        catch {
            throw 'Unable to find the specified Knowledge Base article.'
        }
    }
}

<#
    .SYNOPSIS
        Install-SPPatch
    .DESCRIPTION
        Install-SPPatch reduces the amount of time it takes to install SharePoint patches. This cmdlet supports SharePoint 2013 and above. Additional information
        can be found at https://github.com/Nauplius.
    .PARAMETER Path
        The folder where the patch file(s) reside.
    .PARAMETER Pause
        Pauses the Search Service Application(s) prior to stopping the SharePoint Search Services.
    .PARAMETER Stop
        Stop the SharePoint Search Services without pausing the Search Service Application(s).
    .PARAMETER SilentInstall
        Silently installs the patches without user input. Not specifying this parameter will cause each patch to prompt to install.
    .PARAMETER KeepSearchPaused
        Keeps the Search Service Application(s) in a paused state after the installation of the patch has completed. Useful for when applying the patch to multiple
        servers in the farm. Default to false.
    .PARAMETER OnlySTS
        Only apply the STS (non-language dependent) patch. This switch may be used when only an STS patch is available.
    .EXAMPLE
        Install-SPPatch -Path C:\Updates -Pause -SilentInstall

        Install the available patches in C:\Updates, pauses the Search Service Application(s) on the farm, and performs a silent installation.
    .EXAMPLE
        Install-SPPatch -Path C:\Updates -Pause -KeepSearchPaused:$true -SilentInstall

        Install the available patches in C:\Updates, pauses the Search Service Application(s) on the farm,
        does not resume the Search Service Application(s) after the installation is complete, and performs a silent installation.
    .NOTES
        Author: Trevor Seward
        Date: 01/16/2020
    .LINK
        https://thesharepointfarm.com
    .LINK
        https://github.com/Nauplius
    .LINK
        https://sharepointupdates.com
#>

Function Install-SPPatch {
    param
    (
        [string]
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        $Path,
        [switch]
        [Parameter(Mandatory = $true, ParameterSetName = "PauseSearch")]
        $Pause,
        [switch]
        [Parameter(Mandatory = $true, ParameterSetName = "StopSearch")]
        $Stop,
        [switch]
        [Parameter(Mandatory = $false, ParameterSetName = "PauseSearch")]
        $KeepSearchPaused = $false,
        [switch]
        [Parameter(Mandatory = $false)]
        $SilentInstall,
        [switch]
        [Parameter(Mandatory = $false)]
        $OnlySTS
    )

    $version = (Get-SPFarm).BuildVersion
    $majorVersion = $version.Major
    $startTime = Get-Date
    $exitRebootCodes = @(3010, 17022)
    $searchSvcRunning = $false
    
    Write-Host -ForegroundColor Green "Current build: $version"

    ###########################
    ##Ensure Patch is Present##
    ###########################

    if ($majorVersion -eq '16') {
        $sts = Get-ChildItem -LiteralPath $Path  -Filter *.exe | ? { $_.Name -match 'sts([A-Za-z0-9\-]+).exe' }
        $wssloc = Get-ChildItem -LiteralPath $Path  -Filter *.exe | ? { $_.Name -match 'wssloc([A-Za-z0-9\-]+).exe' }
        
        if ($OnlySTS) {
            if ($sts -eq $null) {
                Write-Host 'Missing the sts patch. Please make sure the sts patch present in the specified directory.' -ForegroundColor Red
                return            
            }
        }
        else {
            if ($sts -eq $null -and $wssloc -eq $null) {
                Write-Host 'Missing the sts and wssloc patch. Please make sure both patches are present in the specified directory.' -ForegroundColor Red
                return
            }

            if ($sts -eq $null -or $wssloc -eq $null) {
                Write-Host '[Warning] Either the sts and wssloc patch is not available. Please make sure both patches are present in the same directory or safely ignore if only single patch is available.' -ForegroundColor Yellow
                return
            }
        }

        if ($OnlySTS) {
            $patchfiles = $sts
            Write-Host -for Yellow "Installing $sts"
        }
        else {
            $patchfiles = $sts, $wssloc
            Write-Host -for Yellow "Installing $sts and $wssloc"
        }
    }
    elseif ($majorVersion -eq '15') {
        $patchfiles = Get-ChildItem -LiteralPath $Path  -Filter *.exe | ? { $_.Name -match '([A-Za-z0-9\-]+)2013-kb([A-Za-z0-9\-]+)glb.exe' }
        
        if ($patchfiles -eq $null) { 
            Write-Host 'Unable to retrieve the file(s). Exiting Script' -ForegroundColor Red 
            return 
        }

        Write-Host -ForegroundColor Yellow "Installing $patchfiles"
    }
    elseif ($majorVersion -lt '15') {
        throw 'This cmdlet supports SharePoint 2013 and above.'
    }

    ########################
    ##Stop Search Services##
    ########################
    ##Checking Search services##

    $oSearchSvc = Get-Service "OSearch$majorVersion" 
    $sPSearchHCSvc = Get-Service "SPSearchHostController"

    if (($oSearchSvc.status -eq 'Running') -or ($sPSearchHCSvc.status -eq 'Running')) { 
        $searchSvcRunning = $true
        if ($Pause) { 
            $ssas = Get-SPEnterpriseSearchServiceApplication

            foreach ($ssa in $ssas) {
                Write-Host -ForegroundColor Yellow "Pausing the Search Service Application: $($ssa.DisplayName)"
                Write-Host  -ForegroundColor Yellow  ' This could take a few minutes...'
                Suspend-SPEnterpriseSearchServiceApplication -Identity $ssa | Out-Null
            }
        }
        elseif ($Stop) { 
            Write-Host -ForegroundColor Cyan ' Continuing without pausing the Search Service Application'
        }
    }

    #We don't need to stop SharePoint Services for 2016 and above
    if ($majorVersion -lt '16') {
        Write-Host -ForegroundColor Yellow 'Stopping Search Services if they are running'

        if ($oSearchSvc.status -eq 'Running') { 
            Set-Service -Name "OSearch$majorVersion" -StartupType Disabled 
            Stop-Service "OSearch$majorVersion" -WA 0
        }

        if ($sPSearchHCSvc.status -eq 'Running') { 
            Set-Service 'SPSearchHostController' -StartupType Disabled 
            Stop-Service 'SPSearchHostController' -WA 0
        }

        Write-Host -ForegroundColor Green 'Search Services are Stopped'
        Write-Host

        #######################
        ##Stop Other Services##
        #######################
        Set-Service -Name 'IISADMIN' -StartupType Disabled 
        Set-Service -Name 'SPTimerV4' -StartupType Disabled

        Write-Host -ForegroundColor Green 'Gracefully stopping IIS...'
        Write-Host 
        iisreset -stop -noforce 
        Write-Host -ForegroundColor Yellow 'Stopping SPTimerV4'
        Write-Host

        $sPTimer = Get-Service 'SPTimerV4' 
        if ($sPTimer.Status -eq 'Running') {
            Stop-Service 'SPTimerV4'
        }

        Write-Host -ForegroundColor Green 'Services are Stopped'
        Write-Host 
        Write-Host
    }

    ##################
    ##Start patching##
    ##################
    Write-Host -ForegroundColor Yellow 'Working on it... Please keep this PowerShell window open...'
    Write-Host

    $patchStartTime = Get-Date

    foreach ($patchfile in $patchfiles) {
        $filename = $patchfile.Fullname
        #unblock the file, to get rid of the prompts
        Unblock-File -Path $filename -Confirm:$false

        if ($SilentInstall) {
            $process = Start-Process $filename -ArgumentList '/passive /quiet' -PassThru -Wait
        }
        else {
            $process = Start-Process $filename -ArgumentList '/norestart' -PassThru -Wait
        }

        if ($exitRebootCodes.Contains($process.ExitCode)) {
            $reboot = $true
        }

        Write-Host -ForegroundColor Yellow "Patch $patchfile installed with Exit Code $($process.ExitCode)"
    }

    $patchEndTime = Get-Date

    Write-Host 
    Write-Host -ForegroundColor Yellow ('Patch installation completed in {0:g}' -f ($patchEndTime - $patchStartTime))
    Write-Host

    if ($majorVersion -lt '16') {
        ##################
        ##Start Services##
        ##################
        Write-Host -ForegroundColor Yellow 'Starting Services'
        Set-Service -Name 'SPTimerV4' -StartupType Automatic 
        Set-Service -Name 'IISADMIN' -StartupType Automatic

        Start-Service 'SPTimerV4'
        Start-Service 'IISAdmin'

        ###Ensuring Search Services were stopped by script before Starting"
        if ($searchSvcRunning = $true) { 
            Set-Service -Name "OSearch$majorVersion" -StartupType Manual 
            Start-Service "OSearch$majorVersion" -WA 0
            Set-Service 'SPSearchHostController' -StartupType Automatic 
            Start-Service 'SPSearchHostController' -WA 0
        }
    }

    ###Resuming Search Service Application if paused###
    if ($Pause -and $KeepSearchPaused -eq $false) { 
        $ssas = Get-SPEnterpriseSearchServiceApplication

        foreach ($ssa in $ssas) {
            Write-Host -ForegroundColor Yellow "Resuming the Search Service Application: $($ssa.DisplayName)"
            Write-Host -ForegroundColor Yellow ' This could take a few minutes...'
            Resume-SPEnterpriseSearchServiceApplication -Identity $ssa | Out-Null
        }
    }
    elseif ($pause -and $KeepSearchPaused -eq $true) {
        Write-Host -ForegroundColor Yellow 'Not resuming the Search Service Application(s)'
    }

    ###Resuming IIS###
    iisreset -start

    $endTime = Get-Date
    Write-Host -ForegroundColor Green 'Services are Started'
    Write-Host 
    Write-Host 
    Write-Host -ForegroundColor Yellow ('Script completed in {0:g}' -f ($endTime - $startTime))
    Write-Host -ForegroundColor Yellow 'Started:'  $startTime 
    Write-Host -ForegroundColor Yellow 'Finished:'  $endTime 

    if ($reboot) {
        Write-Host -ForegroundColor Yellow 'A reboot is required'
    }
}

<#
    .SYNOPSIS
        Invoke-SPConfigWizard
    .DESCRIPTION
        Invoke-SPConfigWizard runs psconfig.exe on the local server with the necessary parameters to update the local server.
        This must be run on all servers in the farm after all servers in the farm have been patched. Additional information
        can be found at https://github.com/Nauplius.
    .EXAMPLE
        Invoke-SPConfigWizard

        Runs the Config Wizard with all the necessary switches to update the local server.
    .NOTES
        Author: Trevor Seward
        Date: 01/16/2020
    .LINK
        https://thesharepointfarm.com
    .LINK
        https://github.com/Nauplius
    .LINK
        https://sharepointupdates.com
#>
Function Invoke-SPConfigWizard {
    PSConfig.exe -cmd upgrade -inplace b2b -wait -cmd applicationcontent -install -cmd installfeatures -cmd secureresources -cmd services -install
}