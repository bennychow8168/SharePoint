<#
.SYNOPSIS
    Sets (adds or updates) a property bag value on a SharePoint Subscription Edition
    (on-premises) site (SPWeb) using the SharePoint Server Management Shell / SSOM.

.DESCRIPTION
    Uses Get-SPWeb and the .AllProperties hashtable to add/update a key-value pair
    in the web's property bag, then persists the change with .Update() and
    .Properties.Update(). Must be run on a SharePoint server (or a machine with the
    SharePoint Management Shell / snap-in loaded) under an account with at least
    Full Control on the target web.

.PARAMETER SiteUrl
    Full URL of the site (SPWeb) to update, e.g. https://intranet.contoso.com/sites/hr

.PARAMETER Key
    Property bag key name.

.PARAMETER Value
    Property bag value to set.

.PARAMETER Indexed
    Optional switch. If specified, marks the property as indexed so it can be used
    in search-driven queries (adds the key to the vti_indexedpropertykeys list).

.EXAMPLE
    .\Set-SPPropertyBagValue.ps1 -SiteUrl "https://intranet.contoso.com/sites/hr" -Key "DeptCode" -Value "HR-01"

.EXAMPLE
    .\Set-SPPropertyBagValue.ps1 -SiteUrl "https://intranet.contoso.com/sites/hr" -Key "DeptCode" -Value "HR-01" -Indexed

.NOTES
    Target: SharePoint Subscription Edition (on-premises), also compatible with 2019/2016.
    Run from an elevated PowerShell session on a SharePoint server, or a Windows
    machine with the SharePoint Management Shell / Microsoft.SharePoint.PowerShell
    snap-in registered.
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [string]$SiteUrl,
    [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [string]$Key,
    [Parameter(Mandatory = $true)] [AllowEmptyString()] [string]$Value,
    [switch]$Indexed
)

# --- Ensure the SharePoint snap-in is loaded ---
if (-not (Get-PSSnapin -Name Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue)) {
    try {
        Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction Stop
    }
    catch {
        Write-Error "Unable to load the Microsoft.SharePoint.PowerShell snap-in. Run this script from a SharePoint server or a machine with the SharePoint Management Shell installed. $($_.Exception.Message)"
        return
    }
}

$web = $null
try {
    $web = Get-SPWeb -Identity $SiteUrl -ErrorAction Stop
}
catch {
    Write-Error "Could not open site '$SiteUrl'. Verify the URL and your permissions. $($_.Exception.Message)"
    return
}

try {
    $existing = $web.AllProperties[$Key]

    if ($PSCmdlet.ShouldProcess("$SiteUrl [$Key]", "Set property bag value to '$Value'")) {

        if ($null -eq $existing) {
            Write-Host "Adding new property bag key '$Key' on '$SiteUrl'..." -ForegroundColor Cyan
        }
        else {
            Write-Host "Updating existing property bag key '$Key' (was: '$existing') on '$SiteUrl'..." -ForegroundColor Cyan
        }

        $web.AllProperties[$Key] = $Value

        if ($Indexed) {
            $indexedKeysProp = "vti_indexedpropertykeys"
            $indexedKeysRaw = $web.AllProperties[$indexedKeysProp]

            # vti_indexedpropertykeys is stored as a base64-encoded, newline-delimited list of key names
            $keyList = New-Object System.Collections.Generic.List[string]
            if (-not [string]::IsNullOrEmpty($indexedKeysRaw)) {
                try {
                    $decoded = [System.Text.Encoding]::Unicode.GetString([System.Convert]::FromBase64String($indexedKeysRaw))
                    $keyList.AddRange(($decoded -split "`n" | Where-Object { $_ -ne "" }))
                }
                catch {
                    Write-Warning "Could not parse existing vti_indexedpropertykeys value; it will be recreated."
                }
            }

            if (-not $keyList.Contains($Key)) {
                $keyList.Add($Key)
                $newRaw = ($keyList -join "`n") + "`n"
                $newEncoded = [System.Convert]::ToBase64String([System.Text.Encoding]::Unicode.GetBytes($newRaw))
                $web.AllProperties[$indexedKeysProp] = $newEncoded
                Write-Host "Marked '$Key' as an indexed property." -ForegroundColor Cyan
            }
        }

        $web.Update()
        $web.Properties.Update()

        Write-Host "Property bag value set successfully." -ForegroundColor Green
        Write-Host ("  Site : {0}" -f $SiteUrl)
        Write-Host ("  Key  : {0}" -f $Key)
        Write-Host ("  Value: {0}" -f $Value)
    }
}
catch {
    Write-Error "Failed to set property bag value: $($_.Exception.Message)"
}
finally {
    if ($web) { $web.Dispose() }
}
