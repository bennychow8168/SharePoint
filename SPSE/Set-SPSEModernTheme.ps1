#Requires -Version 5.1
<#
.SYNOPSIS
    Applies a SharePoint Online-style modern theme to SharePoint Subscription Edition (on-premises).

.DESCRIPTION
    Provisions a Fluent UI colour palette to one or more SPSE site collections via the
    /_api/ThemeManager/ApplyTheme REST endpoint.

    Theme source priority (highest → lowest):
      1. -SpoThemeJson  – raw SPO JSON string or path to a .json file
      2. -ThemePreset   – one of the five built-in presets
      3. Default        – SharePoint Blue

    SPO JSON format accepted (both the outer wrapper and the bare palette object):
      { "name": "...", "isInverted": false, "palette": { "themePrimary": "#0078d4", ... } }
      { "themePrimary": "#0078d4", ... }   <- bare palette object also works

.PARAMETER SiteUrl
    URL of the target site collection.  Accepts pipeline input or a comma-separated list.

.PARAMETER ThemeName
    Friendly name stored in SPSE's theme gallery.  Defaults to "ContosoModern".
    When -SpoThemeJson carries a "name" field that value is used unless you override it here.

.PARAMETER SpoThemeJson
    SharePoint Online theme JSON.  Pass either:
      - A JSON string  : -SpoThemeJson '{ "palette": { "themePrimary": "#c94f24", ... } }'
      - A file path    : -SpoThemeJson "C:\themes\brand.json"
    All hex colour values (#RRGGBB / #RGB) are auto-converted to the rgb() strings SPSE expects.
    Unknown keys in the JSON are passed through so future SPO palette slots are not lost.

.PARAMETER ThemePreset
    Built-in colour preset used when -SpoThemeJson is NOT supplied.
      Default  - SharePoint Blue (#0078d4)
      Teal     - Communication-site teal (#03787c)
      Purple   - SharePoint purple (#5c2d91)
      Green    - SharePoint green (#107c10)
      DarkGray - Dark-mode grey (blue on dark background)
      Custom   - Provide -PrimaryColor / -BodyTextColor

.PARAMETER PrimaryColor
    (Custom preset only) Hex colour for themePrimary, e.g. "#c94f24".

.PARAMETER BodyTextColor
    (Custom preset only) Hex colour for neutralPrimary / bodyText, e.g. "#201f1e".

.PARAMETER Credential
    PSCredential for authentication.  Omit to use the current Windows identity (Kerberos).

.PARAMETER ApplyToSubWebs
    Recursively apply the theme to all sub-sites of each target site.

.PARAMETER WhatIf
    Dry-run: shows what would happen without making any changes.

.EXAMPLE
    # Feed a .json file exported from SPO / M365 theme designer
    .\Set-SPSEModernTheme.ps1 -SiteUrl "https://sp.contoso.com/sites/hr" `
        -SpoThemeJson "C:\themes\brand.json"

.EXAMPLE
    # Feed a raw JSON string inline
    $json = Get-Content "C:\themes\brand.json" -Raw
    .\Set-SPSEModernTheme.ps1 -SiteUrl "https://sp.contoso.com/sites/hr" -SpoThemeJson $json

.EXAMPLE
    # Pipe multiple sites, apply the same JSON theme, recurse sub-webs
    "https://sp/sites/hr","https://sp/sites/it" |
        .\Set-SPSEModernTheme.ps1 -SpoThemeJson "C:\themes\brand.json" -ApplyToSubWebs

.EXAMPLE
    # Fall back to a built-in preset (no JSON supplied)
    .\Set-SPSEModernTheme.ps1 -SiteUrl "https://sp.contoso.com/sites/it" -ThemePreset Teal

.EXAMPLE
    # Dry-run to validate without touching SharePoint
    .\Set-SPSEModernTheme.ps1 -SiteUrl "https://sp.contoso.com/sites/hr" `
        -SpoThemeJson "C:\themes\brand.json" -WhatIf

.NOTES
    Prerequisites
    - SharePoint Server Subscription Edition (Feature Pack 1+ recommended)
    - PowerShell 5.1+ (PowerShell 7 works via Windows Compatibility layer)
    - Executing account must be a Site Collection Administrator on each target site
    - No extra DLLs required: pure REST

    SPO JSON format
    Export from https://aka.ms/themedesigner or retrieve via:
      Invoke-RestMethod "https://<tenant>.sharepoint.com/_api/ThemeManager/GetTenantThemingOptions"
    Both the full wrapper { name, isInverted, palette:{} } and a bare palette object are accepted.
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param (
    [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
    [ValidateNotNullOrEmpty()]
    [string[]] $SiteUrl,

    [string] $ThemeName = "ContosoModern",

    # SPO JSON string or path to a .json file
    [string] $SpoThemeJson,

    [ValidateSet("Default","Teal","Purple","Green","DarkGray","Custom")]
    [string] $ThemePreset = "Default",

    [string] $PrimaryColor  = "#0078d4",
    [string] $BodyTextColor = "#323130",

    [System.Management.Automation.PSCredential] $Credential,

    [switch] $ApplyToSubWebs
)

begin {
    Set-StrictMode -Version Latest
    $ErrorActionPreference = "Stop"

    # ── Helper: is a string a valid hex colour? ──────────────────────────────
    function Test-HexColor ([string]$value) {
        return $value -match '^#([0-9A-Fa-f]{3}|[0-9A-Fa-f]{6})$'
    }

    # ── Helper: expand shorthand #RGB to #RRGGBB ─────────────────────────────
    function Expand-HexColor ([string]$hex) {
        $hex = $hex.TrimStart('#')
        if ($hex.Length -eq 3) {
            $hex = "$($hex[0])$($hex[0])$($hex[1])$($hex[1])$($hex[2])$($hex[2])"
        }
        return "#$hex"
    }

    # ── Helper: convert #RRGGBB to "rgb(R, G, B)" ────────────────────────────
    function ConvertTo-RgbString ([string]$hex) {
        $hex = (Expand-HexColor $hex).TrimStart('#')
        $r   = [Convert]::ToInt32($hex.Substring(0,2), 16)
        $g   = [Convert]::ToInt32($hex.Substring(2,2), 16)
        $b   = [Convert]::ToInt32($hex.Substring(4,2), 16)
        return "rgb($r, $g, $b)"
    }

    # ── Helper: darken / lighten a hex colour by a multiplier ────────────────
    function Adjust-Color ([string]$hex, [double]$ratio) {
        $hex = (Expand-HexColor $hex).TrimStart('#')
        $r   = [Math]::Min(255, [Math]::Max(0, [int]([Convert]::ToInt32($hex.Substring(0,2),16) * $ratio)))
        $g   = [Math]::Min(255, [Math]::Max(0, [int]([Convert]::ToInt32($hex.Substring(2,2),16) * $ratio)))
        $b   = [Math]::Min(255, [Math]::Max(0, [int]([Convert]::ToInt32($hex.Substring(4,2),16) * $ratio)))
        return ('#{0:X2}{1:X2}{2:X2}' -f $r,$g,$b)
    }

    # ════════════════════════════════════════════════════════════════════════
    # PATH A: Parse SPO JSON into an SPSE-ready rgb() palette
    # ════════════════════════════════════════════════════════════════════════
    function ConvertFrom-SpoThemeJson ([string]$jsonInput) {

        # Accept a file path OR a raw JSON string
        if (Test-Path $jsonInput -ErrorAction SilentlyContinue) {
            Write-Verbose "Loading SPO theme from file: $jsonInput"
            $jsonInput = Get-Content $jsonInput -Raw -Encoding UTF8
        }

        try {
            $parsed = $jsonInput | ConvertFrom-Json
        } catch {
            throw "Could not parse the supplied value as JSON. " +
                  "Ensure -SpoThemeJson is a valid JSON string or a path to a .json file.`nParser error: $_"
        }

        # Support both wrapper format  { name, isInverted, palette:{} }
        # and bare palette object      { themePrimary: "#...", ... }
        if ($null -ne $parsed.palette) {
            $paletteObj = $parsed.palette
            # Use the JSON's embedded name unless the caller explicitly set ThemeName
            if ($parsed.name -and $script:ThemeName -eq "ContosoModern") {
                $script:ThemeName = $parsed.name
            }
        } else {
            $paletteObj = $parsed
        }

        # Walk every property: convert hex colours to rgb(), leave others alone
        $result = [ordered]@{}
        foreach ($prop in $paletteObj.PSObject.Properties) {
            $val = $prop.Value
            if ($val -is [string] -and (Test-HexColor $val)) {
                $result[$prop.Name] = ConvertTo-RgbString $val
            } else {
                $result[$prop.Name] = $val
            }
        }

        if (-not $result.ContainsKey("themePrimary")) {
            throw "The supplied JSON does not contain a 'themePrimary' slot. " +
                  "Please verify it is a valid SPO theme palette."
        }

        return $result
    }

    # ════════════════════════════════════════════════════════════════════════
    # PATH B: Build palette from a built-in preset
    # ════════════════════════════════════════════════════════════════════════
    $presets = @{
        Default  = @{ primary = "#0078d4"; text = "#323130"; bg = "#ffffff" }
        Teal     = @{ primary = "#03787c"; text = "#323130"; bg = "#ffffff" }
        Purple   = @{ primary = "#5c2d91"; text = "#323130"; bg = "#ffffff" }
        Green    = @{ primary = "#107c10"; text = "#323130"; bg = "#ffffff" }
        DarkGray = @{ primary = "#0078d4"; text = "#f3f2f1"; bg = "#201f1e" }
        Custom   = @{ primary = $PrimaryColor; text = $BodyTextColor; bg = "#ffffff" }
    }

    function New-PresetPalette ([hashtable]$p) {
        $pri   = $p.primary
        $text  = $p.text
        $bg    = $p.bg
        $shade = Adjust-Color $pri 0.75

        $slots = [ordered]@{
            themeDarker               = Adjust-Color $pri 0.55
            themeDark                 = $shade
            themeDarkAlt              = Adjust-Color $pri 0.85
            themePrimary              = $pri
            themeSecondary            = Adjust-Color $pri 1.10
            themeTertiary             = Adjust-Color $pri 1.40
            themeLight                = Adjust-Color $pri 1.60
            themeLighter              = Adjust-Color $pri 1.75
            themeLighterAlt           = Adjust-Color $pri 1.90
            black                     = "#000000"
            neutralDark               = "#201f1e"
            neutralPrimary            = $text
            neutralPrimaryAlt         = "#3b3a39"
            neutralSecondary          = "#605e5c"
            neutralSecondaryAlt       = "#8a8886"
            neutralTertiary           = "#a19f9d"
            neutralTertiaryAlt        = "#c8c6c4"
            neutralQuaternary         = "#d2d0ce"
            neutralQuaternaryAlt      = "#e1dfdd"
            neutralLight              = "#edebe9"
            neutralLighter            = "#f3f2f1"
            neutralLighterAlt         = "#faf9f8"
            white                     = $bg
            accent                    = $pri
            bodyBackground            = $bg
            bodyText                  = $text
            disabledBackground        = "#f3f2f1"
            errorText                 = "#a4262c"
            focusBorder               = $pri
            inputBorder               = "#8a8886"
            inputBorderHovered        = "#323130"
            inputFocusBorderAlt       = $pri
            link                      = $pri
            linkHovered               = $shade
            menuBackground            = $bg
            menuDivider               = "#edebe9"
            menuHeader                = $pri
            menuIcon                  = $pri
            menuItemBackgroundHovered = "#f3f2f1"
            menuItemText              = $text
            primaryButtonBackground   = $pri
            primaryButtonBackgroundHovered = $shade
            primaryButtonText         = "#ffffff"
        }

        $rgb = [ordered]@{}
        foreach ($key in $slots.Keys) { $rgb[$key] = ConvertTo-RgbString $slots[$key] }
        return $rgb
    }

    # ════════════════════════════════════════════════════════════════════════
    # Resolve which palette to use
    # ════════════════════════════════════════════════════════════════════════
    $usingJson = $PSBoundParameters.ContainsKey("SpoThemeJson") -and $SpoThemeJson

    if ($usingJson) {
        Write-Verbose "Theme source: SPO JSON"
        $themeData = ConvertFrom-SpoThemeJson -jsonInput $SpoThemeJson
    } else {
        Write-Verbose "Theme source: Built-in preset '$ThemePreset'"
        $themeData = New-PresetPalette $presets[$ThemePreset]
    }

    # ── Apply theme to one web via REST ──────────────────────────────────────
    function Invoke-ApplyTheme {
        param (
            [string]                         $WebUrl,
            [System.Collections.IDictionary] $Palette,
            [System.Net.ICredentials]        $Creds
        )

        $useDefault = ($null -eq $Creds)
        $authParams  = if ($useDefault) { @{ UseDefaultCredentials = $true } }
                       else             { @{ Credential = $Creds } }

        # Form digest required for all SharePoint REST POSTs
        $digestResp = Invoke-RestMethod "$WebUrl/_api/contextinfo" -Method Post `
                        -Headers @{ Accept = "application/json;odata=nometadata" } `
                        @authParams

        $headers = @{
            "Accept"          = "application/json;odata=nometadata"
            "Content-Type"    = "application/json;odata=nometadata"
            "X-RequestDigest" = $digestResp.FormDigestValue
        }

        # SPSE expects themeJson as a serialised string nested inside the outer body JSON
        $body = [ordered]@{
            name         = $ThemeName
            themeJson    = ($Palette | ConvertTo-Json -Depth 10 -Compress)
            updateNavBar = $true
        } | ConvertTo-Json -Depth 10

        if ($PSCmdlet.ShouldProcess($WebUrl, "Apply modern theme '$ThemeName'")) {
            return Invoke-RestMethod "$WebUrl/_api/ThemeManager/ApplyTheme" `
                        -Method Post -Body $body -Headers $headers @authParams
        }
    }

    # ── Enumerate immediate sub-webs via REST ─────────────────────────────────
    function Get-SubWebs {
        param ([string]$WebUrl, [System.Net.ICredentials]$Creds)
        $useDefault = ($null -eq $Creds)
        $authParams  = if ($useDefault) { @{ UseDefaultCredentials = $true } }
                       else             { @{ Credential = $Creds } }
        $resp = Invoke-RestMethod "$WebUrl/_api/web/webs?`$select=Url" -Method Get `
                    -Headers @{ Accept = "application/json;odata=nometadata" } @authParams
        return $resp.value | Select-Object -ExpandProperty Url
    }

    # Resolve credentials once
    $netCreds = if ($Credential) { $Credential.GetNetworkCredential() } else { $null }

    # ── Banner ────────────────────────────────────────────────────────────────
    Write-Host "`n══════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host "  SharePoint SE - Modern Theme Provisioner"             -ForegroundColor Cyan
    Write-Host "══════════════════════════════════════════════════════`n" -ForegroundColor Cyan
    Write-Host "  Theme Name   : $ThemeName"
    if ($usingJson) {
        Write-Host "  Source       : SPO JSON  ($($themeData.Count) colour slots loaded)"
        Write-Host "  themePrimary : $($themeData['themePrimary'])"
        if ($themeData.ContainsKey('white'))         { Write-Host "  white        : $($themeData['white'])" }
        if ($themeData.ContainsKey('neutralPrimary')){ Write-Host "  neutralPrimary: $($themeData['neutralPrimary'])" }
    } else {
        $p = $presets[$ThemePreset]
        Write-Host "  Source       : Built-in preset '$ThemePreset'"
        Write-Host "  Primary      : $($p.primary)"
        Write-Host "  Body Text    : $($p.text)"
        Write-Host "  Background   : $($p.bg)"
    }
    Write-Host "  Sub-webs     : $(if ($ApplyToSubWebs) { 'Yes (recursive)' } else { 'No' })`n"
}

process {
    foreach ($url in $SiteUrl) {
        $url = $url.TrimEnd('/')
        Write-Host "► Processing site: $url" -ForegroundColor Yellow

        try {
            Invoke-ApplyTheme -WebUrl $url -Palette $themeData -Creds $netCreds | Out-Null
            Write-Host "  ✓ Theme applied to root web." -ForegroundColor Green

            if ($ApplyToSubWebs) {
                $subWebs = Get-SubWebs -WebUrl $url -Creds $netCreds
                if (-not $subWebs) {
                    Write-Host "  (no sub-sites found)" -ForegroundColor DarkGray
                }
                foreach ($sw in $subWebs) {
                    Write-Host "  ► Sub-web: $sw" -ForegroundColor DarkYellow
                    try {
                        Invoke-ApplyTheme -WebUrl $sw -Palette $themeData -Creds $netCreds | Out-Null
                        Write-Host "    ✓ Applied." -ForegroundColor Green
                    } catch {
                        Write-Warning "    ✗ Failed on $sw : $_"
                    }
                }
            }
        } catch {
            Write-Error "✗ Failed to apply theme on $url : $_"
        }
    }
}

end {
    Write-Host "`n══════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host "  Done.  Theme '$ThemeName' provisioning complete."     -ForegroundColor Cyan
    Write-Host "══════════════════════════════════════════════════════`n" -ForegroundColor Cyan
}