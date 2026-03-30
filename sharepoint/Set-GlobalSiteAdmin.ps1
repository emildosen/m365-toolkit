<#
.SYNOPSIS
    Adds a user as site collection administrator to all SharePoint sites in a tenant.

.DESCRIPTION
    Connects to SharePoint Online using PnP PowerShell (interactive auth) and adds the
    specified user as a site collection admin on every site, excluding OneDrive, the
    admin center, and system sites. Results are exported to an Excel report.
    You must have a app registration with appropriate permissions. See:
    https://pnp.github.io/powershell/articles/registerapplication.html

.PARAMETER AdminUrl
    URL of the SharePoint Admin Center (e.g. https://contoso-admin.sharepoint.com).

.PARAMETER AppId
    Client ID of the Entra ID app registration used for authentication.

.PARAMETER Tenant
    Tenant ID or domain (e.g. contoso.onmicrosoft.com).

.PARAMETER AddAdmin
    UPN of the user to add as site collection administrator (e.g. admin@contoso.com).

.PARAMETER Cloud
    Azure cloud environment. Defaults to 'Production'.
    Accepted values: Production, PPE, China, Germany, USGovernment, USGovernmentHigh, USGovernmentDoD.

.PARAMETER Site
    Optional. URL of a single site to target. Use for testing before running against all sites.

.PARAMETER Out
    Optional. Full path for the output Excel report. Defaults to
    ~/Downloads/<tenant>-GlobalSiteAdmin-<timestamp>.xlsx.

.EXAMPLE
    .\Set-GlobalSiteAdmin.ps1 -AdminUrl https://contoso-admin.sharepoint.com `
        -AppId 00000000-0000-0000-0000-000000000000 `
        -Tenant contoso.onmicrosoft.com `
        -AddAdmin newadmin@contoso.com

.EXAMPLE
    .\Set-GlobalSiteAdmin.ps1 -AdminUrl https://contoso-admin.sharepoint.com `
        -AppId 00000000-0000-0000-0000-000000000000 `
        -Tenant contoso.onmicrosoft.com `
        -AddAdmin newadmin@contoso.com `
        -Site https://contoso.sharepoint.com/sites/TestSite
#>
#Requires -Modules PnP.PowerShell, ImportExcel


[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidatePattern('^https?://')]
    [string]$AdminUrl,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$AppId,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$Tenant,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$AddAdmin,

    [Parameter()]
    [ValidateSet('Production', 'PPE', 'China', 'Germany', 'USGovernment', 'USGovernmentHigh', 'USGovernmentDoD')]
    [string]$Cloud = 'Production',

    [Parameter()]
    [string]$Site,

    [Parameter()]
    [string]$Out
)

function Log { param([string]$Msg, [ConsoleColor]$Color = 'White') Write-Host "[$(Get-Date -Format 'HH:mm')] $Msg" -ForegroundColor $Color }

Log "Connecting to SharePoint Admin..."
Connect-PnPOnline -Url $AdminUrl -ClientId $AppId -Interactive -Tenant $Tenant -AzureEnvironment $Cloud -ErrorAction Stop

$tenantName = ($AdminUrl -replace 'https?://', '' -replace '-admin\.sharepoint\..+$', '')
Log "Connected to SharePoint: $tenantName" Green

# Determine output file name
if (!$Out) {
    $Out = Join-Path $env:USERPROFILE "Downloads\$tenantName-GlobalSiteAdmin-$(Get-Date -Format 'yyyy-MM-dd_HH-mm').xlsx"
} elseif ($Out -notlike '*.xlsx') { $Out += '.xlsx' }

$results = [System.Collections.Generic.List[object]]::new()

if ($Site) {
    # Single site mode for testing
    Log "Single site mode: $Site"
    $sites = @([PSCustomObject]@{ Url = $Site; Title = $Site })
} else {
    # Get all tenant sites (excludes OneDrive by default)
    Log "Fetching all SharePoint sites..."
    $sites = Get-PnPTenantSite | Where-Object {
        $_.Url -notlike '*-admin.sharepoint.*' -and
        $_.Url -notlike '*-my.sharepoint.*' -and
        $_.Url -notlike '*/portals/hub' -and
        $_.Url -notlike '*/search'
    }
    Log "Found $($sites.Count) sites to process"
}

$i = 0
foreach ($s in $sites) {
    $i++
    $siteUrl = $s.Url
    $siteTitle = $s.Title

    Log "[$i/$($sites.Count)] Processing: $siteTitle ($siteUrl)"

    try {
        Set-PnPTenantSite -Identity $siteUrl -Owners $AddAdmin -ErrorAction Stop

        $results.Add([PSCustomObject]@{
            SiteUrl   = $siteUrl
            SiteTitle = $siteTitle
            Status    = 'Success'
            Error     = $null
        })
        Log "  Added $AddAdmin as admin"
    } catch {
        $results.Add([PSCustomObject]@{
            SiteUrl   = $siteUrl
            SiteTitle = $siteTitle
            Status    = 'Failed'
            Error     = $_.Exception.Message
        })
        Log "  Failed: $($_.Exception.Message)" Red
    }
}

# Export results
$results | Export-Excel -Path $Out -WorksheetName 'Results' -AutoSize -FreezeTopRow -TableName 'SiteAdminResults' -TableStyle Medium2

Log "Summary: $($results.Where({$_.Status -eq 'Success'}).Count) succeeded, $($results.Where({$_.Status -eq 'Failed'}).Count) failed"
Log "Exported report: $Out" Green

$answer = Read-Host "Open the report now? [Y/n]"
if ($answer -eq '' -or $answer -match '^y') {
    Start-Process $Out
}