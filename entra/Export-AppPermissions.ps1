<#
.SYNOPSIS
    Exports Entra ID app permissions to an Excel report.

.DESCRIPTION
    Connects to Microsoft Graph and enumerates all service principals in the tenant,
    collecting both delegated (OAuth2) and application (app role) permissions.
    Results are exported to a formatted Excel workbook.

.PARAMETER GraphCloud
    The Microsoft Graph national cloud environment to connect to.
    Defaults to 'Global'. Accepted values: Global, USGov, USGovDoD, Germany, China.

.PARAMETER Type
    Optional. Filters the permission type to export. Use 'app' for application permissions,
    'delegate' for delegated permissions, or omit to export both.

.PARAMETER Claim
    Optional. Filters results by permission claim value. Supports wildcards.
    Examples: 'Mail.Read', 'Mail.*', 'Mail*'

.PARAMETER Tenant
    Optional. The tenant ID or domain to connect to. If omitted, uses the default tenant
    for the authenticated account.

.PARAMETER Out
    Output path for the Excel file. If omitted, saves to the current user's
    Downloads folder with an auto-generated filename.
    Example: 'C:\Reports\permissions.xlsx'

.EXAMPLE
    .\Export-AppPermissions.ps1
    Exports all app and delegated permissions for the default tenant.

.EXAMPLE
    .\Export-AppPermissions.ps1 -Type app -Claim 'Mail.*'
    Exports only application permissions whose claim matches 'Mail.*'.

.EXAMPLE
    .\Export-AppPermissions.ps1 -Tenant contoso.onmicrosoft.com -Out C:\Reports\perms.xlsx
    Exports all permissions for a specific tenant to a custom output path.
#>
#Requires -Modules Microsoft.Graph.Authentication, ImportExcel

param(
    [ValidateSet('Global', 'USGov', 'USGovDoD', 'Germany', 'China')]
    [string]$GraphCloud = 'Global',
    
    [ValidateSet('app', 'delegate')]
    [string]$Type, # Empty = both

    [string]$Claim, # Mail.Read, Mail.*, or Mail*
    [string]$Tenant,
    [string]$Out
)

function Log { param([string]$Msg,[ConsoleColor]$Color='White') Write-Host "[$(Get-Date -Format 'HH:mm')] $Msg" -ForegroundColor $Color }

$scopes = @(
    'Organization.Read.All'
    'Directory.Read.All'
    'Application.Read.All'
    'User.Read.All'
    'DelegatedPermissionGrant.Read.All'
)
Log "Connecting to Graph..."
if ($Tenant) { Connect-MgGraph -Scopes $scopes -Environment $GraphCloud -NoWelcome -ErrorAction Stop -TenantId $Tenant }
else { Connect-MgGraph -Scopes $scopes -Environment $GraphCloud -NoWelcome -ErrorAction Stop }
$Org = Get-MgOrganization
Log "Connected to Graph: $($Org.DisplayName)" Green

if (!$Out) {
    $friendlyName = "AppPermissions"
    $Out = Join-Path $env:USERPROFILE "Downloads\$($Org.DisplayName)-$friendlyName-$(Get-Date -Format 'yyyy-MM-dd_HH-mm').xlsx"
}
elseif ($Out -notlike '*.xlsx') { $Out += '.xlsx' }


$doDelegate = (-not $Type) -or ($Type -eq 'delegate')
$doApp = (-not $Type) -or ($Type -eq 'app')

Log "Fetching service principals..."

$servicePrincipals = Get-MgServicePrincipal -All -ErrorAction Stop

if (-not $servicePrincipals) {
    Log "No service principals found." Yellow
    return
}

Log "Processing $($servicePrincipals.Count) service principal(s)..."

# Simple caches
$resourceSpCache = @{}
$userCache = @{}
$rows = New-Object System.Collections.Generic.List[object]

function Get-ResourceServicePrincipal {
    param(
        [Parameter(Mandatory)]
        [string]$Id
    )

    if ($resourceSpCache.ContainsKey($Id)) {
        return $resourceSpCache[$Id]
    }

    $sp = Get-MgServicePrincipal -ServicePrincipalId $Id -ErrorAction SilentlyContinue
    if ($sp) {
        $resourceSpCache[$Id] = $sp
    }
    return $sp
}

function Get-UserNameFromId {
    param(
        [Parameter(Mandatory)]
        [string]$Id
    )

    if ($userCache.ContainsKey($Id)) {
        return $userCache[$Id]
    }

    $u = Get-MgUser -UserId $Id -ErrorAction SilentlyContinue
    $name = $null

    if ($u) {
        $name = $u.UserPrincipalName
        if (-not $name) { $name = $u.DisplayName }
    }

    if (-not $name) { $name = $Id }

    $userCache[$Id] = $name
    return $name
}

foreach ($sp in $servicePrincipals) {

    # Delegated permissions
    if ($doDelegate) {

        $grants = Get-MgServicePrincipalOauth2PermissionGrant -ServicePrincipalId $sp.Id -All -ErrorAction SilentlyContinue

        if ($grants) {
            $permIndex = @{}  # key: resourceId|scope

            foreach ($grant in $grants) {
                if (-not $grant.Scope) { continue }

                $scopesInGrant = $grant.Scope -split ' ' | Where-Object { $_ -and $_.Trim() }

                $resourceSp = Get-ResourceServicePrincipal -Id $grant.ResourceId
                if (-not $resourceSp) { continue }

                foreach ($scopeValue in $scopesInGrant) {

                    if ($Claim -and ($scopeValue -notlike $Claim)) {
                        continue
                    }

                    $permissionDef = $null
                    if ($resourceSp.Oauth2PermissionScopes) {
                        $permissionDef = $resourceSp.Oauth2PermissionScopes |
                        Where-Object { $_.Value -eq $scopeValue }
                    }

                    if ($permissionDef) {
                        $permissionDisplayName = $permissionDef.UserConsentDisplayName
                        if (-not $permissionDisplayName) { $permissionDisplayName = $permissionDef.AdminConsentDisplayName }
                    }

                    $resourceKey = "$($grant.ResourceId)|$scopeValue"

                    if (-not $permIndex.ContainsKey($resourceKey)) {
                        $row = [PSCustomObject]@{
                            DisplayName    = $sp.DisplayName
                            Type           = 'Delegated'
                            Claim          = $scopeValue
                            Description    = $permissionDisplayName
                            ResourceName   = $resourceSp.DisplayName
                            AdminConsented = $false
                            UserConsented  = $false
                            ConsentedUsers = ''
                            AppId          = $sp.AppId
                            AppObjectId    = $sp.Id
                            ResourceAppId  = $resourceSp.AppId
                        }
                        $permIndex[$resourceKey] = $row
                    }

                    $current = $permIndex[$resourceKey]

                    if ($grant.ConsentType -eq 'AllPrincipals') {
                        $current.AdminConsented = $true
                    }
                    elseif ($grant.ConsentType -eq 'Principal') {
                        $current.UserConsented = $true
                        if ($grant.PrincipalId) {
                            $userName = Get-UserNameFromId -Id $grant.PrincipalId
                            if ($userName) {
                                if ([string]::IsNullOrWhiteSpace($current.ConsentedUsers)) {
                                    $current.ConsentedUsers = $userName
                                }
                                else {
                                    $existing = $current.ConsentedUsers -split ';' | ForEach-Object { $_.Trim() }
                                    if ($existing -notcontains $userName) {
                                        $current.ConsentedUsers += ";$userName"
                                    }
                                }
                            }
                        }
                    }
                }
            }

            foreach ($entry in $permIndex.GetEnumerator()) {
                $row = $entry.Value
                $rows.Add($row)
            }
        }
    }

    # Application permissions
    if ($doApp) {

        $appAssignments = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $sp.Id -All -ErrorAction SilentlyContinue

        if ($appAssignments) {
            foreach ($assignment in $appAssignments) {
                if ($null -ne $assignment.DeletedDateTime) { continue }

                $resourceSp = Get-ResourceServicePrincipal -Id $assignment.ResourceId
                if (-not $resourceSp) { continue }

                $appRole = $null
                if ($resourceSp.AppRoles) {
                    $appRole = $resourceSp.AppRoles |
                    Where-Object { $_.Id -eq $assignment.AppRoleId }
                }

                $claimValue = 'Unknown'
                $permissionDisplay = '[UnknownAppRole]'
                if ($appRole) {
                    $claimValue = $appRole.Value
                    $permissionDisplay = $appRole.DisplayName
                }

                if ($Claim -and $claimValue -and ($claimValue -notlike $Claim)) {
                    continue
                }

                $row = [PSCustomObject]@{
                    DisplayName    = $sp.DisplayName
                    Type           = 'Application'
                    Claim          = $claimValue
                    Description    = $permissionDisplay
                    ResourceName   = $resourceSp.DisplayName
                    AdminConsented = $true
                    UserConsented  = $false
                    ConsentedUsers = ''
                    AppId          = $sp.AppId
                    AppObjectId    = $sp.Id
                    ResourceAppId  = $resourceSp.AppId
                }

                $rows.Add($row)
            }
        }
    }
}

if (-not $rows -or $rows.Count -eq 0) {
    if ($Claim) {
        Log "No permissions found matching claim '$Claim'." Yellow
    }
    else {
        Low "No permissions found." Yellow
    }
    return
}

$sheetName = switch ($Type) {
    'app' { 'AppPermissions' }
    'delegate' { 'DelegatedPermissions' }
    default { 'AllPermissions' }
}
$tableName = $sheetName

$rows |
Export-Excel -Path $Out `
    -AutoSize `
    -FreezeTopRow `
    -WorksheetName $sheetName `
    -TableName $tableName `
    -TableStyle Medium2

Log "Exported report: $Out" Green

$answer = Read-Host "Open the report now? [Y/n]"
if ($answer -eq '' -or $answer -match '^y') {
    Start-Process $Out
}