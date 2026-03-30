<#
.SYNOPSIS
    Exports Entra ID Conditional Access policies and named locations to an Excel report.

.DESCRIPTION
    Connects to Microsoft Graph and retrieves all Conditional Access policies along with
    named locations. Resolves user, group, role, and application IDs to display names,
    then exports the results to a formatted Excel workbook with two worksheets.

.PARAMETER GraphCloud
    The Microsoft Graph national cloud environment to connect to.
    Defaults to 'Global'. Valid values: Global, USGov, USGovDoD, Germany, China.

.PARAMETER Tenant
    The tenant ID or domain name to authenticate against. If omitted, the default
    tenant for the authenticated account is used.

.PARAMETER Out
    Full path for the output Excel file (.xlsx). If omitted, the file is saved to the
    current user's Downloads folder with an auto-generated name including the tenant
    display name and timestamp.

.EXAMPLE
    .\Export-ConditionalAccess.ps1

.EXAMPLE
    .\Export-ConditionalAccess.ps1 -Tenant contoso.onmicrosoft.com -Out C:\Reports\CA.xlsx
#>
#Requires -Modules Microsoft.Graph.Authentication, ImportExcel

[CmdletBinding()]
param(
    [ValidateSet('Global', 'USGov', 'USGovDoD', 'Germany', 'China')]
    [string]$GraphCloud = 'Global',
    [string]$Tenant,
    [string]$Out
)

function Log { param([string]$Msg, [ConsoleColor]$Color = 'White') Write-Host "[$(Get-Date -Format 'HH:mm')] $Msg" -ForegroundColor $Color }

$scopes = @(
    'Policy.Read.ConditionalAccess',
    'User.Read.All',
    'Group.Read.All',
    'RoleManagement.Read.Directory',
    'Application.Read.All',
    'Policy.Read.All',
    'Organization.Read.All'
)
Log "Connecting to Graph..."
if ($Tenant) { Connect-MgGraph -Scopes $scopes -Environment $GraphCloud -NoWelcome -ErrorAction Stop -TenantId $Tenant }
else { Connect-MgGraph -Scopes $scopes -Environment $GraphCloud -NoWelcome -ErrorAction Stop }
$Org = Get-MgOrganization
Log "Connected to Graph: $($Org.DisplayName)" Green

if (!$Out) {
    $friendlyName = "ConditionalAccess"
    $Out = Join-Path $env:USERPROFILE "Downloads\$($Org.DisplayName)-$friendlyName-$(Get-Date -Format 'yyyy-MM-dd_HH-mm').xlsx"
}
elseif ($Out -notlike '*.xlsx') { $Out += '.xlsx' }


# Helpers
function Convert-DateString {
    param(
        [string]$InputString
    )

    if ([string]::IsNullOrWhiteSpace($InputString)) {
        return $null
    }

    try {
        $dt = [datetime]::Parse($InputString, [System.Globalization.CultureInfo]::InvariantCulture)
        return $dt.ToString('yyyy-MM-dd HH:mm')
    }
    catch {
        return $null
    }
}

# Main script

Log "Fetching conditional access policies and directory roles..."
# Get all conditional access policies
$policies = Get-MgIdentityConditionalAccessPolicy -All

# Get directory roles and role templates for lookup
$directoryRoleTemplates = Get-MgDirectoryRoleTemplate -All
$roleTemplateMap = @{}
foreach ($template in $directoryRoleTemplates) {
    $roleTemplateMap[$template.Id] = $template.DisplayName
}

Log "Fetching users, groups, and apps..."
# Get users, groups, apps for lookup
$allUsers = @{}
$allGroups = @{}
$allApps = @{}

Get-MgUser -All -Property Id, UserPrincipalName | ForEach-Object {
    $allUsers[$_.Id] = $_.UserPrincipalName
}

Get-MgGroup -All -Property Id, DisplayName | ForEach-Object {
    $allGroups[$_.Id] = $_.DisplayName
}

Get-MgApplication -All -Property Id, DisplayName, AppId | ForEach-Object {
    $allApps[$_.AppId] = $_.DisplayName
}

Get-MgServicePrincipal -All -Property AppId, DisplayName | ForEach-Object {
    if (-not $allApps.ContainsKey($_.AppId)) {
        $allApps[$_.AppId] = $_.DisplayName
    }
}


Log "Fetching and processing named locations..."
# Get named locations before processing policies
$namedLocations = Get-MgIdentityConditionalAccessNamedLocation -All
$locationMap = @{}
$locationDetails = @()

foreach ($location in $namedLocations) {
    $locationMap[$location.Id] = $location.DisplayName
    
    # Collect location details for separate sheet
    $locationType = 'Unknown'
    $locationValues = @()
    
    if ($location.AdditionalProperties -and $location.AdditionalProperties['@odata.type']) {
        if ($location.AdditionalProperties['@odata.type'] -eq '#microsoft.graph.countryNamedLocation') {
            $locationType = 'Country'
            $locationValues = $location.AdditionalProperties.countriesAndRegions
        }
        elseif ($location.AdditionalProperties['@odata.type'] -eq '#microsoft.graph.ipNamedLocation') {
            $locationType = 'IP Range'
            $locationValues = $location.AdditionalProperties.ipRanges | ForEach-Object {
                if ($_.cidrAddress) { $_.cidrAddress } else { "$($_.addressPrefix)/$($_.prefixLength)" }
            }
        }
    }
    
    $locationDetails += [PSCustomObject]@{
        'Location Name' = $location.DisplayName
        'Type'          = $locationType
        'Values'        = $locationValues -join '; '
        'Is Trusted'    = if ($location.AdditionalProperties.isTrusted) { 'Yes' } else { 'No' }
        'Created'       = Convert-DateString $location.CreatedDateTime
        'Modified'      = Convert-DateString $location.ModifiedDateTime
    }
}


Log "Processing $($policies.Count) conditional access policies..."
# Process policies
$results = foreach ($policy in $policies) {
    # Process state
    $state = switch ($policy.State) {
        'enabledForReportingButNotEnforced' { 'report-only' }
        default { $policy.State }
    }
    
    # Process includes and excludes
    $includedObjects = @()
    $excludedObjects = @()
    
    # Included users
    if ($policy.Conditions.Users.IncludeUsers) {
        foreach ($userId in $policy.Conditions.Users.IncludeUsers) {
            if ($userId -eq 'All') {
                $includedObjects += 'All Users'
            }
            elseif ($userId -eq 'GuestsOrExternalUsers') {
                $includedObjects += 'Guests or External Users'
            }
            else {
                $includedObjects += $allUsers[$userId] ?? $userId
            }
        }
    }

    # Included guests/external users
    if ($policy.Conditions.Users.IncludeGuestsOrExternalUsers) {
        $types = $policy.Conditions.Users.IncludeGuestsOrExternalUsers.GuestOrExternalUserTypes
        $ext = $policy.Conditions.Users.IncludeGuestsOrExternalUsers.ExternalTenants

        $label = ($types) ? "Guests or External Users: $types" : 'Guests or External Users'

        $extLabel = $null
        if ($ext) {
            if ($ext.PSObject.Properties.Match('MembershipKind').Count -gt 0 -and $ext.MembershipKind) {
                if ($ext.MembershipKind -eq 'all') { $extLabel = 'ExternalTenants: All' }
            }
            if (-not $extLabel -and $ext.PSObject.Properties.Match('Members').Count -gt 0 -and $ext.Members) {
                $extLabel = "ExternalTenants: $($ext.Members -join ',')"
            }

            if (-not $extLabel) {
                $pairs = $ext.PSObject.Properties |
                Where-Object { $_.Name -ne 'AdditionalProperties' -and $_.Value } |
                ForEach-Object {
                    if ($_.Value -is [System.Array]) { "$($_.Name)=$($($_.Value -join ','))" }
                    elseif ($_.Value -isnot [System.Collections.IDictionary]) { "$($_.Name)=$($_.Value)" }
                }
                $extSummary = ($pairs | Where-Object { $_ }) -join '; '
                if ($extSummary) { $extLabel = "ExternalTenants: $extSummary" }
            }
        }

        if ($extLabel) { $label = "$label ($extLabel)" }
        $includedObjects += $label
    }
    
    # Included groups
    if ($policy.Conditions.Users.IncludeGroups) {
        foreach ($groupId in $policy.Conditions.Users.IncludeGroups) {
            $includedObjects += "Group: $($allGroups[$groupId] ?? $groupId)"
        }
    }
    
    # Included roles
    if ($policy.Conditions.Users.IncludeRoles) {
        foreach ($roleId in $policy.Conditions.Users.IncludeRoles) {
            $roleName = $roleTemplateMap[$roleId] ?? $roleId
            $includedObjects += "Role: $roleName"
        }
    }
    
    # Excluded users
    if ($policy.Conditions.Users.ExcludeUsers) {
        foreach ($userId in $policy.Conditions.Users.ExcludeUsers) {
            if ($userId -eq 'GuestsOrExternalUsers') {
                $excludedObjects += 'Guests or External Users'
            }
            else {
                $excludedObjects += $allUsers[$userId] ?? $userId
            }
        }
    }

    # Excluded guests/external users
    if ($policy.Conditions.Users.ExcludeGuestsOrExternalUsers) {
        $types = $policy.Conditions.Users.ExcludeGuestsOrExternalUsers.GuestOrExternalUserTypes
        $ext = $policy.Conditions.Users.ExcludeGuestsOrExternalUsers.ExternalTenants

        $label = ($types) ? "Guests or External Users: $types" : 'Guests or External Users'

        $extLabel = $null
        if ($ext) {
            if ($ext.PSObject.Properties.Match('MembershipKind').Count -gt 0 -and $ext.MembershipKind) {
                if ($ext.MembershipKind -eq 'all') { $extLabel = 'ExternalTenants: All' }
            }
            if (-not $extLabel -and $ext.PSObject.Properties.Match('Members').Count -gt 0 -and $ext.Members) {
                $extLabel = "ExternalTenants: $($ext.Members -join ',')"
            }

            if (-not $extLabel) {
                $pairs = $ext.PSObject.Properties |
                Where-Object { $_.Name -ne 'AdditionalProperties' -and $_.Value } |
                ForEach-Object {
                    if ($_.Value -is [System.Array]) { "$($_.Name)=$($($_.Value -join ','))" }
                    elseif ($_.Value -isnot [System.Collections.IDictionary]) { "$($_.Name)=$($_.Value)" }
                }
                $extSummary = ($pairs | Where-Object { $_ }) -join '; '
                if ($extSummary) { $extLabel = "ExternalTenants: $extSummary" }
            }
        }

        if ($extLabel) { $label = "$label ($extLabel)" }
        $excludedObjects += $label
    }
    
    # Excluded groups
    if ($policy.Conditions.Users.ExcludeGroups) {
        foreach ($groupId in $policy.Conditions.Users.ExcludeGroups) {
            $excludedObjects += "Group: $($allGroups[$groupId] ?? $groupId)"
        }
    }
    
    # Excluded roles
    if ($policy.Conditions.Users.ExcludeRoles) {
        foreach ($roleId in $policy.Conditions.Users.ExcludeRoles) {
            $roleName = $roleTemplateMap[$roleId] ?? $roleId
            $excludedObjects += "Role: $roleName"
        }
    }
    
    # Process target applications and user actions
    $targets = @()
    
    # Applications
    if ($policy.Conditions.Applications.IncludeApplications) {
        foreach ($appId in $policy.Conditions.Applications.IncludeApplications) {
            if ($appId -eq 'All') {
                $targets += 'All resources'
            }
            elseif ($appId -eq 'Office365') {
                $targets += 'Office 365'
            }
            else {
                $targets += $allApps[$appId] ?? $appId
            }
        }
    }
    
    # User actions
    if ($policy.Conditions.Applications.IncludeUserActions) {
        foreach ($action in $policy.Conditions.Applications.IncludeUserActions) {
            $actionName = switch ($action) {
                'urn:user:registersecurityinfo' { 'Register security information' }
                'urn:user:registerdevice' { 'Register or join devices' }
                default { $action }
            }
            $targets += "User Action: $actionName"
        }
    }
    
    if ($policy.Conditions.Applications.ExcludeApplications) {
        foreach ($appId in $policy.Conditions.Applications.ExcludeApplications) {
            $targets += "Excluded: $($allApps[$appId] ?? $appId)"
        }
    }
    
    # Process conditions
    $conditions = @()
    
    # Locations with named location display names
    if ($policy.Conditions.Locations) {
        if ($policy.Conditions.Locations.IncludeLocations) {
            $includeLocNames = @()
            foreach ($locId in $policy.Conditions.Locations.IncludeLocations) {
                if ($locId -eq 'All') {
                    $includeLocNames += 'All locations'
                }
                elseif ($locId -eq 'AllTrusted') {
                    $includeLocNames += 'All trusted locations'
                }
                else {
                    $includeLocNames += $locationMap[$locId] ?? $locId
                }
            }
            $conditions += "Include Locations: $($includeLocNames -join ', ')"
        }
        if ($policy.Conditions.Locations.ExcludeLocations) {
            $excludeLocNames = @()
            foreach ($locId in $policy.Conditions.Locations.ExcludeLocations) {
                if ($locId -eq 'AllTrusted') {
                    $excludeLocNames += 'All trusted locations'
                }
                else {
                    $excludeLocNames += $locationMap[$locId] ?? $locId
                }
            }
            $conditions += "Exclude Locations: $($excludeLocNames -join ', ')"
        }
    }
    
    # Platforms
    if ($policy.Conditions.Platforms) {
        if ($policy.Conditions.Platforms.IncludePlatforms) {
            $conditions += "Include Platforms: $($policy.Conditions.Platforms.IncludePlatforms -join ', ')"
        }
        if ($policy.Conditions.Platforms.ExcludePlatforms) {
            $conditions += "Exclude Platforms: $($policy.Conditions.Platforms.ExcludePlatforms -join ', ')"
        }
    }
    
    # Client app types
    if ($policy.Conditions.ClientAppTypes) {
        $conditions += "Client Apps: $($policy.Conditions.ClientAppTypes -join ', ')"
    }
    
    # Risk levels
    if ($policy.Conditions.UserRiskLevels) {
        $conditions += "User Risk: $($policy.Conditions.UserRiskLevels -join ', ')"
    }
    if ($policy.Conditions.SignInRiskLevels) {
        $conditions += "Sign-in Risk: $($policy.Conditions.SignInRiskLevels -join ', ')"
    }
    
    # Process grant controls
    $grantBlock = @()
    if ($policy.GrantControls) {
        if ($policy.GrantControls.BuiltInControls) {
            if ($policy.GrantControls.BuiltInControls -contains 'block') {
                $grantBlock += 'Block access'
            }
            else {
                $operator = if ($policy.GrantControls.Operator -eq 'AND') { ' AND ' } else { ' OR ' }
                $controls = $policy.GrantControls.BuiltInControls | ForEach-Object {
                    switch ($_) {
                        'mfa' { 'Require MFA' }
                        'compliantDevice' { 'Require compliant device' }
                        'domainJoinedDevice' { 'Require domain joined device' }
                        'approvedApplication' { 'Require approved app' }
                        'compliantApplication' { 'Require app protection policy' }
                        'passwordChange' { 'Require password change' }
                        default { $_ }
                    }
                }
                $grantBlock += "Grant: $($controls -join $operator)"
            }
        }
        
        if ($policy.GrantControls.CustomAuthenticationFactors) {
            $grantBlock += "Custom factors: $($policy.GrantControls.CustomAuthenticationFactors -join ', ')"
        }
        
        if ($policy.GrantControls.TermsOfUse) {
            $grantBlock += "Terms of use required"
        }
    }
    
    # Process session controls
    $sessionControls = @()
    if ($policy.SessionControls) {
        if ($policy.SessionControls.ApplicationEnforcedRestrictions.IsEnabled) {
            $sessionControls += 'App enforced restrictions'
        }
        if ($policy.SessionControls.CloudAppSecurity.IsEnabled) {
            $sessionControls += "Cloud App Security: $($policy.SessionControls.CloudAppSecurity.CloudAppSecurityType)"
        }
        if ($policy.SessionControls.SignInFrequency.IsEnabled) {
            $sessionControls += "Sign-in frequency: $($policy.SessionControls.SignInFrequency.Value) $($policy.SessionControls.SignInFrequency.Type)"
        }
        if ($policy.SessionControls.PersistentBrowser.IsEnabled) {
            $sessionControls += "Persistent browser: $($policy.SessionControls.PersistentBrowser.Mode)"
        }
    }
    
    # Create output object
    [PSCustomObject]@{
        'Policy Name'      = $policy.DisplayName
        'State'            = $state
        'Included Objects' = $includedObjects -join '; '
        'Excluded Objects' = $excludedObjects -join '; '
        'Targets'          = $targets -join '; '
        'Conditions'       = $conditions -join '; '
        'Grant/Block'      = $grantBlock -join '; '
        'Session'          = $sessionControls -join '; '
        'Created'          = Convert-DateString $policy.CreatedDateTime
        'Modified'         = Convert-DateString $policy.ModifiedDateTime
    }
}



# Export to Excel
$excelParams = @{
    Path         = $Out
    TableStyle   = "Medium2"
    AutoSize     = $true
    FreezeTopRow = $true
}
$results | Export-Excel @excelParams -WorksheetName 'Conditional Access Policies'
$excel = $locationDetails | Export-Excel @excelParams -WorksheetName 'Named Locations' -PassThru
    
# Set column width
$ws = $excel.Workbook.Worksheets["Conditional Access Policies"]
$ws.Column(3).Width = 20  # Included Objects
$ws.Column(4).Width = 20  # Excluded Objects
$ws.Column(5).Width = 20  # Targets
$ws.Column(6).Width = 20  # Conditions
$ws.Column(7).Width = 20  # Grant/Block
$ws.Column(8).Width = 20  # Session
$ws.Column(9).AutoFit()   # Created
$ws.Column(9).Width += 3
$ws.Column(10).AutoFit()  # Modified
$ws.Column(10).Width += 3
    
$ws = $excel.Workbook.Worksheets["Named Locations"]
$ws.Column(3).Width = 50  # Values
$ws.Column(5).AutoFit()  # Created
$ws.Column(5).Width += 3
$ws.Column(6).AutoFit()  # Modified
$ws.Column(6).Width += 3
    
Close-ExcelPackage $excel
    
Log "Exported report: $Out" Green

$answer = Read-Host "Open the report now? [Y/n]"
if ($answer -eq '' -or $answer -match '^y') {
    Start-Process $Out
}