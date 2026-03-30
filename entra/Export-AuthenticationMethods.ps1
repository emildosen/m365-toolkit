<#
.SYNOPSIS
    Exports user authentication methods and tenant policy settings to an Excel report.

.DESCRIPTION
    Connects to Microsoft Graph and retrieves all users' registered authentication methods,
    sign-in preferences, and tenant-level authentication method policy configurations.
    Outputs an Excel workbook with two sheets: a Summary (method usage counts and tenant
    policy state) and a Details sheet (one row per user with a Boolean column per method).

.PARAMETER GraphCloud
    The Microsoft Graph national cloud environment to connect to.
    Defaults to 'Global'. Accepted values: Global, USGov, USGovDoD, Germany, China.

.PARAMETER Tenant
    Optional tenant ID or domain to target when authenticating to Microsoft Graph.
    If omitted, the default tenant for the authenticated account is used.

.PARAMETER Out
    Optional output path for the Excel file (.xlsx).
    Defaults to: ~\Downloads\<OrgName>-UserAuthMethods-<timestamp>.xlsx

.EXAMPLE
    .\Export-AuthenticationMethods.ps1

    Connects to the global cloud and saves the report to the Downloads folder.

.EXAMPLE
    .\Export-AuthenticationMethods.ps1 -Tenant contoso.onmicrosoft.com -Out C:\Reports\auth.xlsx

    Targets a specific tenant and writes the report to a custom path.
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
    'User.Read.All',
    'Organization.Read.All',
    'UserAuthenticationMethod.Read.All',
    'Policy.Read.All'
)
if ($Tenant) { Connect-MgGraph -Scopes $scopes -Environment $GraphCloud -NoWelcome -ErrorAction Stop -TenantId $Tenant }
else { Connect-MgGraph -Scopes $scopes -Environment $GraphCloud -NoWelcome -ErrorAction Stop }

$Org = Get-MgOrganization
Log "Connected to Graph: $($Org.DisplayName)" Green

if (!$Out) {
    $friendlyName = "UserAuthMethods"
    $Out = Join-Path $env:USERPROFILE "Downloads\$($Org.DisplayName)-$friendlyName-$(Get-Date -Format 'yyyy-MM-dd_HH-mm').xlsx"
}
elseif ($Out -notlike '*.xlsx') { $Out += '.xlsx' }

function Convert-MethodType {
    param([string]$OdataType)
    if (-not $OdataType) { return 'Unknown' }
    $t = $OdataType -replace '#microsoft.graph.', ''
    switch -Regex ($t) {
        '^emailAuthenticationMethod$' { 'Email' ; break }
        '^fido2AuthenticationMethod$' { 'FIDO2' ; break }
        '^microsoftAuthenticatorAuthenticationMethod$' { 'MSAuthenticator' ; break }
        '^passwordAuthenticationMethod$' { 'Password' ; break }
        '^phoneAuthenticationMethod$' { 'Phone' ; break }
        '^softwareOathAuthenticationMethod$' { 'SoftwareOATH' ; break }
        '^temporaryAccessPassAuthenticationMethod$' { 'TemporaryAccessPass' ; break }
        '^windowsHelloForBusinessAuthenticationMethod$' { 'WindowsHelloForBusiness' ; break }
        '^x509CertificateAuthenticationMethod$' { 'Certificate' ; break }
        default { $t }
    }
}

Log "Fetching all users..."
$users = Get-MgUser -All -Property id, displayName, userPrincipalName, accountEnabled, createdDateTime

# Collect all method types present across the directory
$allMethodTypes = New-Object System.Collections.Generic.HashSet[string]
$userMethodMap = @{}

foreach ($u in $users) {
    try {
        $methods = Get-MgUserAuthenticationMethod -UserId $u.Id -ErrorAction Stop
        $userMethodMap[$u.Id] = $methods
        foreach ($m in $methods) {
            $null = $allMethodTypes.Add( (Convert-MethodType -OdataType $m.AdditionalProperties['@odata.type']) )
        }
    }
    catch {
        Write-Warning "Failed to retrieve authentication methods for $($u.UserPrincipalName): $_"
        $userMethodMap[$u.Id] = @()
    }
}

# Create a stable, readable order for the dynamic columns
$preferredOrder = @(
    'Password', 'MSAuthenticator', 'Phone', 'Email', 'FIDO2', 'WindowsHelloForBusiness', 'SoftwareOATH', 'TemporaryAccessPass', 'Certificate', 'Unknown'
)
$dynamicColumns = @()
foreach ($name in $preferredOrder) {
    if ($allMethodTypes.Contains($name)) { $dynamicColumns += $name }
}
# Append any remaining uncommon types that were discovered
$remaining = $allMethodTypes | Where-Object { $_ -notin $dynamicColumns }
$dynamicColumns += ($remaining | Sort-Object)

# Build rows with one Boolean column per discovered method + default & system preferred status
Log "Fetching sign-in preferences..."
$rows = foreach ($u in $users) {
    $row = [ordered]@{
        DisplayName       = $u.DisplayName
        UserPrincipalName = $u.UserPrincipalName
        AccountEnabled    = $u.AccountEnabled
        CreatedDate       = $u.CreatedDateTime
        MethodCount       = 0
        DefaultAuthMethod = $null
        IsSystemPreferred = $null
    }
    foreach ($col in $dynamicColumns) { $row[$col] = $false }

    $count = 0

    foreach ($m in $userMethodMap[$u.Id]) {
        $typeName = Convert-MethodType -OdataType $m.AdditionalProperties['@odata.type']
        if ($row.Contains($typeName)) { $row[$typeName] = $true }
        $count++
    }

    # Beta, get default authentication method or the system preferred method
    try {
        $pref = Get-MgBetaUserAuthenticationSignInPreference -UserId $u.Id -ErrorAction Stop
        $row['IsSystemPreferred'] = [bool]$pref.IsSystemPreferredAuthenticationMethodEnabled
        $row['PreferredMethod'] = $pref.UserPreferredMethodForSecondaryAuthentication

        $sysPref = $null
        if ($pref.PSObject.Properties.Name -contains 'SystemPreferredAuthenticationMethod') {
            $sysPref = $pref.SystemPreferredAuthenticationMethod
        }
        elseif ($pref.AdditionalProperties -and $pref.AdditionalProperties.ContainsKey('systemPreferredAuthenticationMethod')) {
            $sysPref = $pref.AdditionalProperties['systemPreferredAuthenticationMethod']
        }
        if ($sysPref) { $row['SystemPreferredMethod'] = $sysPref }

        # System preferred wins if enabled; otherwise, user's preferred selection
        if ($row['IsSystemPreferred']) { $row['DefaultAuthMethod'] = $row['SystemPreferredMethod'] }
        else { $row['DefaultAuthMethod'] = $row['PreferredMethod'] }
    }
    catch {
        # leave preference fields null if not available
    }

    $row['MethodCount'] = $count
    [PSCustomObject]$row
}

# Summary data 
$summaryCounts = foreach ($name in $dynamicColumns) {
    [PSCustomObject]@{
        Method          = $name
        UsersConfigured = ($rows | Where-Object { $_.$name -eq $true }).Count
    }
}

# Tenant authentication method configurations (enabled/disabled)
Log "Fetching tenant authentication settings..."
$summaryConfig = @()
try {
    $policy = Get-MgPolicyAuthenticationMethodPolicy -ErrorAction Stop
    $configs = $policy.AuthenticationMethodConfigurations

    foreach ($c in $configs) {
        # Prefer AdditionalProperties['@odata.type'] when present; else fall back to Id mapping
        $odata = $null
        if ($c.AdditionalProperties -and $c.AdditionalProperties.ContainsKey('@odata.type')) {
            $odata = [string]$c.AdditionalProperties['@odata.type']
            $odata = $odata -replace '#microsoft.graph.', ''
        }

        $name = switch ($odata) {
            'emailAuthenticationMethodConfiguration' { 'Email' }
            'fido2AuthenticationMethodConfiguration' { 'FIDO2' }
            'microsoftAuthenticatorAuthenticationMethodConfiguration' { 'MSAuthenticator' }
            'passwordAuthenticationMethodConfiguration' { 'Password' }
            'phoneAuthenticationMethodConfiguration' { 'Phone' }
            'softwareOathAuthenticationMethodConfiguration' { 'SoftwareOATH' }
            'temporaryAccessPassAuthenticationMethodConfiguration' { 'TemporaryAccessPass' }
            'windowsHelloForBusinessAuthenticationMethodConfiguration' { 'WindowsHelloForBusiness' }
            'x509CertificateAuthenticationMethodConfiguration' { 'Certificate' }
            'smsAuthenticationMethodConfiguration' { 'SMS' }
            'voiceAuthenticationMethodConfiguration' { 'Voice' }
            Default {
                switch ($c.Id) {
                    'Fido2' { 'FIDO2' }
                    'MicrosoftAuthenticator' { 'MSAuthenticator' }
                    'Password' { 'Password' }
                    'TemporaryAccessPass' { 'TemporaryAccessPass' }
                    'SoftwareOath' { 'SoftwareOATH' }
                    'Email' { 'Email' }
                    'Voice' { 'Voice' }
                    'Sms' { 'SMS' }
                    'Phone' { 'Phone' }
                    'WindowsHelloForBusiness' { 'WindowsHelloForBusiness' }
                    'X509Certificate' { 'Certificate' }
                    Default { if ($c.DisplayName) { $c.DisplayName } else { $c.Id } }
                }
            }
        }

        $enabled = $null
        if ($c.PSObject.Properties.Name -contains 'State') { $enabled = ($c.State -eq 'enabled') }
        elseif ($c.PSObject.Properties.Name -contains 'IsEnabled') { $enabled = [bool]$c.IsEnabled }

        $summaryConfig += [PSCustomObject]@{
            Method  = $name
            Enabled = $enabled
        }
    }
}
catch {
    Write-Warning "Failed to retrieve authentication method policy: $_"
}


# Export to Excel
$excelParams = @{
    Path         = $Out
    AutoSize     = $true
    FreezeTopRow = $true
    TableStyle   = 'Medium2'
}
# User method summary
$summaryCounts | Export-Excel @excelParams -WorksheetName 'Summary' -TableName 'AuthMethodCounts'
# Tenant auth methods summary
$startRow = ($summaryCounts.Count + 3)
$summaryConfig | Export-Excel -WorksheetName 'Summary' -TableName 'AuthMethodConfig' -StartRow $startRow @excelParams
# Details
$rows | Export-Excel @excelParams -WorksheetName 'Details' -TableName 'UserAuthDetails'

Log "Exported report: $Out" Green

$answer = Read-Host "Open the report now? [Y/n]"
if ($answer -eq '' -or $answer -match '^y') {
    Start-Process $Out
}