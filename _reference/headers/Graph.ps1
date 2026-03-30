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
    'Organization.Read.All'
)
Log "Connecting to Graph..."
if ($Tenant) { Connect-MgGraph -Scopes $scopes -Environment $GraphCloud -NoWelcome -ErrorAction Stop -TenantId $Tenant }
else { Connect-MgGraph -Scopes $scopes -Environment $GraphCloud -NoWelcome -ErrorAction Stop }
$Org = Get-MgOrganization
Log "Connected to Graph: $($Org.DisplayName)" Green

if (!$Out) {
    # CHANGE FRIENDLY NAME
    $friendlyName = "UserAuthMethods"
    $Out = Join-Path $env:USERPROFILE "Downloads\$($Org.DisplayName)-$friendlyName-$(Get-Date -Format 'yyyy-MM-dd_HH-mm').xlsx"
}
elseif ($Out -notlike '*.xlsx') { $Out += '.xlsx' }