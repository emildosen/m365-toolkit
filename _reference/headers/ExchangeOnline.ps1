#Requires -Modules ExchangeOnlineManagement, ImportExcel

[CmdletBinding()]
param(
    [ValidateSet('O365Default', 'O365USGovGCCHigh', 'O365USGovDoD', 'O365GermanyCloud', 'O365China')]
    [string]$ExchangeCloud = 'O365Default',
    [string]$Tenant,
    [string]$Out
)

function Log { param([string]$Msg, [ConsoleColor]$Color = 'White') Write-Host "[$(Get-Date -Format 'HH:mm')] $Msg" -ForegroundColor $Color }

Log "Connecting to ExchangeOnline..."
$connectParams = @{
    ExchangeEnvironmentName = $ExchangeCloud
    DisableWAM              = $true
    ShowBanner              = $false
    ErrorAction             = 'Stop'
}
if ($Tenant) { $connectParams['Organization'] = $Tenant }
Connect-ExchangeOnline @connectParams
$Org = Get-OrganizationConfig
Log "Connected to Exchange: $($Org.DisplayName)" Green

if (!$Out) {
    # CHANGE FRIENDLY NAME
    $friendlyName = "EmailAliases"
    $Out = Join-Path $env:USERPROFILE "Downloads\$($Org.DisplayName)-$friendlyName-$(Get-Date -Format 'yyyy-MM-dd_HH-mm').xlsx"
}
elseif ($Out -notlike '*.xlsx') { $Out += '.xlsx' }