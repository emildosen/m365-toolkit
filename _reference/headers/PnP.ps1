#Requires -Modules PnP.PowerShell, ImportExcel

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidatePattern('^https?://')]
    [string]$Url,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$AppId,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$Tenant,

    [ValidateSet('Production', 'PPE', 'China', 'Germany', 'USGovernment', 'USGovernmentHigh', 'USGovernmentDoD')]
    [string]$Cloud = 'Production',
    [string]$Out
)

function Log { param([string]$Msg, [ConsoleColor]$Color = 'White') Write-Host "[$(Get-Date -Format 'HH:mm')] $Msg" -ForegroundColor $Color }

Log "Connecting to PnP..."
Connect-PnPOnline -Url $Url -ClientId $AppId -Interactive -Tenant $Tenant -AzureEnvironment $Cloud -ErrorAction Stop
$tenantName = ($Tenant.Trim() -split '\.')[0]
Log "Connected to PnP: $tenantName" Green

if (!$Out) {
    # CHANGE FRIENDLY NAME
    $friendlyName = "FileLockCheck"
    $Out = Join-Path $env:USERPROFILE "Downloads\$($Org.DisplayName)-$friendlyName-$(Get-Date -Format 'yyyy-MM-dd_HH-mm').xlsx"
}
elseif ($Out -notlike '*.xlsx') { $Out += '.xlsx' }