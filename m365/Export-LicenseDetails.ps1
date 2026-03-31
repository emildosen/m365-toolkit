<#
.SYNOPSIS
    Exports Microsoft 365 license and subscription details to an Excel report.

.DESCRIPTION
    Connects to Microsoft Graph to retrieve subscription and license assignment data for
    the tenant. Resolves SKU part numbers to friendly display names using Microsoft's
    published mapping, calculates commitment and billing terms, and exports the results
    to a formatted Excel workbook.

.PARAMETER GraphCloud
    The Microsoft Graph cloud environment to connect to.
    Valid values: Global, USGov, USGovDoD, Germany, China. Defaults to Global.

.PARAMETER Tenant
    The tenant ID to connect to. If omitted, the default tenant for the authenticated
    account is used.

.PARAMETER Out
    Output file path for the Excel report. Defaults to the user's Downloads folder
    with a timestamped filename. Automatically appends .xlsx if not present.

.EXAMPLE
    .\Export-LicenseDetails.ps1

.EXAMPLE
    .\Export-LicenseDetails.ps1 -Tenant "contoso.onmicrosoft.com" -Out "C:\Reports\licenses.xlsx"

.EXAMPLE
    .\Export-LicenseDetails.ps1 -GraphCloud USGov
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
    "Directory.Read.All",
    'Organization.Read.All'
)
Log "Connecting to Graph..."
if ($Tenant) { Connect-MgGraph -Scopes $scopes -Environment $GraphCloud -NoWelcome -ErrorAction Stop -TenantId $Tenant }
else { Connect-MgGraph -Scopes $scopes -Environment $GraphCloud -NoWelcome -ErrorAction Stop }
$Org = Get-MgOrganization
Log "Connected to Graph: $($Org.DisplayName)" Green

if (!$Out) {
    $friendlyName = "LicenseDetails"
    $Out = Join-Path $env:USERPROFILE "Downloads\$($Org.DisplayName)-$friendlyName-$(Get-Date -Format 'yyyy-MM-dd_HH-mm').xlsx"
}
elseif ($Out -notlike '*.xlsx') { $Out += '.xlsx' }

# Download Microsoft's SKU display names
Log "Downloading SKU display names..."
$csv = Invoke-RestMethod "https://download.microsoft.com/download/e/3/e/e3e9faf2-f28b-490a-9ada-c6089a1fc5b0/Product%20names%20and%20service%20plan%20identifiers%20for%20licensing.csv"
$productNames = $csv | ConvertFrom-Csv

# Create lookup table
$skuLookup = @{}
foreach ($product in $productNames | Sort-Object String_Id -Unique) {
    if ($product.String_Id) {
        $skuLookup[$product.String_Id] = $product.Product_Display_Name
    }
}

# Get subscription data
Log "Fetching subscription data..."
$subs = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/directory/subscriptions"

# Get SKU data for assigned licenses
$skus = Get-MgSubscribedSku
$skuAssigned = @{}
foreach ($sku in $skus) {
    $skuAssigned[$sku.SkuId] = @{
        Consumed  = $sku.ConsumedUnits
        Total     = $sku.PrepaidUnits.Enabled
        Available = $sku.PrepaidUnits.Enabled - $sku.ConsumedUnits
    }
}

# Create results array
$results = @()

foreach ($sub in $subs.value) {
    # Handle null dates
    $createdDate = if ($sub.createdDateTime) { [DateTime]$sub.createdDateTime } else { $null }
    $renewalDate = if ($sub.nextLifecycleDateTime) { [DateTime]$sub.nextLifecycleDateTime } else { $null }
    
    # Calculate commitment term
    $commitmentTerm = "Unknown"
    $billingTerm = "Unknown"
    if ($createdDate -and $renewalDate) {
        $termDays = ($renewalDate - $createdDate).Days
        if ($termDays -le 31) { 
            $commitmentTerm = "Monthly"
            $billingTerm = "Monthly"
        }
        elseif ($termDays -le 366) { 
            $commitmentTerm = "1-Year"
            $billingTerm = "Annual"
        }
        elseif ($termDays -le 732) { 
            $commitmentTerm = "2-Year"
            $billingTerm = "Annual"
        }
        else { 
            $commitmentTerm = "3-Year"
            $billingTerm = "Annual"
        }
    }
    elseif ($sub.isTrial) {
        $commitmentTerm = "Trial"
        $billingTerm = "Trial"
    }
    
    # Get display name
    $displayName = if ($skuLookup[$sub.skuPartNumber]) { 
        $skuLookup[$sub.skuPartNumber] 
    }
    else { 
        $sub.skuPartNumber 
    }
    
    # Get assigned license info
    $assignedInfo = $skuAssigned[$sub.skuId]
    
    $results += [PSCustomObject]@{
        'Subscription Name'  = $displayName
        'Status'             = $sub.status
        'Total Licenses'     = $sub.totalLicenses
        'Assigned Licenses'  = if ($assignedInfo) { $assignedInfo.Consumed } else { 0 }
        'Available Licenses' = if ($assignedInfo) { $assignedInfo.Available } else { $sub.totalLicenses }
        'Created Date'       = if ($createdDate) { $createdDate.ToString("yyyy-MM-dd") } else { "N/A" }
        'Renewal Date'       = if ($renewalDate) { $renewalDate.ToString("yyyy-MM-dd") } else { "N/A" }
        'Commitment'         = $commitmentTerm
        'Billing'            = $billingTerm
        'Is Trial'           = $sub.isTrial
        'Subscription ID'    = $sub.commerceSubscriptionId
        'SKU Part Number'    = $sub.skuPartNumber
        'SKU ID'             = $sub.skuId
    }
}

# Export to Excel
$results | Export-Excel -Path $Out `
    -AutoSize `
    -AutoFilter `
    -TableName "Subscriptions" `
    -TableStyle Medium2 `
    -FreezeTopRow `
    -WorksheetName "Subscriptions"

Log "Exported report: $Out" Green

$answer = Read-Host "Open the report now? [Y/n]"
if ($answer -eq '' -or $answer -match '^y') {
    Start-Process $Out
}