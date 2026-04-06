#Requires -Modules PnP.PowerShell, ImportExcel

<#
.SYNOPSIS
    Exports SharePoint site recycle bin contents to an Excel report.

.DESCRIPTION
    Connects to a SharePoint site via PnP PowerShell and retrieves all recycle bin items,
    exporting them to an Excel workbook with Summary and Details worksheets.
    Supports optional filtering by deletion date range and/or the user who deleted the item.
    Date/time values are converted to local time for display. The report is saved to the
    user's Downloads folder by default.

.PARAMETER Url
    The SharePoint site URL to connect to. Must start with http:// or https://.

.PARAMETER AppId
    The Entra ID app registration (client) ID used for authentication.

.PARAMETER Tenant
    The tenant domain (e.g. contoso.onmicrosoft.com).

.PARAMETER Cloud
    The Azure cloud environment to connect to. Defaults to 'Production'.
    Valid values: Production, PPE, China, Germany, USGovernment, USGovernmentHigh, USGovernmentDoD.

.PARAMETER Out
    Optional output file path for the Excel report. If omitted, the report is saved to
    the user's Downloads folder with a timestamped filename. A .xlsx extension is appended
    if not provided.

.PARAMETER DeletedBy
    Optional filter by the email address/UPN of the user who deleted the item.

.PARAMETER Start
    Optional start date/time filter. Only items deleted on or after this date are included.
    Accepts PowerShell datetime strings (e.g. "2025-10-30 18:00", "2025-10-30", "2025-10-30T18:00").
    Can be used individually or combined with -End to define a range.

.PARAMETER End
    Optional end date/time filter. Only items deleted on or before this date are included.
    If a date-only value is provided, it is treated as end-of-day for inclusive comparison.
    Accepts PowerShell datetime strings. Can be used individually or combined with -Start.

.EXAMPLE
    .\Export-RecycleBin.ps1 -Url "https://contoso.sharepoint.com/sites/HR" -AppId "00000000-0000-0000-0000-000000000000" -Tenant "contoso.onmicrosoft.com"

.EXAMPLE
    .\Export-RecycleBin.ps1 -Url "https://contoso.sharepoint.com" -AppId "00000000-0000-0000-0000-000000000000" -Tenant "contoso.onmicrosoft.com" -Start "2025-10-01" -End "2025-10-31" -DeletedBy "user@contoso.com"
#>

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

    [string]$Out,
    [string]$DeletedBy,
    [datetime]$Start,
    [datetime]$End
)

function Log { param([string]$Msg, [ConsoleColor]$Color = 'White') Write-Host "[$(Get-Date -Format 'HH:mm')] $Msg" -ForegroundColor $Color }

Log "Connecting to SharePoint Admin..."
Connect-PnPOnline -Url $Url -ClientId $AppId -Interactive -Tenant $Tenant -AzureEnvironment $Cloud -ErrorAction Stop

$tenantName = ($Tenant.Trim() -split '\.')[0]
Log "Connected to $tenantName PnP" Green

if (!$Out) {
    $friendlyName = "RecycleBin"
    $Out = Join-Path $env:USERPROFILE "Downloads\$tenantName-$friendlyName-$(Get-Date -Format 'yyyy-MM-dd_HH-mm').xlsx"
} elseif ($Out -notlike '*.xlsx') { $Out += '.xlsx' }

# Get current site server-relative URL so we can strip it from item paths
# Example: "/" (root site) OR "/sites/HR" (non-root)
$web = Get-PnPWeb -Includes ServerRelativeUrl, Title
$siteRel = $web.ServerRelativeUrl           # e.g. "/" or "/sites/SiteName"
$siteRelTrim = $siteRel.Trim('/')           # "" or "sites/SiteName"
$siteName = if ([string]::IsNullOrWhiteSpace($web.Title)) { "(root)" } else { $web.Title }

function Remove-SitePrefix {
    [CmdletBinding()]
    param([Parameter()] [string]$Path)
    if ([string]::IsNullOrEmpty($Path)) { return $Path }
    $p = ($Path -replace '\\', '/').TrimStart('/')      # normalize slashes and leading "/"
    if ([string]::IsNullOrEmpty($siteRelTrim)) { return $p }   # root site: nothing to strip
    if ($p -eq $siteRelTrim) { return "" }
    if ($p.StartsWith("$siteRelTrim/")) { return $p.Substring($siteRelTrim.Length + 1) }
    return $p
}

function Get-LocalDeletedDate([datetime]$utcDate) {
    if ($utcDate.Kind -eq [System.DateTimeKind]::Utc) {
        return $utcDate.ToLocalTime()
    }
    else {
        return [datetime]::SpecifyKind($utcDate, [System.DateTimeKind]::Utc).ToLocalTime()
    }
}

# Normalize date-only End to end-of-day for inclusive comparisons
if ($PSBoundParameters.ContainsKey('End') -and $End.TimeOfDay -eq [TimeSpan]::Zero) {
    $End = $End.Date.AddDays(1).AddMilliseconds(-1)
}


# Get Recycle Bin Items
$RecycleItemsRaw = Get-PnPRecycleBinItem | Where-Object {
    $localDeleted = Get-LocalDeletedDate $_.DeletedDate

    $dateMatch =
    if ($PSBoundParameters.ContainsKey('Start') -and $PSBoundParameters.ContainsKey('End')) {
        $localDeleted -ge $Start -and $localDeleted -le $End
    }
    elseif ($PSBoundParameters.ContainsKey('Start')) {
        $localDeleted -ge $Start
    }
    elseif ($PSBoundParameters.ContainsKey('End')) {
        $localDeleted -le $End
    }
    else {
        $true
    }

    $deletedByMatch =
    if ($PSBoundParameters.ContainsKey('DeletedBy') -and $DeletedBy) {
        $_.DeletedByEmail -eq $DeletedBy
    }
    else {
        $true
    }

    $dateMatch -and $deletedByMatch
} | Select-Object DirName, LeafName, DeletedByEmail, DeletedDate, DeletedDateLocalFormatted, ItemType, AuthorEmail, ItemState, Id


# Build DeletedItem without the site prefix, add Order + Site + RecycleId
$RecycleItems = $RecycleItemsRaw | ForEach-Object {
    $dir = Remove-SitePrefix -Path $_.DirName
    $leaf = ($_.LeafName -replace '\\', '/')

    $deletedItem = if ([string]::IsNullOrEmpty($dir)) { $leaf }
    elseif ($dir -eq $leaf) { $dir }
    elseif ($dir -match '/$') { "$dir$leaf" }
    else { "$dir/$leaf" }

    # Order = path depth - 1 (min 1). Same depth => same Order value (by design).
    $depth = ($deletedItem -split '/').Count
    $order = [Math]::Max(1, $depth - 1)

    [pscustomobject]@{
        DeletedItem               = $deletedItem                 # no sites/SiteName prefix here
        Order                     = $order
        RecycleId                 = $_.Id
        DeletedByEmail            = $_.DeletedByEmail
        DeletedDateLocalFormatted = $_.DeletedDateLocalFormatted
        ItemType                  = $_.ItemType
        AuthorEmail               = $_.AuthorEmail
        ItemState                 = $_.ItemState
        Site                      = $siteName
    }
}

# Summary: IsDeleted = Yes/No; ChildCount excludes the top-level folder itself
$ChildCounts = @{}
$TopDeleted = @{}

$RecycleItems | ForEach-Object {
    $parts = $_.DeletedItem -split '/'
    # top = first two segments of path (post-strip), or the single segment if there's only one
    $top = if ($parts.Count -gt 1) { "$($parts[0])/$($parts[1])" } else { $_.DeletedItem }

    if (-not $ChildCounts.ContainsKey($top)) { $ChildCounts[$top] = 0 }
    if (-not $TopDeleted.ContainsKey($top)) { $TopDeleted[$top] = $false }

    if ($_.DeletedItem -eq $top -and $_.ItemType -eq 'Folder') {
        $TopDeleted[$top] = $true
    }
    elseif ($_.DeletedItem -like ($top + '/*')) {
        $ChildCounts[$top]++
    }
}

$AllTops = ($ChildCounts.Keys + $TopDeleted.Keys) | Sort-Object -Unique
$Summary = $AllTops | ForEach-Object {
    $folder = $_
    [pscustomobject]@{
        Site       = $siteName
        Folder     = $folder
        IsDeleted  = if ($TopDeleted[$folder]) { 'Yes' } else { 'No' }
        ChildCount = $ChildCounts[$folder]
    }
}

# Export to Excel
$Summary      | Export-Excel -Path $Out -WorksheetName "Summary" -AutoSize -FreezeTopRow -TableName "Summary" -TableStyle Medium2
$RecycleItems | Export-Excel -Path $Out -WorksheetName "Details" -AutoSize -FreezeTopRow -TableName "Details" -TableStyle Medium2

Log "Exported report: $Out" Green

$answer = Read-Host "Open the report now? [Y/n]"
if ($answer -eq '' -or $answer -match '^y') {
    Start-Process $Out
}