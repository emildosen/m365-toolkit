<#
.SYNOPSIS
    Finds empty folders within a SharePoint document library or folder path.

.DESCRIPTION
    Connects to a SharePoint site via PnP PowerShell and recursively walks the entire
    folder tree under the specified folder path to identify folders that contain no
    files and no sub-folders (leaf-level empty folders). All nesting levels are scanned.
    Results are exported to an Excel report showing the full relative path of each
    empty folder. The report is saved to the user's Downloads folder by default.

.PARAMETER Url
    The SharePoint site URL to connect to. Must start with http:// or https://.

.PARAMETER AppId
    The Entra ID app registration (client) ID used for authentication.

.PARAMETER Tenant
    The tenant domain (e.g. contoso.onmicrosoft.com).

.PARAMETER Folder
    The folder path relative to the site URL to scan (e.g. "Shared Documents" or
    "Shared Documents/Archive"). This should be the server-relative path segment
    after the site URL.

.PARAMETER Cloud
    The Azure cloud environment to connect to. Defaults to 'Production'.
    Valid values: Production, PPE, China, Germany, USGovernment, USGovernmentHigh, USGovernmentDoD.

.PARAMETER Out
    Optional output file path for the Excel report. If omitted, the report is saved to
    the user's Downloads folder with a timestamped filename. A .xlsx extension is appended
    if not provided.

.EXAMPLE
    .\Find-EmptyFolders.ps1 -Url "https://contoso.sharepoint.com/sites/HR" -AppId "00000000-0000-0000-0000-000000000000" -Tenant "contoso.onmicrosoft.com" -Folder "Shared Documents"

.EXAMPLE
    .\Find-EmptyFolders.ps1 -Url "https://contoso.sharepoint.com/sites/HR" -AppId "00000000-0000-0000-0000-000000000000" -Tenant "contoso.onmicrosoft.com" -Folder "Shared Documents/Archive" -Out "C:\Reports\empty-folders.xlsx"
#>
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

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$Folder,

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
    $friendlyName = "EmptyFolders"
    $Out = Join-Path $env:USERPROFILE "Downloads\$tenantName-$friendlyName-$(Get-Date -Format 'yyyy-MM-dd_HH-mm').xlsx"
} elseif ($Out -notlike '*.xlsx') { $Out += '.xlsx' }

$web = Get-PnPWeb -Includes ServerRelativeUrl
$siteRelUrl = $web.ServerRelativeUrl.TrimEnd('/')

$emptyFolders = [System.Collections.Generic.List[object]]::new()
$checked = 0

function Find-EmptyFoldersRecursive {
    param([string]$RelativeUrl)

    $script:checked++
    if ($script:checked % 50 -eq 0) { Log "Checked $($script:checked) folders so far..." }

    $files = Get-PnPFolderItem -FolderSiteRelativeUrl $RelativeUrl -ItemType File -ErrorAction SilentlyContinue
    $subFolders = Get-PnPFolderItem -FolderSiteRelativeUrl $RelativeUrl -ItemType Folder -ErrorAction SilentlyContinue

    if ((-not $files -or $files.Count -eq 0) -and (-not $subFolders -or $subFolders.Count -eq 0)) {
        $emptyFolders.Add([pscustomobject]@{
            FolderPath = $RelativeUrl
            Name       = ($RelativeUrl -split '/')[-1]
        })
        return
    }

    foreach ($sub in $subFolders) {
        $subRelUrl = $sub.ServerRelativeUrl.Substring($siteRelUrl.Length + 1)
        Find-EmptyFoldersRecursive -RelativeUrl $subRelUrl
    }
}

Log "Scanning all folders under: $Folder"
Find-EmptyFoldersRecursive -RelativeUrl $Folder

Log "Found $($emptyFolders.Count) empty folders ($checked folders checked)"

if ($emptyFolders.Count -gt 0) {
    $emptyFolders | Export-Excel -Path $Out -WorksheetName "EmptyFolders" -AutoSize -FreezeTopRow -TableName "EmptyFolders" -TableStyle Medium2
    Log "Exported report: $Out" Green

    $answer = Read-Host "Open the report now? [Y/n]"
    if ($answer -eq '' -or $answer -match '^y') {
        Start-Process $Out
    }
} else {
    Log "No empty folders found. No report generated." Green
}
