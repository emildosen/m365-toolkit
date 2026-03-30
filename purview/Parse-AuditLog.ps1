<#
.SYNOPSIS
    Parses a Microsoft Purview audit log CSV export into a formatted Excel report.

.DESCRIPTION
    Imports a CSV exported from the Microsoft Purview compliance portal, expands the
    nested AuditData JSON column into individual columns, converts timestamps to both
    UTC and the system's local time zone, sorts results newest-first, and writes two
    worksheets to an Excel workbook: a Filtered sheet (noise columns removed) and a
    Detailed sheet (all columns).

.PARAMETER CsvPath
    Path to the audit log CSV file exported from the Microsoft Purview compliance portal.
    The file must exist and be a valid CSV containing an AuditData column.

.PARAMETER Out
    Optional path for the output Excel file. Defaults to
    ParsedAuditLog-<timestamp>.xlsx in the current user's Downloads folder. The .xlsx
    extension is appended automatically if omitted.

.EXAMPLE
    .\Parse-AuditLog.ps1 -CsvPath .\audit_export.csv

.EXAMPLE
    .\Parse-AuditLog.ps1 -CsvPath .\audit_export.csv -Out C:\Reports\audit_2026.xlsx
#>
#Requires -Modules ImportExcel

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateScript({
            if (-not (Test-Path $_ -PathType Leaf)) {
                throw "The specified CSV path '$_' does not exist or is not a file."
            }
            $true
        })]
    [string]$CsvPath,
    [string]$Out
)

function Log { param([string]$Msg, [ConsoleColor]$Color = 'White') Write-Host "[$(Get-Date -Format 'HH:mm')] $Msg" -ForegroundColor $Color }

if (!$Out) {
    $friendlyName = "ParsedAuditLog"
    $Out = Join-Path $env:USERPROFILE "Downloads\$friendlyName-$(Get-Date -Format 'yyyy-MM-dd_HH-mm').xlsx"
}
elseif ($Out -notlike '*.xlsx') { $Out += '.xlsx' }

# Columns to skip in the filtered sheet
$SkipColumns = @(
    'RecordId',
    'RecordType',
    'AssociatedAdminUnits',
    'AssociatedAdminUnitsNames',
    'AppAccessContext.APIId',
    'AppAccessContext.AuthTime',
    'AppAccessContext.ClientAppId',
    'AppAccessContext.ClientAppName',
    'AppAccessContext.CorrelationId',
    'AppAccessContext.TokenIssuedAtTime',
    'AppAccessContext.UniqueTokenId',
    'AppAccessContext.UserObjectId',
    'ApplicationId',
    'AssertingApplicationId',
    'AuthenticationType',
    'BrowserName',
    'BrowserVersion',
    'CorrelationId',
    'CreationTime',
    'CrossScopeSyncDelete',
    'DeviceDisplayName',
    'DoNotDistributeEvent',
    'EventData',
    'EventSignature',
    'FileSizeBytes',
    'FileSyncBytesCommitted',
    'GeoLocation',
    'HighPriorityMediaProcessing',
    'Id',
    'ImplicitShare',
    'ListBaseType',
    'ListId',
    'ListItemUniqueId',
    'ListServerTemplate',
    'MachineId',
    'ObjectId',
    'OrganizationId',
    'Site',
    'TargetUserOrGroupName',
    'TargetUserOrGroupType',
    'UniqueSharingId',
    'UserAgent',
    'UserKey',
    'UserType',
    'Version',
    'WebId',
    'Workload'
)


function ConvertFrom-AuditJSON {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return $null }
    try { return ($Text | ConvertFrom-Json -Depth 100 -ErrorAction Stop) } catch {
        $fixed = $Text.Trim('"') -replace '""', '"'
        try { return ($fixed | ConvertFrom-Json -Depth 100 -ErrorAction Stop) } catch { return $null }
    }
}

function Expand-AuditJSON {
    param(
        [Parameter(Mandatory)] [object]$InputObject,
        [string]$Prefix = ''
    )
    $out = [ordered]@{}

    $add = {
        param($k, $v)
        if ($k) { $out[$k] = $v }
    }

    if ($null -eq $InputObject) {
        & $add $Prefix $null
        return $out
    }

    if ($InputObject -is [pscustomobject] -or $InputObject -is [System.Collections.IDictionary]) {
        foreach ($p in $InputObject.PSObject.Properties) {
            $key = if ($Prefix) { "$Prefix.$($p.Name)" } else { $p.Name }
            $child = Expand-AuditJSON -InputObject $p.Value -Prefix $key
            foreach ($kv in $child.GetEnumerator()) { $out[$kv.Key] = $kv.Value }
        }
        return $out
    }

    if ($InputObject -is [System.Collections.IEnumerable] -and -not ($InputObject -is [string])) {
        $items = foreach ($i in $InputObject) {
            if ($i -is [pscustomobject] -or $i -is [System.Collections.IDictionary]) {
                $pairs = foreach ($pp in $i.PSObject.Properties) { "$($pp.Name)=$($pp.Value)" }
                ($pairs -join ',')
            }
            else {
                "$i"
            }
        }
        & $add $Prefix ($items -join '; ')
        return $out
    }

    & $add $Prefix $InputObject
    return $out
}

$rows = Import-Csv -Path $CsvPath
if (-not $rows) { throw "CSV is empty: $CsvPath" }
Log "Parsing $($rows.Count) audit log results..."

$baseCols = $rows[0].PSObject.Properties.Name | Where-Object { $_ -ne 'AuditData' }

$jsonKeySet = [System.Collections.Generic.HashSet[string]]::new()

foreach ($r in $rows) {
    $j = ConvertFrom-AuditJSON -Text $r.AuditData
    if ($j) {
        $flat = Expand-AuditJSON -InputObject $j
        foreach ($k in $flat.Keys) { [void]$jsonKeySet.Add($k) }
    }
}

$jsonKeys = $jsonKeySet | Sort-Object -Unique

$final = New-Object System.Collections.Generic.List[object]

foreach ($r in $rows) {
    $obj = [ordered]@{}
    foreach ($c in $baseCols) { $obj[$c] = $r.$c }

    $j = ConvertFrom-AuditJSON -Text $r.AuditData
    $flat = if ($j) { Expand-AuditJSON -InputObject $j } else { @{} }

    foreach ($k in $jsonKeys) {
        $obj[$k] = if ($flat.Contains($k)) { $flat[$k] } else { $null }
    }

    # Format date and time, showing both UTC and local time
    if ($obj.Contains('CreationDate')) {
        $val = $obj['CreationDate']
        $localTz = [System.TimeZoneInfo]::Local
        $localLabel = "Date ($($localTz.Id))"
        if ($val) {
            try {
                $dto = [datetimeoffset]::Parse($val, [System.Globalization.CultureInfo]::InvariantCulture,
                    [System.Globalization.DateTimeStyles]::AssumeUniversal -bor [System.Globalization.DateTimeStyles]::AdjustToUniversal)

                $utc = $dto.UtcDateTime
                $local = [System.TimeZoneInfo]::ConvertTimeFromUtc($utc, $localTz)

                $obj['Date (UTC)'] = $utc.ToString('yyyy-MM-dd HH:mm')
                $obj[$localLabel] = $local.ToString('yyyy-MM-dd HH:mm')
            }
            catch {
                $obj['Date (UTC)'] = $val
                $obj[$localLabel] = $null
            }
        }
        else {
            $obj['Date (UTC)'] = $null
            $obj[$localLabel] = $null
        }

        $obj.Remove('CreationDate')

        # Move Date (UTC) and local date to the start
        $newObj = [ordered]@{
            'Date (UTC)' = $obj['Date (UTC)']
            $localLabel  = $obj[$localLabel]
        }
        foreach ($k in $obj.Keys | Where-Object { $_ -notin @('Date (UTC)', $localLabel) }) {
            $newObj[$k] = $obj[$k]
        }
        $obj = $newObj
    }

    $final.Add([pscustomobject]$obj) | Out-Null

    if (($final.Count % 500) -eq 0) {
        Log "Processed $($final.Count) rows..."
    }
}

# Sort by date, newest first
$sorted = $final
if ($final.Count -gt 0 -and $final[0].PSObject.Properties.Name -contains 'Date (UTC)') {
    $fmt = 'yyyy-MM-dd HH:mm'
    $ci = [System.Globalization.CultureInfo]::InvariantCulture
    $sorted = $final | Sort-Object {
        if ($_.'Date (UTC)') {
            [datetime]::ParseExact($_.'Date (UTC)', $fmt, $ci)
        }
        else {
            [datetime]::MinValue
        }
    } -Descending
}

# Build the filtered sheet, skipping the columns defined at the top of the script.
$allCols = $sorted[0].PSObject.Properties.Name
$keepCols = $allCols | Where-Object { $_ -notin $SkipColumns }

# Export to Excel
$excelParams = @{
    Path         = $Out
    TableStyle   = "Medium2"
    AutoSize     = $true
    FreezeTopRow = $true
}

$sorted | Select-Object -Property $keepCols |
Export-Excel -WorksheetName 'Filtered' -TableName 'AuditFiltered' @excelParams
$sorted | Export-Excel -WorksheetName 'Detailed' -TableName 'AuditDetailed' @excelParams

Log "Exported report: $Out" Green

$answer = Read-Host "Open the report now? [Y/n]"
if ($answer -eq '' -or $answer -match '^y') {
    Start-Process $Out
}