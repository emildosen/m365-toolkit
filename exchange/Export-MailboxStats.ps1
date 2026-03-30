<#
.SYNOPSIS
    Exports mailbox size statistics for all user and shared mailboxes to an Excel report.

.DESCRIPTION
    Connects to Exchange Online and retrieves mailbox statistics for all UserMailbox and
    SharedMailbox recipients, including primary mailbox size, archive size, quota limits,
    and last logon time. Results are exported to a formatted Excel workbook.

.PARAMETER ExchangeCloud
    The Exchange Online environment to connect to. Defaults to 'O365Default'.
    Valid values: O365Default, O365USGovGCCHigh, O365USGovDoD, O365GermanyCloud, O365China.

.PARAMETER Tenant
    Optional. The tenant domain or organization name to connect to. Useful when the
    account has access to multiple tenants.

.PARAMETER Out
    Optional. Full path for the output .xlsx file. If omitted, the file is saved to the
    current user's Downloads folder with a timestamped filename. A .xlsx extension is
    appended automatically if not provided.

.EXAMPLE
    .\Export-MailboxStats.ps1

.EXAMPLE
    .\Export-MailboxStats.ps1 -Tenant contoso.onmicrosoft.com -Out C:\Reports\mailboxes.xlsx

.EXAMPLE
    .\Export-MailboxStats.ps1 -ExchangeCloud O365USGovGCCHigh
#>
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
    $friendlyName = "MailboxStats"
    $Out = Join-Path $env:USERPROFILE "Downloads\$($Org.DisplayName)-$friendlyName-$(Get-Date -Format 'yyyy-MM-dd_HH-mm').xlsx"
}
elseif ($Out -notlike '*.xlsx') { $Out += '.xlsx' }

function Convert-ToBytes {
    param($Size)
    if ($null -eq $Size) { return 0 }
    $s = $Size.ToString()
    $m = [regex]::Match($s, '\((\d+)\s+bytes\)')
    if ($m.Success) { return [int64]$m.Groups[1].Value }
    $m2 = [regex]::Match($s, '([0-9]*\.?[0-9]+)\s*(B|KB|MB|GB|TB|PB)', 'IgnoreCase')
    if ($m2.Success) {
        $val = [double]$m2.Groups[1].Value
        $unit = $m2.Groups[2].Value.ToUpperInvariant()
        $mult = switch ($unit) { 'B' { 1 } 'KB' { 1KB } 'MB' { 1MB } 'GB' { 1GB } 'TB' { 1TB } 'PB' { 1PB } }
        return [int64]($val * $mult)
    }
    0
}

function Convert-QuotaGB {
    param($QuotaObj)
    if ($null -eq $QuotaObj) { return $null }
    $s = $QuotaObj.ToString()
    if ($s -match 'Unlimited') { return $null }
    $bytes = Convert-ToBytes $QuotaObj
    if ($bytes -le 0) { return $null }
    [math]::Round($bytes / 1GB, 2)
}

function Test-HasArchive {
    param([string]$GuidString)
    try { return ([guid]$GuidString -ne [guid]::Empty) } catch { return $false }
}

try {
    Log "Fetching all mailboxes..."

    $mailboxes = Get-EXOMailbox -ResultSize Unlimited `
        -RecipientTypeDetails UserMailbox, SharedMailbox `
        -PropertySets Archive, Quota `
        -Properties DisplayName, RecipientTypeDetails, PrimarySmtpAddress

    Log "Processing $($mailboxes.Count) mailboxes..."

    $rows = foreach ($m in $mailboxes) {
        $identity = $m.PrimarySmtpAddress
        $statP = $null
        $statA = $null

        try { $statP = Get-MailboxStatistics -Identity $identity -ErrorAction Stop } catch {
            Log ("Primary stats failed for {0}: {1}" -f $identity, $_.Exception.Message) Yellow
        }

        $hasArchive = Test-HasArchive $m.ArchiveGuid

        if ($hasArchive) {
            try { $statA = Get-MailboxStatistics -Identity $identity -Archive -ErrorAction Stop } catch {
                Log ("Archive stats not available for {0}: {1}" -f $identity, $_.Exception.Message) Yellow
            }
        }

        $pBytes = if ($statP) { Convert-ToBytes $statP.TotalItemSize } else { 0 }
        $pDeleted = if ($statP) { Convert-ToBytes $statP.TotalDeletedItemSize } else { 0 }
        $aBytes = if ($statA) { Convert-ToBytes $statA.TotalItemSize } else { 0 }
        $aDeleted = if ($statA) { Convert-ToBytes $statA.TotalDeletedItemSize } else { 0 }

        [pscustomobject]@{
            UserPrincipalName    = $m.UserPrincipalName
            DisplayName          = $m.DisplayName
            MailboxType          = $m.RecipientTypeDetails
            HasArchive           = $hasArchive
            AutoExpandingArchive = [bool]$m.AutoExpandingArchiveEnabled
            MailboxSizeGB        = [math]::Round($pBytes / 1GB, 2)
            MailboxDeletedGB     = [math]::Round($pDeleted / 1GB, 2)
            MailboxItems         = if ($statP) { $statP.ItemCount } else { $null }
            ArchiveSizeGB        = [math]::Round($aBytes / 1GB, 2)
            ArchiveDeletedGB     = [math]::Round($aDeleted / 1GB, 2)
            ArchiveItems         = if ($statA) { $statA.ItemCount } else { 0 }
            SoftLimit            = Convert-QuotaGB $m.IssueWarningQuota
            HardLimit            = Convert-QuotaGB $m.ProhibitSendReceiveQuota
            LastLogonTime        = if ($statP) { $statP.LastLogonTime } else { $null }
            Database             = if ($statP) { $statP.Database } else { $null }
        }
    }

}
catch {
    Log "Encountered an error: $($_.Exception.Message)" Red
}


# Normal file export to Excel
$rows |
Sort-Object -Property PrimarySizeGB -Descending |
Export-Excel -Path $Out -WorksheetName 'MailboxStats' -TableName 'MailboxStats' -AutoSize -FreezeTopRow -TableStyle Medium2

Log "Exported report: $Out" Green

$answer = Read-Host "Open the report now? [Y/n]"
if ($answer -eq '' -or $answer -match '^y') {
    Start-Process $Out
}