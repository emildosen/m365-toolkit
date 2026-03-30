<#
.SYNOPSIS
    Exports all email addresses and aliases from Exchange Online to an Excel report.

.DESCRIPTION
    Connects to Exchange Online and retrieves all mail-enabled recipients (excluding
    contacts and guest mail users), including Microsoft 365 Groups. Outputs a flat
    list of every primary SMTP address and alias, enriched with recipient
    and group details.

.PARAMETER ExchangeCloud
    The Exchange Online environment to connect to. Defaults to 'O365Default'.
    Accepted values: O365Default, O365USGovGCCHigh, O365USGovDoD, O365GermanyCloud, O365China

.PARAMETER Tenant
    Optional. The organization domain or tenant ID to pass to Connect-ExchangeOnline.
    Use this when running in a multi-tenant context.

.PARAMETER Out
    Optional. Output file path for the Excel report. Defaults to a timestamped file
    in the current user's Downloads folder, named after the Exchange organization.
    The .xlsx extension is appended if omitted.
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
    $friendlyName = "EmailAliases"
    $Out = Join-Path $env:USERPROFILE "Downloads\$($Org.DisplayName)-$friendlyName-$(Get-Date -Format 'yyyy-MM-dd_HH-mm').xlsx"
}
elseif ($Out -notlike '*.xlsx') { $Out += '.xlsx' }

# All mail-enabled recipients, excluding contacts and guest mail users
$recipients = Get-EXORecipient -ResultSize Unlimited -PropertySets All | Where-Object {
    $_.RecipientType -ne 'MailContact' -and
    $_.RecipientTypeDetails -ne 'GuestMailUser'
}
Log "Fetched $($recipients.Count) mailboxes..."

# Microsoft 365 Groups with extended properties
$unifiedGroups = Get-UnifiedGroup -ResultSize Unlimited -IncludeAllProperties
Log "Fetched $($unifiedGroups.Count) M365 groups..."

# Map M365 group info by primary SMTP (lowercased)
$groupMeta = @{}
foreach ($g in $unifiedGroups) {
    if ($g.PrimarySmtpAddress) {
        $key = $g.PrimarySmtpAddress.ToString().ToLower()
        $groupMeta[$key] = $g
    }
}

$rows = [System.Collections.Generic.List[object]]::new()
$typeMap = @{
    'PrimarySmtpAddress' = 'Primary'
    'EmailAddresses'     = 'Alias'
}

Log "Processing..."

foreach ($r in $recipients) {
    $seen = [System.Collections.Generic.HashSet[string]]::new()

    $primarySmtp = $null
    if ($r.PrimarySmtpAddress) {
        $primarySmtp = $r.PrimarySmtpAddress.ToString()
    }

    $primaryKey = $null
    if ($primarySmtp) {
        $primaryKey = $primarySmtp.ToLower()
    }

    $groupInfo = $null
    if ($primaryKey -and $groupMeta.ContainsKey($primaryKey)) {
        $groupInfo = $groupMeta[$primaryKey]
    }

    $groupType = $null
    switch ($r.RecipientTypeDetails) {
        'GroupMailbox' { $groupType = 'M365 Group' }
        'MailUniversalDistributionGroup' { $groupType = 'Distribution Group' }
        'MailUniversalSecurityGroup' { $groupType = 'Mail-enabled Security Group' }
        'DynamicDistributionGroup' { $groupType = 'Dynamic Distribution Group' }
        'RoomList' { $groupType = 'Room List' }
    }

    $mkRow = {
        param(
            [string]$Email,
            [string]$Source
        )

        if (-not $Email) { return }
        Log "Processing: $Email"

        $key = $Email.ToLower()
        if ($seen.Contains($key)) { return }
        [void]$seen.Add($key)

        $localPart = $null
        $domain = $null
        if ($Email -like '*@*') {
            $localPart = $Email.Split('@')[0]
            $domain = $Email.Split('@')[1]
        }

        $friendlyType = $typeMap[$Source]
        if (-not $friendlyType) {
            $friendlyType = $Source
        }

        $rows.Add([pscustomobject]@{
                EmailAddress         = $Email
                LocalPart            = $localPart
                Domain               = $domain
                Type                 = $friendlyType
                UPN                  = $r.WindowsLiveID
                DisplayName          = $r.DisplayName
                Mailbox              = $r.RecipientTypeDetails
                Hidden               = $r.HiddenFromAddressListsEnabled

                GroupType            = $groupType
                AllowExternalSenders = if ($groupInfo) { -not $groupInfo.RequireSenderAuthenticationEnabled } else { $null }
                CopyToMemberInbox    = if ($groupInfo) { $groupInfo.AutoSubscribeNewMembers } else { $null }
                AccessType           = if ($groupInfo) { $groupInfo.AccessType } else { $null }
            })
    }

    # Primary SMTP
    if ($primarySmtp) {
        & $mkRow -Email $primarySmtp -Source 'PrimarySmtpAddress'
    }

    # All SMTP proxy addresses (primary + aliases)
    if ($r.EmailAddresses) {
        foreach ($addr in $r.EmailAddresses) {
            $s = $addr.ToString()

            if ($s -like 'SMTP:*' -or $s -like 'smtp:*') {
                $smtp = $s.Substring(5)
                & $mkRow -Email $smtp -Source 'EmailAddresses'
            }
        }
    }
}

# Export to Excel
$rows |
Sort-Object EmailAddress, Mailbox, DisplayName |
Export-Excel -Path $Out -WorksheetName "Emails" -TableName "Emails" -TableStyle "Medium2" -AutoSize -FreezeTopRow

Log "Exported report: $Out" Green

$answer = Read-Host "Open the report now? [Y/n]"
if ($answer -eq '' -or $answer -match '^y') {
    Start-Process $Out
}