<#
.SYNOPSIS
    Exports all Intune configuration policy settings with friendly names to an Excel report.

.DESCRIPTION
    Connects to Microsoft Graph and retrieves all Settings Catalog policies (Beta API) and
    Device Configuration policies (v1.0 API). Resolves friendly setting names and values using
    inline setting definitions, then outputs an Excel workbook with two sheets:
    - Settings: one row per setting across all policies with category, name, and value
    - Duplicates: settings that appear in more than one policy

.PARAMETER GraphCloud
    The Microsoft Graph national cloud environment to connect to.
    Defaults to 'Global'. Accepted values: Global, USGov, USGovDoD, Germany, China.

.PARAMETER Tenant
    Optional tenant ID or domain to target when authenticating to Microsoft Graph.
    If omitted, the default tenant for the authenticated account is used.

.PARAMETER Out
    Optional output path for the Excel file (.xlsx).
    Defaults to: ~\Downloads\<OrgName>-Settings Export-<timestamp>.xlsx

.EXAMPLE
    .\Export-ConfigurationSettings.ps1

    Connects to the global cloud and saves the report to the Downloads folder.

.EXAMPLE
    .\Export-ConfigurationSettings.ps1 -Tenant contoso.onmicrosoft.com -Out C:\Reports\settings.xlsx

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
    'DeviceManagementConfiguration.Read.All',
    'Organization.Read.All'
)
Log "Connecting to Graph..."
if ($Tenant) { Connect-MgGraph -Scopes $scopes -Environment $GraphCloud -NoWelcome -ErrorAction Stop -TenantId $Tenant }
else { Connect-MgGraph -Scopes $scopes -Environment $GraphCloud -NoWelcome -ErrorAction Stop }
$Org = Get-MgOrganization
Log "Connected to Graph: $($Org.DisplayName)" Green

if (!$Out) {
    $friendlyName = "SettingsExport"
    $Out = Join-Path $env:USERPROFILE "Downloads\$($Org.DisplayName)-$friendlyName-$(Get-Date -Format 'yyyy-MM-dd_HH-mm').xlsx"
}
elseif ($Out -notlike '*.xlsx') { $Out += '.xlsx' }

# --- Helper: Graph pagination ---
function Get-AllGraphPages {
    param([string]$Uri)
    $results = @()
    $currentUri = $Uri
    while ($currentUri) {
        $response = Invoke-MgGraphRequest -Method GET -Uri $currentUri
        $results += $response.value
        $currentUri = $response.'@odata.nextLink'
    }
    return $results
}

# --- Helper: Convert camelCase to friendly name ---
function ConvertTo-FriendlyName {
    param([string]$Name)
    $spaced = $Name -creplace '([a-z])([A-Z])', '$1 $2'
    $spaced = $spaced -creplace '([A-Z]+)([A-Z][a-z])', '$1 $2'
    return $spaced.Substring(0, 1).ToUpper() + $spaced.Substring(1)
}

# --- Helper: Fallback name from definition ID ---
function Get-FallbackName {
    param([string]$DefinitionId)
    $segments = $DefinitionId -split '_'
    $last = $segments[-1]
    return ConvertTo-FriendlyName $last
}

# --- Helper: Resolve friendly value for a choice setting ---
function Resolve-ChoiceValue {
    param($RawValue, $Definition)
    if ($Definition -and $Definition.options) {
        foreach ($opt in $Definition.options) {
            if ($opt.itemId -eq $RawValue -or $opt.name -eq $RawValue) {
                return $opt.displayName
            }
        }
    }
    # Fallback: last segment
    $segments = $RawValue -split '_'
    return ConvertTo-FriendlyName $segments[-1]
}

# --- Helper: Format simple value ---
function Format-SimpleValue {
    param($Value, $Definition)
    if ($null -eq $Value) { return '(not set)' }
    # Try option matching if definition has options
    if ($Definition -and $Definition.options) {
        foreach ($opt in $Definition.options) {
            if ($opt.itemId -eq $Value -or $opt.name -eq "$Value" -or $opt.itemId -match "_$Value$") {
                return $opt.displayName
            }
        }
    }
    if ($Value -is [bool]) { return if ($Value) { 'Enabled' } else { 'Disabled' } }
    if ($Value -is [array]) { return ($Value -join ', ') }
    return [string]$Value
}

# --- Helper: Resolve category path ---
function Resolve-CategoryPath {
    param([string]$CategoryId, [hashtable]$CategoryMap)
    if (-not $CategoryId -or -not $CategoryMap.ContainsKey($CategoryId)) {
        return @{ Category = ''; Subcategory = '' }
    }
    $chain = @()
    $current = $CategoryMap[$CategoryId]
    $visited = @{}
    while ($current) {
        if ($visited.ContainsKey($current.id)) { break }
        $visited[$current.id] = $true
        $chain += $current.displayName
        if (-not $current.parentCategoryId -or $current.parentCategoryId -eq $current.id) { break }
        if ($CategoryMap.ContainsKey($current.parentCategoryId)) {
            $current = $CategoryMap[$current.parentCategoryId]
        }
        else { break }
    }
    # chain is leaf-to-root; reverse for root-to-leaf
    [array]::Reverse($chain)
    $category = $chain[0]
    $subcategory = if ($chain.Count -gt 1) { ($chain[1..($chain.Count - 1)] -join ' > ') } else { '' }
    return @{ Category = $category; Subcategory = $subcategory }
}

# --- Helper: Find definition by ID in definitions array ---
function Find-Definition {
    param([string]$DefinitionId, [array]$Definitions)
    foreach ($d in $Definitions) {
        if ($d.id -eq $DefinitionId) { return $d }
    }
    return $null
}

# --- Helper: Recursively process setting instances ---
function Expand-SettingInstance {
    param(
        $Instance,
        [array]$Definitions,
        [string]$PolicyName,
        [string]$Category,
        [string]$Subcategory,
        [hashtable]$CategoryMap
    )

    $rows = @()
    if (-not $Instance) { return $rows }

    $odataType = $Instance.'@odata.type'
    $defId = $Instance.settingDefinitionId
    $def = Find-Definition -DefinitionId $defId -Definitions $Definitions

    $settingName = if ($def) { $def.displayName } else { Get-FallbackName $defId }

    # Resolve category from definition if not already set
    if (-not $Category -and $def -and $def.categoryId) {
        $catPath = Resolve-CategoryPath -CategoryId $def.categoryId -CategoryMap $CategoryMap
        $Category = $catPath.Category
        $Subcategory = $catPath.Subcategory
    }

    switch -Wildcard ($odataType) {
        '*choiceSettingInstance' {
            $rawValue = $Instance.choiceSettingValue.value
            $friendlyValue = Resolve-ChoiceValue -RawValue $rawValue -Definition $def
            $rows += [PSCustomObject]@{
                'Policy Name'   = $PolicyName
                'Category'      = $Category
                'Subcategory'   = $Subcategory
                'Setting Name'  = $settingName
                'Setting Value' = $friendlyValue
            }
            # Process children
            if ($Instance.choiceSettingValue.children) {
                foreach ($child in $Instance.choiceSettingValue.children) {
                    $rows += Expand-SettingInstance -Instance $child -Definitions $Definitions -PolicyName $PolicyName -Category $Category -Subcategory $Subcategory -CategoryMap $CategoryMap
                }
            }
        }
        '*choiceSettingCollectionInstance' {
            $values = @()
            $childRows = @()
            foreach ($item in $Instance.choiceSettingCollectionValue) {
                $rawValue = $item.value
                $values += Resolve-ChoiceValue -RawValue $rawValue -Definition $def
                if ($item.children) {
                    foreach ($child in $item.children) {
                        $childRows += Expand-SettingInstance -Instance $child -Definitions $Definitions -PolicyName $PolicyName -Category $Category -Subcategory $Subcategory -CategoryMap $CategoryMap
                    }
                }
            }
            $rows += [PSCustomObject]@{
                'Policy Name'   = $PolicyName
                'Category'      = $Category
                'Subcategory'   = $Subcategory
                'Setting Name'  = $settingName
                'Setting Value' = ($values -join ', ')
            }
            $rows += $childRows
        }
        '*simpleSettingInstance' {
            $rawValue = $Instance.simpleSettingValue.value
            $friendlyValue = Format-SimpleValue -Value $rawValue -Definition $def
            $rows += [PSCustomObject]@{
                'Policy Name'   = $PolicyName
                'Category'      = $Category
                'Subcategory'   = $Subcategory
                'Setting Name'  = $settingName
                'Setting Value' = $friendlyValue
            }
        }
        '*simpleSettingCollectionInstance' {
            $values = @()
            foreach ($item in $Instance.simpleSettingCollectionValue) {
                $values += [string]$item.value
            }
            $rows += [PSCustomObject]@{
                'Policy Name'   = $PolicyName
                'Category'      = $Category
                'Subcategory'   = $Subcategory
                'Setting Name'  = $settingName
                'Setting Value' = ($values -join ', ')
            }
        }
        '*groupSettingCollectionInstance' {
            $groupCount = 0
            if ($Instance.groupSettingCollectionValue) { $groupCount = $Instance.groupSettingCollectionValue.Count }
            $rows += [PSCustomObject]@{
                'Policy Name'   = $PolicyName
                'Category'      = $Category
                'Subcategory'   = $Subcategory
                'Setting Name'  = $settingName
                'Setting Value' = "$groupCount rule(s) configured"
            }
            foreach ($group in $Instance.groupSettingCollectionValue) {
                if ($group.children) {
                    foreach ($child in $group.children) {
                        $rows += Expand-SettingInstance -Instance $child -Definitions $Definitions -PolicyName $PolicyName -Category $Category -Subcategory $Subcategory -CategoryMap $CategoryMap
                    }
                }
            }
        }
        default {
            # Unknown type, output raw
            $rows += [PSCustomObject]@{
                'Policy Name'   = $PolicyName
                'Category'      = $Category
                'Subcategory'   = $Subcategory
                'Setting Name'  = $settingName
                'Setting Value' = $Instance.value
            }
        }
    }

    return $rows
}

# =====================================================================
# STEP 1: Fetch Settings Catalog Policies
# =====================================================================
Log "Fetching Settings Catalog policies..."
$catalogPolicies = Get-AllGraphPages -Uri "beta/deviceManagement/configurationPolicies?`$select=id,name,description,platforms,templateReference"
Log "Found $($catalogPolicies.Count) Settings Catalog policies" Green

# =====================================================================
# STEP 2: Fetch Categories
# =====================================================================
Log "Fetching configuration categories..."
$categories = Get-AllGraphPages -Uri "beta/deviceManagement/configurationCategories?`$select=id,displayName,description,parentCategoryId,rootCategoryId"
$categoryMap = @{}
foreach ($cat in $categories) {
    $categoryMap[$cat.id] = $cat
}
Log "Loaded $($categories.Count) categories" Green

# =====================================================================
# STEP 3: Fetch settings for each Settings Catalog policy (batched)
# =====================================================================
Log "Fetching settings for Settings Catalog policies..."
$settingsData = @()

# Process in batches of 20
for ($i = 0; $i -lt $catalogPolicies.Count; $i += 20) {
    $batch = $catalogPolicies[$i..[Math]::Min($i + 19, $catalogPolicies.Count - 1)]
    $batchRequests = @()
    for ($j = 0; $j -lt $batch.Count; $j++) {
        $policy = $batch[$j]
        $batchRequests += @{
            id     = "${j}_$($policy.id)"
            method = 'GET'
            url    = "/deviceManagement/configurationPolicies/$($policy.id)/settings?`$expand=settingDefinitions"
        }
    }

    $batchBody = @{ requests = $batchRequests } | ConvertTo-Json -Depth 10
    $batchResponse = Invoke-MgGraphRequest -Method POST -Uri "beta/`$batch" -Body $batchBody -ContentType 'application/json'

    foreach ($resp in $batchResponse.responses) {
        $policyId = ($resp.id -split '_', 2)[1]
        $policy = $catalogPolicies | Where-Object { $_.id -eq $policyId }
        $policyName = $policy.name

        if ($resp.status -ne 200) {
            Log "  Warning: Failed to fetch settings for '$policyName' (HTTP $($resp.status))" Yellow
            continue
        }

        # Handle pagination within batch response
        $allSettings = @()
        $allSettings += $resp.body.value
        $nextLink = $resp.body.'@odata.nextLink'
        while ($nextLink) {
            $pageResponse = Invoke-MgGraphRequest -Method GET -Uri $nextLink
            $allSettings += $pageResponse.value
            $nextLink = $pageResponse.'@odata.nextLink'
        }

        foreach ($setting in $allSettings) {
            $definitions = $setting.settingDefinitions
            $instance = $setting.settingInstance

            $settingsData += Expand-SettingInstance `
                -Instance $instance `
                -Definitions $definitions `
                -PolicyName $policyName `
                -Category '' `
                -Subcategory '' `
                -CategoryMap $categoryMap
        }
    }

    $processed = [Math]::Min($i + 20, $catalogPolicies.Count)
    Log "  Processed $processed / $($catalogPolicies.Count) policies..."
}

# =====================================================================
# STEP 4: Fetch Device Configuration Policies
# =====================================================================
Log "Fetching Device Configuration policies..."
$deviceConfigs = Get-AllGraphPages -Uri "v1.0/deviceManagement/deviceConfigurations"
Log "Found $($deviceConfigs.Count) Device Configuration policies" Green

$excludedProperties = @('id', 'displayName', 'description', 'createdDateTime', 'lastModifiedDateTime', 'version', 'omaSettings', 'payload')

foreach ($config in $deviceConfigs) {
    $policyName = $config.displayName

    # Process OMA-URI settings if present
    if ($config.omaSettings) {
        foreach ($oma in $config.omaSettings) {
            $omaType = $oma.'@odata.type'
            $value = switch ($omaType) {
                '#microsoft.graph.omaSettingString'        { $oma.value }
                '#microsoft.graph.omaSettingInteger'       { [string]$oma.value }
                '#microsoft.graph.omaSettingFloatingPoint' { [string]$oma.value }
                '#microsoft.graph.omaSettingBoolean'       { if ($oma.value) { 'True' } else { 'False' } }
                '#microsoft.graph.omaSettingDateTime'      { [string]$oma.value }
                '#microsoft.graph.omaSettingBase64'        {
                    if ($oma.fileName) { "[Binary file: $($oma.fileName)]" } else { '[Binary data]' }
                }
                '#microsoft.graph.omaSettingStringXml' {
                    try { [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($oma.value)) }
                    catch { $oma.value }
                }
                default { [string]$oma.value }
            }
            $settingsData += [PSCustomObject]@{
                'Policy Name'   = $policyName
                'Category'      = ''
                'Subcategory'   = ''
                'Setting Name'  = $oma.displayName
                'Setting Value' = $value
            }
        }
    }

    # Process built-in key/value settings
    $props = $config.Keys | Where-Object {
        $_ -notin $excludedProperties -and
        $_ -notlike '@odata*' -and
        $null -ne $config[$_]
    }

    foreach ($prop in $props) {
        $rawValue = $config[$prop]
        $value = if ($null -eq $rawValue) { '(not set)' }
        elseif ($rawValue -is [bool]) { if ($rawValue) { 'Enabled' } else { 'Disabled' } }
        elseif ($rawValue -is [array]) { ($rawValue | ForEach-Object { [string]$_ }) -join ', ' }
        elseif ($rawValue -is [hashtable] -or $rawValue -is [System.Collections.IDictionary]) { $rawValue | ConvertTo-Json -Compress }
        else { [string]$rawValue }

        $settingsData += [PSCustomObject]@{
            'Policy Name'   = $policyName
            'Category'      = ''
            'Subcategory'   = ''
            'Setting Name'  = ConvertTo-FriendlyName $prop
            'Setting Value' = $value
        }
    }
}

Log "Total settings collected: $($settingsData.Count)" Green

# =====================================================================
# STEP 5: Detect cross-policy duplicates
# =====================================================================
Log "Detecting duplicate settings across policies..."
$dupMap = @{}
foreach ($row in $settingsData) {
    $key = $row.'Setting Name'
    if (-not $dupMap.ContainsKey($key)) { $dupMap[$key] = @{} }
    $pn = $row.'Policy Name'
    if (-not $dupMap[$key].ContainsKey($pn)) { $dupMap[$key][$pn] = @() }
    $dupMap[$key][$pn] += $row.'Setting Value'
}

$duplicatesData = @()
foreach ($settingName in $dupMap.Keys | Sort-Object) {
    $policies = $dupMap[$settingName]
    if ($policies.Count -lt 2) { continue }
    $policyNames = @()
    $policyValues = @()
    foreach ($pn in $policies.Keys | Sort-Object) {
        $policyNames += $pn
        $policyValues += ($policies[$pn] -join '; ')
    }
    $duplicatesData += [PSCustomObject]@{
        'Setting Name' = $settingName
        'Policies'     = $policyNames -join '; '
        'Values'       = $policyValues -join '; '
    }
}

Log "Found $($duplicatesData.Count) duplicate settings" Green

# =====================================================================
# STEP 6: Export to Excel
# =====================================================================
Log "Exporting to Excel..."
$excelParams = @{
    Path         = $Out
    AutoSize     = $true
    FreezeTopRow = $true
    TableStyle   = 'Medium2'
}

$settingsData | Export-Excel @excelParams -WorksheetName 'Settings' -TableName 'SettingsTable' -PassThru | ForEach-Object {
    $ws = $_.Workbook.Worksheets['Settings']
    $valueCol = $ws.Dimension.Columns  # Setting Value is the last column
    foreach ($row in 2..$ws.Dimension.Rows) {
        $ws.Cells[$row, $valueCol].Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Left
    }
    $_.Save(); $_.Dispose()
}

if ($duplicatesData.Count -gt 0) {
    $duplicatesData | Export-Excel @excelParams -WorksheetName 'Duplicates' -TableName 'DuplicatesTable'
}
else {
    # Write an empty sheet with headers
    [PSCustomObject]@{ 'Setting Name' = ''; 'Policies' = ''; 'Values' = '' } |
        Export-Excel @excelParams -WorksheetName 'Duplicates' -TableName 'DuplicatesTable'
}

Log "Exported report: $Out" Green

$answer = Read-Host "Open the report now? [Y/n]"
if ($answer -eq '' -or $answer -match '^y') {
    Start-Process $Out
}
