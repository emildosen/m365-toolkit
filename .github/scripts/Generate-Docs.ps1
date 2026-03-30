[CmdletBinding()]
param(
    [string]$RepoRoot = '.',
    [string]$DocsRoot = './docs'
)

$sections = @('entra', 'exchange', 'purview', 'sharepoint')

$skipParams = @(
    'Verbose', 'Debug', 'ErrorAction', 'WarningAction', 'InformationAction',
    'ErrorVariable', 'WarningVariable', 'InformationVariable', 'OutVariable',
    'OutBuffer', 'PipelineVariable', 'WhatIf', 'Confirm'
)

foreach ($section in $sections) {
    $sectionPath = Join-Path $RepoRoot $section
    $outputDir   = Join-Path $DocsRoot $section

    if (-not (Test-Path $sectionPath)) {
        Write-Warning "Section folder not found: $sectionPath"
        continue
    }

    $null = New-Item -ItemType Directory -Path $outputDir -Force

    $scripts = Get-ChildItem -Path $sectionPath -Filter '*.ps1' | Sort-Object Name

    foreach ($script in $scripts) {
        Write-Host "Processing $($script.Name)..."

        $help = Get-Help $script.FullName -Full -ErrorAction SilentlyContinue
        $name = $script.BaseName
        $lines = [System.Collections.Generic.List[string]]::new()

        $lines.Add("# $name")
        $lines.Add('')

        # Synopsis
        if ($help.Synopsis -and $help.Synopsis.Trim()) {
            $lines.Add($help.Synopsis.Trim())
            $lines.Add('')
        }

        # Description
        if ($help.Description) {
            $desc = ($help.Description | ForEach-Object { $_.Text }) -join "`n`n"
            if ($desc.Trim()) {
                $lines.Add('## Description')
                $lines.Add('')
                $lines.Add($desc.Trim())
                $lines.Add('')
            }
        }

        # Parameters
        $params = $help.Parameters.Parameter | Where-Object { $_.Name -notin $skipParams }
        if ($params) {
            $lines.Add('## Parameters')
            $lines.Add('')
            foreach ($param in $params) {
                $lines.Add("### -$($param.Name)")
                $lines.Add('')
                if ($param.Description) {
                    $paramDesc = ($param.Description | ForEach-Object { $_.Text }) -join ' '
                    $lines.Add($paramDesc.Trim())
                    $lines.Add('')
                }
                $lines.Add('| Property | Value |')
                $lines.Add('|----------|-------|')
                if ($param.Type.Name) {
                    $lines.Add("| Type | ``$($param.Type.Name)`` |")
                }
                $lines.Add("| Required | $($param.Required) |")
                if ($param.DefaultValue -and $param.DefaultValue -notin @('none', 'None', '')) {
                    $lines.Add("| Default | $($param.DefaultValue) |")
                }
                if ($param.ValidValues) {
                    $validStr = ($param.ValidValues | ForEach-Object { "``$_``" }) -join ', '
                    $lines.Add("| Valid values | $validStr |")
                }
                $lines.Add('')
            }
        }

        # Examples
        if ($help.Examples.Example) {
            $lines.Add('## Examples')
            $lines.Add('')
            $i = 1
            foreach ($example in $help.Examples.Example) {
                $title = if ($example.Title) {
                    ($example.Title -replace '^[\s\-]+|[\s\-]+$', '').Trim()
                } else {
                    "Example $i"
                }
                if (-not $title) { $title = "Example $i" }

                $lines.Add("### $title")
                $lines.Add('')
                if ($example.Code) {
                    $lines.Add('```powershell')
                    $lines.Add($example.Code.Trim())
                    $lines.Add('```')
                    $lines.Add('')
                }
                if ($example.Remarks) {
                    $remarks = ($example.Remarks |
                        Where-Object { $_.Text.Trim() } |
                        ForEach-Object { $_.Text.Trim() }) -join "`n"
                    if ($remarks.Trim()) {
                        $lines.Add($remarks.Trim())
                        $lines.Add('')
                    }
                }
                $i++
            }
        }

        # Full script source
        $rawUrl = "https://raw.githubusercontent.com/emildosen/m365-toolkit/main/$section/$($script.Name)"
        $lines.Add('## Script')
        $lines.Add('')
        $lines.Add("[Download $($script.Name)]($rawUrl){ .md-button }")
        $lines.Add('')
        $lines.Add('```powershell')
        $lines.Add((Get-Content $script.FullName -Raw).TrimEnd())
        $lines.Add('```')
        $lines.Add('')

        $outFile = Join-Path $outputDir "$name.md"
        $lines -join "`n" | Set-Content $outFile -Encoding UTF8 -NoNewline
        Write-Host "  -> $outFile"
    }
}

Write-Host "Doc generation complete."
