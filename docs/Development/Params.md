## Commonly used params

### CSV file input
```powershell
[Parameter(Mandatory = $true)]
[Alias('Csv', 'InputCsv', 'CsvFile')]
[ValidateScript({
        if (-not (Test-Path $_ -PathType Leaf)) {
            throw "The specified CSV path '$_' does not exist or is not a file."
        }
        $true
    })]
[string]$CsvPath
```

### CSV template generator
```powershell
[CmdletBinding(DefaultParameterSetName = 'Process')]
param(
  [Parameter(Mandatory = $true, ParameterSetName = 'Process')]
  [string]$ConfigPath,

  [Parameter(Mandatory = $true, ParameterSetName = 'Process')]
  [string]$CsvPath,

  [Parameter(ParameterSetName = 'Template')]
  [Alias('Template')]
  [switch]$CreateTemplate
)

# Create template
if ($PSCmdlet.ParameterSetName -eq 'Template') {
  if (!$Out) { $Out = "GuestInviteTemplate-$(Get-Date -Format 'yyyy-MM-dd_HH-mm').csv" }
  elseif ($Out -notlike '*.csv') { $Out += '.csv' }

  [PSCustomObject]@{
    GreetingName = 'Example User'
    DisplayName  = 'Example User (Contoso)'
    CompanyName  = 'Contoso Corporation'
    EmailAddress = 'example@contoso.com'
    Status       = ''
  } | Export-Csv -Path $Out -NoTypeInformation
  Log "Template created: $Out" Green
  return
}
```
