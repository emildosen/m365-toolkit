## Various Excel import/export snippets

### Params
-AutoSize       Automatically size column width to fit content
-WorksheetName  Name the sheet
-TableName      Name the table
-TableStyle     Set the table design (Medium2 = light blue)
-FreezeTopRow   Freezes the table header
-PassThru       Returns the saved Excel document
-ClearSheet     Empties the sheet if the document exists
-Title
-TitleBold
-TitleSize 14

### Default export settings
```powershell
$excelParams = @{
    Path          = $Out
    WorksheetName = "Sheet"
    TableName     = "Table"
    TableStyle    = "Medium2"
    AutoSize      = $true
    FreezeTopRow  = $true
}

$productDetails | Export-Excel @excelParams
```

Or as a one-liner:
```powershell
$productDetails | Export-Excel -Path $Out -WorksheetName "Sheet" -TableName "Table" -TableStyle "Medium2" -AutoSize -FreezeTopRow
```


### Pivot
```powershell
$pivotData = @{
    Endpoint  = 'Count'
    BaseScore = 'Max'
}

$outputData | Export-Excel -Path $Out `
    -IncludePivotTable -PivotTableName "Applications" -PivotRows "Application" -PivotData $pivotData -PivotDataToColumn `
```


### Conditional formatting
```powershell
$formatting = $(
    New-ConditionalText -Text 'y' -Range 'A:A' -BackgroundColor LightGreen
)

$outputData | Export-Excel -Path $Out `
    -ConditionalFormat $formatting
```


### Set column width
```powershell
$excel = $outputData | Export-Excel -Path $Out -PassThru

$ws = $excel.Workbook.Worksheets["Policies"]
# Fixed width
$ws.Column(2).Width = 50  # Description
$ws.Column(4).Width = 15  # Platform
# Auto size + append to the width
$ws.Column(6).AutoFit()  # Modified
$ws.Column(6).Width += 3

Close-ExcelPackage $excel
```


### Column formatting
```powershell
$excel = $report | Export-Excel -Path $output -WorksheetName 'Report' -AutoSize -TableName 'TimesheetReport' -TableStyle "Medium2" -FreezeTopRow -PassThru
$ws = $excel.Workbook.Worksheets['Report']
Set-ExcelColumn -Worksheet $ws -Column 5 -NumberFormat '0.00%'
Close-ExcelPackage $excel
```