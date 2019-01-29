function Export-TervisExcelCustomerReports {
    param (
        [Parameter(Mandatory)]$Workbook,
        [Parameter(Mandatory)]$ExportDirectory
    )
    
    $Date = Get-Date -Format "yyyy-MM-dd"
    $WorksheetNamesToExport = "By Channel", "ITC (By Sales Rep)", "By Channel (Qtrly)", "TOP SUMMARY SHEET"
    $Workbook.Sheets | Where-Object "Name" -In $WorksheetNamesToExport | ForEach-Object {
        $ExportPath = Join-Path -Path $ExportDirectory -ChildPath "$Date $($_.Name).pdf"
        $_.ExportAsFixedFormat(0, $ExportPath)
    }
}

function Invoke-TervisTopCustomerReportExport {
    param (
        [Parameter(Mandatory)]$ReportDirectory,
        [Parameter(Mandatory)]$ExcelFilePath,
        [Parameter(Mandatory)]$ExportDirectory,
        [Parameter(Mandatory)]$WriteResPassword
    )
    $Excel = New-ExcelInstance
    $Workbook = Open-ExcelFile -ExcelInstance $Excel -ExcelFilePath $ExcelFilePath -WriteResPassword $WriteResPassword -IgnoreReadOnlyRecommended $true
    Update-ExcelFile -Workbook $Workbook
    Export-TervisExcelCustomerReports -Workbook $Workbook -ExportDirectory $ExportDirectory
    Stop-ExcelInstance -ExcelInstance $Excel -SaveBeforeQuit
}
