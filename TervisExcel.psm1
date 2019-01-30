function Export-TervisExcelCustomerReports {
    param (
        [Parameter(Mandatory)][ref]$Workbook,
        [Parameter(Mandatory)]$ExportDirectory
    )
    
    $Date = Get-Date -Format "yyyy-MM-dd HH.mm"
    $WorksheetNamesToExport = "By Channel", "ITC (By Sales Rep)", "By Channel (Qtrly)", "TOP SUMMARY SHEET"
    $Workbook.Value.Sheets | Where-Object "Name" -In $WorksheetNamesToExport | ForEach-Object {
        $ExportPath = Join-Path -Path $ExportDirectory -ChildPath "$Date $($_.Name).pdf"
        $_.ExportAsFixedFormat(0, $ExportPath)
    }
}

function Invoke-TervisTopCustomerReportExport {
    param (
        [Parameter(Mandatory)]$ExcelFilePath,
        [Parameter(Mandatory)]$ExportDirectory,
        [Parameter(Mandatory)]$WriteResPassword
    )
    $Excel = New-ExcelInstance
    $Workbook = Open-ExcelFile -ExcelInstance ([ref]$Excel) -ExcelFilePath $ExcelFilePath -WriteResPassword $WriteResPassword -IgnoreReadOnlyRecommended $true
    Update-ExcelFile -Workbook ([ref]$Workbook)
    Export-TervisExcelCustomerReports -Workbook ([ref]$Workbook) -ExportDirectory $ExportDirectory
    Stop-ExcelInstance -ExcelInstance ([ref]$Excel) -SaveBeforeQuit
}
