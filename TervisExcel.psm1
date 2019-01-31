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

function Invoke-TervisTopCustomerReportExtract {
    $ReportCredential = Get-PasswordstatePassword -ID 5699
    $ReportParameters = $ReportCredential.GenericField1 | ConvertFrom-Json

    $Excel = New-ExcelInstance
    $Workbook = Open-ExcelFile -ExcelInstance ([ref]$Excel) -ExcelFilePath $ReportParameters.ExcelFilePath -WriteResPassword $ReportCredential.Password -IgnoreReadOnlyRecommended $true
    Update-ExcelFile -Workbook ([ref]$Workbook)
    Export-TervisExcelCustomerReports -Workbook ([ref]$Workbook) -ExportDirectory $ReportParameters.ExportDirectory
    Stop-ExcelInstance -ExcelInstance ([ref]$Excel) -SaveBeforeQuit
}

function Invoke-ExcelTaskApplicationProvision {
    $ApplicationName = "ExcelTask"
    $EnvironmentName = "Infrastructure"
    Invoke-ApplicationProvision -ApplicationName $ApplicationName -EnvironmentName $EnvironmentName
    $Nodes = Get-TervisApplicationNode -ApplicationName $ApplicationName -EnvironmentName $EnvironmentName
    $Nodes | Push-TervisPowershellModulesToRemoteComputer
    $Nodes | ForEach-Object {Invoke-Command -ComputerName $_.ComputerName -ScriptBlock {Add-LocalGroupMember -Group Administrators -Member "Privilege_InfrastructureScheduledTasksAdministrator"}}
    $Credential = Get-PasswordstatePassword -ID 5574 -AsCredential
    # Need to change this action to trigger the Polaris instance instead
    $Action = New-ScheduledTaskAction -Execute "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe" -Argument "-Command Invoke-TervisTopCustomerReportExtract"
    $ModuleName = "TervisTopCustomerReportExtract"
    $Port = 9000
    $ServiceName = "$ModuleName $Port"
    $Nodes | Install-TervisScheduledTask -TaskName "Top Customer Report Export" -Action $Action -RepetitionIntervalName "EveryDayAt730am" -Credential $Credential
    $Nodes | Install-PowerShellApplicationPolaris -EnvironmentName Infrastructure -ModuleName $ModuleName -TervisModuleDependencies TervisExcel,ExcelPowerShell,TervisPasswordstatePowershell,PasswordstatePowershell -Ports 9000 -CommandString Start-TervisTervisTopCustomerReportPolaris
    $Nodes | ForEach-Object { Invoke-Command -ComputerName $_.ComputerName -ScriptBlock { $C = $using:Credential; nssm set $using:ServiceName ObjectName $C.Username $C.GetNetworkCredential().Password } }
    # $Nodes | ForEach-Object { Invoke-Command -ComputerName $_.ComputerName -ScriptBlock { $using:ModuleName; nssm get $using:ServiceName ObjectName } }
    $Nodes | ForEach-Object { Invoke-Command -ComputerName $_.ComputerName -ScriptBlock { $using:ModuleName; nssm start $using:ServiceName } }
}

function Start-TervisTervisTopCustomerReportPolaris {
    param (
        [Parameter(Mandatory)]$Port
    )
    New-PolarisRoute -Path "/StartExtract" -Method "POST" -Scriptblock {
        $Response.Send("Starting Top Customer Report Extract")
        Invoke-TervisTopCustomerReportExtract
    } -Force

    New-PolarisRoute -Path "/*" -Method "GET" -Scriptblock {
        $Response.Send("TervisExcel Polaris server is up.")
    } -Force

    $Polaris = Start-Polaris -Https -Port $Port

    while ($Polaris.Listener.IsListening) {
        Wait-Event callbackeventbridge.callbackcomplete
    }
}


