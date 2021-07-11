$computers = @(

       [pscustomobject]@{PSComputerName='Computer 1';Status='UP'}
       [pscustomobject]@{PSComputerName='Computer 2';Status='UP'}
       [pscustomobject]@{PSComputerName='Computer 3';Status='UP'}
   )

# Create an Excel object
$ExcelObj = New-Object -comobject Excel.Application
# $ExcelObj.Visible = $true
$ExcelObj.DisplayAlerts = $false
# Add a workbook
$ExcelWorkBook = $ExcelObj.Workbooks.Add()
$ExcelWorkSheet = $ExcelWorkBook.Worksheets.Item(1)
# Rename a worksheet
$ExcelWorkSheet.Name = 'Server Status'
# Fill in the head of the table
$ExcelWorkSheet.Cells.Item(1,1) = 'Server Name'
$ExcelWorkSheet.Cells.Item(1,2) = 'Server Status'
# Make the table head bold, set the font size and the column width
$ExcelWorkSheet.Rows.Item(1).Font.Bold = $true
$ExcelWorkSheet.Rows.Item(1).Font.size=15
$ExcelWorkSheet.Columns.Item(1).ColumnWidth=28
$ExcelWorkSheet.Columns.Item(2).ColumnWidth=28
$ExcelWorkSheet.Columns.Item(3).ColumnWidth=28
# Get the list of all Windows Servers in the domain
$counter=2
# Connect to each computer and get the service status
foreach ($computer in $computers) {
# Fill in Excel cells with the data obtained from the server
$ExcelWorkSheet.Columns.Item(1).Rows.Item($counter) = $computer.PSComputerName
$ExcelWorkSheet.Columns.Item(2).Rows.Item($counter) = $computer.Status
$counter++
}

$pwd = Get-Location
$date = Get-Date -Format "yyyyMMddHHmmss"
$fileName = "$($pwd)\Server_report_$date.xlsx" 
Write-Host $fileName

# Save the report and close Excel:
$ExcelWorkBook.SaveAs("$fileName")
$ExcelWorkBook.close($true)