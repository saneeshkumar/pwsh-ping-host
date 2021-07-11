Function Get-ServersFile {
    Param(
    [string]$file
    )
    Process
    {
        $read = New-Object System.IO.StreamReader($file)
        $serverarray = @()

        while ( $null -ne ($line = $read.ReadLine()))
        {
            $serverarray += $line
        }

        $read.Dispose()
        return $serverarray
    }
}

$servers = Get-ServersFile -file "pcName.txt"
Write-Host $servers.Count

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

$count = 0
$counter=2
Write-Progress -Activity "Gathering Information" -status "Pinging Hosts..." -percentComplete ($count / $servers.Count *100)
foreach ($server in $servers)
{
    $count++
    $ping_status = Test-Connection $server -Quiet
    Write-Host "$server, $ping_status"
    # if ($False -eq $ping_status) {
    #     Write-Host "$server, $ping_status"
    # }
    # Fill in Excel cells with the data obtained from the server
    $ExcelWorkSheet.Columns.Item(1).Rows.Item($counter) = $server
    $ExcelWorkSheet.Columns.Item(2).Rows.Item($counter) = $ping_status
    $counter++
    Write-Progress -Activity "Gathering Information" -status "Pinging Hosts..." -percentComplete ($count / $servers.Count *100)
}
Write-Progress -Activity "Gathering Information" -status "Pinging Hosts..." -Completed

# Save the report and close Excel:
$ExcelWorkBook.SaveAs('D:\Sample_Applications\powershell\Server_report.xlsx')
$ExcelWorkBook.close($true)