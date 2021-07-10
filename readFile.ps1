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

$count = 0
Write-Progress -Activity "Gathering Information" -status "Pinging Hosts..." -percentComplete ($count / $servers.Count *100)
foreach ($server in $servers)
{
    $count++
    $ping_status = Test-Connection $server -Quiet
    Write-Host "$server, $ping_status"
    # if ($False -eq $ping_status) {
    #     Write-Host "$server, $ping_status"
    # }
    Write-Progress -Activity "Gathering Information" -status "Pinging Hosts..." -percentComplete ($count / $servers.Count *100)
}
Write-Progress -Activity "Gathering Information" -status "Pinging Hosts..." -Completed