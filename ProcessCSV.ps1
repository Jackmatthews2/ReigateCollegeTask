$csv = Import-Csv "C:\Users\JackNEW\Documents\_Jack\_spreadsheets\Budgetting\oct18\oct18.csv"

# Not working yet
$csv | select *,@{name="type";expression={
        $payment = $_.payment; ? {$_.payment -like "*apple*"}; $payment
    }
}

<# Notes
$wmiServices = Get-CimInstance -ClassName win32_service -Property Name,PathName

Get-Service | Select-Object -Property *,
@{Name = 'PathName'; Expression = { $serviceName = $_.Name; (@($wmiServices).where({ $_.Name -eq $serviceName })).PathName }}
#>