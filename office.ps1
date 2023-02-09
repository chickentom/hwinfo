$office=Get-WmiObject -Class Win32_Product -ComputerName CZC607B3GY  -filter "name like '%Office%'"
$office=$office.Name -join ', '
$officefinal=$office.Split(",")
write-host $office
$office=$officefinal[0]
write-host $office