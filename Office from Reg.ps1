$HKEY_LOCAL_MACHINE = 2147483650
$strKeyPath = "Software\RegisteredApplications"

$temp = Get-WmiObject -List -Namespace "root\default"  -ComputerName "CZC607B3GY" `  | Where-Object {$_.Name -eq "StdRegProv"}
$res=$temp.EnumValues($HKEY_LOCAL_MACHINE, $strKeyPath)
foreach ($subkey in ($res.sNames))
{
write-host $subKey
if ($subKey -match "Excel.Application"){
     $OfficeVersion = ($subKey.Replace("Excel.Application.","")+".0")
     write-host "Office installed:" + $OfficeVersion  
     #[System.Windows.Forms.MessageBox]::Show($OfficeVersion,"Titel",0)
    }
}