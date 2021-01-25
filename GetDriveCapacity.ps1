#checks a computer for how much free space is remaining on the C drive

Add-Type -AssemblyName Microsoft.VisualBasic
$comp = [Microsoft.VisualBasic.Interaction]::InputBox('Enter a computer name', 'Drive Capacity', "Enter a computer")

$disk = Get-WmiObject Win32_LogicalDisk -ComputerName $comp -Filter "DeviceID='C:'" |
Select-Object Size,FreeSpace

$final = [math]::Round($disk.FreeSpace / $disk.Size * 100)

$final = "$final% drive capacity available"

[System.Windows.MessageBox]::Show("$comp has $final")

