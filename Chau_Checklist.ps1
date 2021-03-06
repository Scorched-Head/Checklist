﻿$server = Read-Host -Prompt "Nombre del server"
$xl=New-Object -ComObject "Excel.Application"
$wb=$xl.Workbooks.Open($env:USERPROFILE+"\Desktop\CheckLists_V2.xlsx")
$ws=$wb.ActiveSheet
$xl.Visible=$True
$cells=$ws.Cells

function Get-InfoNIC {
    $nics = Get-WmiObject -class Win32_NetworkAdapterConfiguration -ComputerName $server -Filter "IPEnabled=True"
    foreach ($nic in $nics) {
        
        $index = $nic.Index

        if ($nic.DefaultIPGateway) {
          $gateway = $nic.DefaultIPGateway[0]
        }
        else {
          $gateway = "Not Set"
        }
        
        $props = [ordered]@{'IP'=$nic.IPAddress[0]                   
                   'Subnet'=$nic.IPSubnet[0] 
                   'Default Gateway'= $gateway 
                   'Nombre'=(Get-WmiObject -class Win32_NetworkAdapter -ComputerName $server -Filter "DeviceId = $index" | 
                                   select -ExpandProperty NetConnectionID)}
        New-Object -TypeName PSObject -Property $props
    }
}

function Get-DiskInfo {

GWMI Win32_LogicalDisk -Computername $server -Filter "DriveType=3"|
select SystemName, Name, VolumeName, FreeSpace, BlockSize, Size| 
% {$_.BlockSize=(($_.FreeSpace)/($_.Size))*100;
   $_.FreeSpace=($_.FreeSpace/1GB);$_}|
Select @{n='Server';e={$_.SystemName}},
       @{n='Unidad de Disco';e={$_.Name}},
       @{n='Nombre del Volumen' ;e={$_.VolumeName}},
       @{n='Tamaño (GB)';e={$_.Size / 1GB -as [int]}},
       @{n='Espacio Libre' ;e={'{0:N2}'-f $_.FreeSpace}},
       @{n='% Libre';e={'{0:N2}'-f $_.BlockSize}}
 } 
            

#DATOS DEL SERVER
			
$cells.item(5,3)=($server)
$cells.item(7,3)=(get-Date)
$cells.item(4,3)=(GWMI Win32_OperatingSystem -ComputerName $server | % {$_.Caption})

#PARAMETROS

$cells.item(13,6)=(GWMI –Class Win32_Processor -Computername $server | measure -Property NumberOfLogicalProcessors -sum).sum,' CPUs' -join "`r"
$cells.item(14,6)=(Get-WmiObject -Class win32_operatingsystem -Computername $server | % {[math]::round($_.TotalVisibleMemorySize /1MB)}),' GB' -join "`r"
$cells.item(15,6)=(GWMI Win32_LogicalDisk -Computername $server -Filter "DeviceID = 'C:'" | % {$_.DeviceID})
$cells.item(16,6)=(GWMI Win32_LogicalDisk -Computername $server -Filter "DriveType=3" | Where {$_.DeviceID -eq 'E:'} | % {$_.DeviceID})
$cells.item(17,6)=(GWMI Win32_LogicalDisk -Computername $server -Filter "DriveType=3" | Where {$_.DeviceID -eq 'F:'} | % {$_.DeviceID})
$cells.item(18,6)=(GWMI Win32_LogicalDisk -Computername $server -Filter "DeviceID = 'D:'" | % {$_.DeviceID})
#$cells.item(19,6)=(Get-InfoNIC | % {$_.IP})
$cells.item(24,6)=(GWMI -Class CIM_DataFile -ComputerName $server -Filter "Name='C:\\Program Files\\VMware\\VMware Tools\\vmtoolsd.exe'" | % {$_.Version})
$cells.item(25,6)=(GWMI -Class CIM_DataFile -ComputerName $server -Filter "Name='C:\\Program Files (x86)\\Symantec\\Symantec Endpoint Protection\\Smc.exe'" | % {$_.Version})
$cells.item(26,6)=(Invoke-Command -ComputerName $server -ScriptBlock {Get-ItemProperty -Path "HKLM:\SOFTWARE\Symantec\Symantec Endpoint Protection\CurrentVersion\public-opstate"} | % {$_.LatestVirusDefsDate})
$cells.item(27,6)=(Get-Service -ComputerName $server -Name SepMasterService | % {$_.Status}), '' -join "`r"
$cells.item(28,6)=(GWMI -Class CIM_DataFile -ComputerName $server -Filter "Name='C:\\Program Files\\Altiris\\Altiris Agent\\AeXNSAgent.exe'" | % {$_.Version})
$cells.item(29,6)=(Get-Service -ComputerName $server -Name AeXNSClient | % {$_.Status}), '' -join "`r"

#ALMACENAMIENTO

$cells.item(33,5)=(GWMI Win32_LogicalDisk -Computername $server -Filter "DeviceID = 'C:'" | % {$_.BlockSize=(($_.FreeSpace)/($_.Size))*100;
   $_.FreeSpace=($_.FreeSpace/1GB);$_} | % {$_.BlockSize}),' % Libre' -join "`r"
$cells.item(34,5)=(GWMI Win32_LogicalDisk -Computername $server -Filter "DeviceID = 'D:'" | % {$_.BlockSize=(($_.FreeSpace)/($_.Size))*100;
   $_.FreeSpace=($_.FreeSpace/1GB);$_} | % {$_.BlockSize}),' % Libre' -join "`r"
$cells.item(35,5)=(GWMI Win32_LogicalDisk -Computername $server -Filter "DriveType=3" | Where {$_.DeviceID -eq 'E:'} | % {$_.BlockSize=(($_.FreeSpace)/($_.Size))*100;
   $_.FreeSpace=($_.FreeSpace/1GB);$_} | % {$_.BlockSize}),' % Libre' -join "`r"
$cells.item(36,5)=(GWMI Win32_LogicalDisk -Computername $server -Filter "DriveType=3" | Where {$_.DeviceID -eq 'F:'} | % {$_.BlockSize=(($_.FreeSpace)/($_.Size))*100;
   $_.FreeSpace=($_.FreeSpace/1GB);$_} | % {$_.BlockSize}),' % Libre' -join "`r"
   
#DESEMPEÑO
 
$cells.item(39,6)=(GWMI –class Win32_processor -Computername $server | measure -Property LoadPercentage -Average).Average,' % Utilizado' -join "`r"
$cells.item(40,6)=(Get-WmiObject -Class win32_operatingsystem -Computername $server| 
% {[math]::round(($_.FreePhysicalMemory)/($_.TotalVisibleMemorySize)*100,2)}),' % Libre' -join "`r"

#SEGURIDAD

$cells.item(43,6)=(Get-Ciminstance -Class "Win32_UserAccount" -ComputerName $server | Where-Object {$_.Name -like '*localadmin*'} | % {$_.Name})

#BACKUP/RESTORE

$cells.item(48,6)=(GWMI -Class CIM_DataFile -ComputerName $server -Filter "Name='C:\\Program Files\\Veritas\\NetBackup\\bin\\NbWin.exe'" | % {$_.Version})

#OTROS

$cells.item(52,6)=(Invoke-Command -ComputerName $server -ScriptBlock {Get-ScheduledTask -TaskName Backup*}|% {$_.TaskName})


$wb.SaveAs($env:USERPROFILE+"\Desktop\$server.xlsx")
#$wb.Close()
#$xl.Quit()

Get-DiskInfo | FT
Get-InfoNIC