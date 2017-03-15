
#find directory yourself
$Script:PSScriptRoot = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition



###определяет тип компьютера
function get-TypePC($PCname){
$PCmodel=""
$WMIinfo=gwmi win32_systemenclosure -ComputerName $PCname
switch($WMIinfo.chassistypes){
1 {$PCmodel="Другое";break}
2 {$PCmodel="Unknown";break}
3 {$PCmodel="Настольный ПК";break}
4 {$PCmodel="Low Profile Desktop";break}
5 {$PCmodel="Pizza Box";break}
6 {$PCmodel="Mini Tower";break}
7 {$PCmodel="Tower";break}
8 {$PCmodel="Portable";break}
9 {$PCmodel="Laptop";break}
10 {$PCmodel="Ноутбук";break}
11 {$PCmodel="Handheld";break}
12 {$PCmodel="Docking Station";break}
13 {$PCmodel="All-in-One";break}
14 {$PCmodel="Sub-Notebook";break}
15 {$PCmodel="Space Saving";break}
16 {$PCmodel="Lunch Box";break}
17 {$PCmodel="Main System Chassis";break}
18 {$PCmodel="Expansion Chassis";break}
19 {$PCmodel="Sub-Chassis";break}
20 {$PCmodel="Bus Expansion Chassis";break}
21 {$PCmodel="Peripheral Chassis";break}
22 {$PCmodel="Storage Chassis";break}
23 {$PCmodel="Rack Mount Chassis";break}
24 {$PCmodel="Sealed-Case PC";break}
}
return $PCmodel
}



###Заменяет неверные размерности хардов
function get-HHDRightSize($size){
switch($size){
75 {$size=80;break}
149 {$size=160;break}
233 {$size=240;break}
234 {$size=240;break}
298 {$size=320;break}
466 {$size=500;break}
596 {$size=650;break}
default {$size=$size;break}
}
return $size
}



###Получение актуального списа AD компьютеров
function get-ADList{
import-module activedirectory
$Object=Get-ADComputer -filter * -Properties name, LastLogon, operatingsystem|Select-Object name, operatingsystem -Skip 1|sort-object name
$Object|Add-Member -MemberType NoteProperty -Name OS_Architecture -Value '' -Force
$Object|Add-Member -MemberType NoteProperty -Name OS_SerialNumber -Value '' -Force
$Object|Add-Member -MemberType NoteProperty -Name HDD_size -Value '' -Force
$Object|Add-Member -MemberType NoteProperty -Name HDD_SerialNumber -Value '' -Force
$Object|Add-Member -MemberType NoteProperty -Name RAM -Value '' -Force
$Object|Add-Member -MemberType NoteProperty -Name RAM_SerialNumber -Value '' -Force
$Object|Add-Member -MemberType NoteProperty -Name CPU -Value '' -Force
$Object|Add-Member -MemberType NoteProperty -Name MB_name -Value '' -Force
$Object|Add-Member -MemberType NoteProperty -Name MB_Manufacturer -Value '' -Force
$Object|Add-Member -MemberType NoteProperty -Name MB_SerialNumber -Value '' -Force
$Object|Add-Member -MemberType NoteProperty -Name PCType -Value '' -Force
$Object|Add-Member -MemberType NoteProperty -Name printer -Value '' -Force
$Object|Add-Member -MemberType NoteProperty -Name Status -Value NONE -Force

return $Object
}



###Архивация
function start-Archive([int32]$CountArchive){

if((Test-Path "$PSScriptRoot\Archive") -eq $false){new-item "$PSScriptRoot\Archive" -Type directory}   #archive data directory

 $date=[string](Get-Date -Format "%d_%M_%y-%h_%m")
 cd ${env:ProgramFiles}
 .\7-Zip\7z.exe a -tzip -mx9 "$PSScriptRoot\Archive\$date.zip" "$PSScriptRoot\Data"

 #delete old archive
dir $PSScriptRoot\Archive\*.zip|Sort-Object LastWriteTime -Descending|Select-Object -Skip $CountArchive|foreach{del $_.fullname}
return
}



###Обновить лист
function get-UpdateList{
$OldList=Import-Clixml $PSScriptRoot\Data\adcomp.xml
$NewList=get-ADList

for($i=0;$i -lt $NewList.length;$i++)
    {
    for($j=0;$j -lt $OldList.length;$j++)
        {
        if($NewList[$i].name -like $OldList[$j].name)
            {
            $OldList[$j].Status="Ok"
            $NewList[$i]=$OldList[$j]
            if($j -eq 0){$OldList=$OldList[($j+1)..($OldList.length-1)]}
            else{$OldList=$OldList[0..($j-1)+($j+1)..($OldList.length-1)]}
            } 
        else{$OldList[$j].Status="NONE"}
        }
    }
 
$ADList=$NewList+$OldList
return $ADList
}


###Сравнение списков
function get-CompareList($NewList, $OldList){

for($i=0;$i -lt $NewList.length;$i++)
    {
    for($j=0;$j -lt $OldList.length;$j++)
        {
        if($NewList[$i] -like $OldList[$j])
            {
            $OldList[$j].Status="Ok"
            $NewList[$i]=$OldList[$j]
            if($j -eq 0){$OldList=$OldList[($j+1)..($OldList.length-1)]}
            else{$OldList=$OldList[0..($j-1)+($j+1)..($OldList.length-1)]}
            } 
        else{$OldList[$j].Status="NONE"}
        }
    }
 
}


###Новая инвентаризация
function get-NewInventory(){

if((Test-Path "$PSScriptRoot\Data") -eq $false){new-item "$PSScriptRoot\Data" -Type directory}

$ADFile=get-ADList

$ADFile|Export-Clixml $PSScriptRoot\Data\adcomp.xml

start-Inventory
}



###Основная инвентаризация
function start-Inventory(){

if((Test-Path "$PSScriptRoot\Data") -eq $false){get-NewInventory;break}
else
{
start-Archive 100 #Создаем архив, хранит не более 100 архивов
$ADList=get-UpdateList
$CompareList=get-UpdateList

$ADlist|where{$_.status -like "Ok"}|foreach{if((Test-Connection $_.name -Count 1 -Quiet) -eq $false) #if comp on -> set wmi parametrs
{
Write-Host $_.name "Выключен..." -ForegroundColor Red
$_.Status="Off"
}
else
{
Write-Host $_.name "Включен..." -ForegroundColor Green

$hdd=get-wmiobject -ComputerName $_.name -Class win32_diskdrive|where{$_.DeviceID -like "*PHYSICALDRIVE0"}|Select-Object SerialNumber, @{n="size"; e={[int32]($_.size/1048576/1024)}}
$os=Get-WmiObject -ComputerName $_.name -Class win32_operatingsystem|Select-Object OSArchitecture, serialnumber
$mb=Get-WmiObject -ComputerName $_.name -Class win32_baseboard|Select-Object product, serialnumber, Manufacturer
$cpu=Get-WmiObject -ComputerName $_.name -Class win32_processor|Select-Object name
$printer=Get-WmiObject -ComputerName $_.name -Class Win32_Printer|where{$_.name -notlike "*pdf*" -and $_.name -notlike "fax"-and $_.name -notlike "microsoft xps*"}|Select-Object name
$ram=Get-WmiObject -ComputerName $_.name -Class win32_physicalmemory|Select-Object serialnumber, @{n="size"; e={($_.capacity/1024/1024/1024)}}

$_.RAM='';$_.RAM+=$ram|foreach{[string]$_.size+"GB"}
$_.RAM_SerialNumber='';$_.RAM_SerialNumber+=$ram|foreach{$_.serialnumber+";"}
$_.OS_Architecture=$os.OSArchitecture
$_.OS_SerialNumber=$os.serialnumber
$_.Status="Ok"
$_.HDD_size= [string](get-HHDRightSize $hdd.size)+"GB"
$_.HDD_SerialNumber=$hdd.SerialNumber
$_.MB_name=$mb.product
$_.MB_Manufacturer=$mb.Manufacturer
$_.MB_SerialNumber=$mb.serialnumber
$_.printer='';$_.printer+=$printer.name|foreach{$_+";"}
$_.CPU=$cpu.name
$_.PCType=get-TypePC $_.name
}
}
}
$ADlist|foreach{if($_.HDD_SerialNumber -notlike $null -and $_.HDD_SerialNumber.Length -ge 30)
{$hdd_dex=$_.HDD_SerialNumber -split '(.{2})' |%{ if ($_ -ne ""){[CHAR]([CONVERT]::toint16("$_",16))}}
$_.HDD_SerialNumber='';$_.HDD_SerialNumber+=$hdd_dex|foreach{$_};$_.HDD_SerialNumber=$_.HDD_SerialNumber -replace " "}}


###функция для сравнения двух списков и выявления смены комплектующих


$ADlist=$ADlist|sort-object name
$ADlist|Export-Clixml $PSScriptRoot\Data\adcomp.xml
$ADlist|ConvertTo-Html > $PSScriptRoot\fullcomp.html
}



start-Inventory