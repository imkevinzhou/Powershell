<###########################################################################################################
# File name  : ParseFileImportExcel.ps1
# Version : V1.0
#
#############################################################################################################>

$color =  45126   #RGB(0,176.80)=0x00B050=45126
$newdevice = Get-Content -Encoding UTF8 .\NewDevice.txt
$updatedevice = Get-Content -Encoding UTF8 .\UpdateDevice.csv

$hash = ConvertFrom-StringData((Get-Content .\config.ini) -join "`n")

$xl=New-Object -ComObject excel.application
$xl.visible=$false
$wb=$xl.workbooks.add()
$sheet=$wb.sheets.item(1)

$sheet.cells.item(1,2)=$hash.FirmwareVersion
$sheet.cells.item(2,1)="Acroview"
$sheet.cells.item(2,2)="Programmer Type"
$sheet.cells.item(2,3)="Itmes"
#$sheet.cells.item(2,4)="Manufacture"
#$sheet.cells.item(2,5)="Chip Name"
#$sheet.cells.item(2,6)="Package"
#$sheet.cells.item(2,7)="Adapter"

$sheet.cells.item(3,2)="AP8000"
$sheet.cells.item(3,3)="New Devices"


for($i=0; $i -lt $newdevice.Count; $i++){
    $arr = $newdevice[$i] -split ","
    $sheet.cells.item($i+2,4)=$arr[0]
    $sheet.cells.item($i+2,5)=$arr[1]
    $sheet.cells.item($i+2,6)=$arr[2]
    $sheet.cells.item($i+2,7)=$arr[3]
}

$sheet.cells.item($newdevice.Count+2,3)="Update Devices"

for($i=0; $i -lt $updatedevice.Count; $i++){
    $arr = $updatedevice[$i] -split ","
    $sheet.cells.item($i+2+$newdevice.Count,4)=$arr[0]
    $sheet.cells.item($i+2+$newdevice.Count,5)=$arr[1]
    $sheet.cells.item($i+2+$newdevice.Count,6)=$arr[2]
    $sheet.cells.item($i+2+$newdevice.Count,7)=$arr[3]
}

$sheet.cells.item($newdevice.Count+2+$updatedevice.Count,3)="GUI modifcations"
$sheet.cells.item($newdevice.Count+2+$updatedevice.Count,4)=$hash.GUImodifcations

$sheet.cells.item($newdevice.Count+2+$updatedevice.Count+1,3)="Firmware modifications"
$sheet.cells.item($newdevice.Count+2+$updatedevice.Count+1,4)=$hash.Firmwaremodifications

$sheet.cells.item($newdevice.Count+2+$updatedevice.Count+2,3)="MultiAprog modifications"
$sheet.cells.item($newdevice.Count+2+$updatedevice.Count+2,4)=$hash.MultiAprogmodifications


$y = $newdevice.Count+$updatedevice.Count+4


$sheet.columns.item("A:G").EntireColumn.AutoFit() | out-null

      
$sheet.Range("B3:B$($newdevice.Count+$updatedevice.Count+1)").Merge() = $true
$sheet.Range("C3:C$($newdevice.Count+1)").Merge() = $true
$sheet.Range("C$($newdevice.Count+2):C$($newdevice.Count+$updatedevice.Count+1)").Merge() = $true

$sheet.Range("B$($newdevice.Count+$updatedevice.Count+2):B$($newdevice.Count+$updatedevice.Count+4)").Merge() = $true
$sheet.Range("D$($newdevice.Count+$updatedevice.Count+2):G$($newdevice.Count+$updatedevice.Count+2)").Merge() = $true
$sheet.Range("D$($newdevice.Count+$updatedevice.Count+3):G$($newdevice.Count+$updatedevice.Count+3)").Merge() = $true
$sheet.Range("D$($newdevice.Count+$updatedevice.Count+4):G$($newdevice.Count+$updatedevice.Count+4)").Merge() = $true


$sheet.Range("A1:G1").Interior.Color = $color  
$sheet.Range("A1:A$y").Interior.Color = $color 

$sheet.columns.item("A:G").EntireColumn.Font.Name = "Arial"
$sheet.columns.item("A:G").EntireColumn.Font.Size = 10

$sheet.Range("A1:G1").Font.Bold = $true 
$sheet.Range("A2:G1").Font.Bold = $true 


$xl.visible=$true