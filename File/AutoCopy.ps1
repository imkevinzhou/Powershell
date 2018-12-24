<###########################################################################################################
# File name  : AutoCopy.ps1
使用System.IO.FileSystemWatcher这个.NET对象，新建文件、删除文件、重命名文件等操作均会被监控或监视。
#############################################################################################################>

# 监控的源文件夹
$Source_Path="C:\Users\Kevin\Desktop\Template\Container_0"
# 接收文件的目标文件夹
$Destination_Path="C:\Users\Kevin\Desktop\Template\Container_1"
# 复制的文件类型
$FileSuffix=".ko",".chk",".xml",".extdata",".txt",".c"
# 复制的文件名
$FileName=$($args[0])
# 每次监控的间隔时间=1000ms
$timeout=1000

Write-Host ""  
Write-Host "-------------------------------------------------------------"   

if ($FileName -eq $null)
{
    Write-Host " Please append file name! For example: .\AutoCopy.ps1 FileName"
    exit
}

Write-Host "AutoCopy $FileName .... " 
<# @1 
# 创建文件系统监视对象
$FileSystemWatcher = New-Object System.IO.FileSystemWatcher $Source_Path
 


Write-Host "AutoCopy $FileName .... " 

while ($true) 
{
    # 监控文件夹内的所有变化
    $result = $FileSystemWatcher.WaitForChanged('all', $timeout)
    if ($result.TimedOut -eq $false)
    {
        $sum = $FileSuffix.Length   
        For($i=0; $i -lt $sum; $i++)
        {  
            $tFile=$Source_Path+'\'+$FileName+$($FileSuffix[$i])  	
            if (Test-Path -Path $tFile)
            {
                #Write-Host "Send $FileName$($FileSuffix[$i]) Pass"
                cp $tFile $Destination_Path
            }else
            {
                #Write-Host "$FileName$($FileSuffix[$i]) not exist"	
            }	
        }
    }
} 
#>

<# @2 #>
$tFile=$Source_Path+$FileName+$FileSuffix[0]
# 创建文件系统监视对象
$FileSystemWatcher = New-Object System.IO.FileInfo($tFile)

while ($true) 
{
    $upFlag=$FileSystemWatcher.LastWriteTime
    $FileSystemWatcher = New-Object System.IO.FileInfo($tFile)

    if ($upFlag -ne $FileSystemWatcher.LastWriteTime)
    { 
        For($i=0; $i -lt $FileSuffix.Length; $i++)
        {  
            $tFile=$Source_Path+$FileName+$($FileSuffix[$i])  	
            if (Test-Path -Path $tFile)
            {
                Write-Host "Send $FileName$($FileSuffix[$i]) Pass"
                cp $tFile $Destination_Path
            }else
            {
#               Write-Host "$FileName$($FileSuffix[$i]) not exist"	
            }	
        }
    }

    $upFlag=$FileSystemWatcher.LastWriteTime
    Start-Sleep -s 1 
} 
