<# 
  .Synopsis
  OpenMUCMainScript.ps1 is script of the scripts for open muc data collection process. 
  It comprises  Extract, Transform, Control and  Load scripts. 
#>
param ([String[]] $InputPaths, [string]$StagingDataPath, [string]$SummaryDataPath)

foreach($Path in $InputPaths){
  Write-Host "Matrix99 OpenMUCMainScript - input path $Path"
}

Write-Host "Matrix99 OpenMUCMainScript - datapath path $StagingDataPath"
Write-Host "Matrix99 OpenMUCScriptMain - datapath path $SummaryDataPath"

# create command line for File Watcher = Extract
$commandExtract = "$PSScriptRoot\OPenMUCFileWatcher.ps1 -Paths "
for($i=0;$i -lt $InputPaths.Length -1;$i++){
  $commandExtract += "'"+ $InputPaths[$i] +"',"
}
$commandExtract += "'" + $InputPaths[$InputPaths.Length -1] +"'" + " -DestinationPath '$StagingDataPath' -OffLineStartFile '20210401' -FileFilter '202*.dat'"
Write-Host "Matrix99 ETL - Extract $commandExtract"
#Invoke-Expression -Command $commandExtract
#pause

$commandTransform = "$PSScriptRoot\OpenMUCDataScript.ps1 -mUCPath '$StagingDataPath' -summaryFilesPath '$SummaryDataPath'  -ServiceMultiplier  (1/30)"
Write-Host "Matrix99 ETL - Transform $commandTransform"
#Invoke-Expression -Command $commandTransform
#pause
#param([string]$mUCPath = "C:\OpenMUCLinux\Test_StagingData",[string]$summaryFilesPath = "C:\OpenMUCLinux\Test_Summary",[double] $ServiceMultiplier = 1/30)
$commandControl = "$PSScriptRoot\LachNerDataControl.ps1 -mUCPath '$StagingDataPath' -summaryFilesPath '$SummaryDataPath'  -ServiceMultiplier  (1/30)"
Write-Host "Matrix99 ETL - Control $commandControl"





while (1) {
$startDateScript = Get-Date
$startDate = Get-Date

Write-Host "================== Matrix99 ETL DLMS start: $(Get-Date)====================================="
Write-Host "*************************************************************************"
Write-Host "Matrix99 OpenMUC Main - EXTRACT start $startDate"
Write-Host "*************************************************************************"
Invoke-Expression -Command $commandExtract
#& "$PSScriptRoot\OPenMUCFileWatcher.ps1"
Write-Host "Matrix99 OpenMUC EXTRACT start $startDate - EXTRACT End $(Get-Date)"


$startDate = Get-Date
Write-Host "**************************************************************************"
Write-Host "Matrix99 OpenMUC Main - TRANSFORM start $startDate"
Write-Host "*************************************************************************"
#& "$PSScriptRoot\OpenMUCDataScript.ps1"
Invoke-Expression -Command $commandTransform
Write-Host "Matrix99 OpenMUC TRANSFORM $startDate - TRANSFORM End $(Get-Date)"

$startDate = Get-Date
Write-Host "*************************************************************************"
Write-Host "Matrix99 OpenMUC Main - CONTROL start $startDate"
Write-Host "*************************************************************************"
#& "$PSScriptRoot\OpenMUC_DataControll.ps1"
Invoke-Expression -Command $commandControl
Write-Host "Matrix99 OpenMUC CONTROL start $startDate - CONTROL End $(Get-Date)"

Write-Host "==========================================================================="
Write-Host "================== Matrix99 ETL DLMS SCRIPT ENDS -  start: $startDateScript  end: $(Get-Date)====================================="
Write-Host "==========================================================================="
Start-Sleep -Seconds 120
}



