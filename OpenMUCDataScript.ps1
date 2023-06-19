<#    
      .Synopsis
      This script scans files in OpenMUC directory and converst 
      OpenMUC files so data are organized in such manner:
      device  (CIM EndDEvice - meter) - directory
      channel (CIM EdndDeviceFunction) - file inside device directory; channel = meter measurement 
      (active power, reactive power etc..,  instantenous and demand).

      Also, there is summary directory that contains all readings for the channels (measurement).
      Each channel forms a file in the summary directory

    
#>


 
param([string]$mUCPath = "C:\OpenMUCLinux\Test_StagingData",[string]$summaryFilesPath = "C:\OpenMUCLinux\Test_Summary", [double] $ServiceMultiplier = 1/30)
#param([string]$mUCPath = "C:\Users\Mirko Jelcic\OneDrive\LachNer\Test_StagingData",[string]$summaryFilesPath = "C:\Users\Mirko Jelcic\OneDrive\LachNer\Test_Summary", [double] $ServiceMultiplier = 1/30)
#LOG definition
$Logfile = $mUCPath + "\"+ "$(gc env:computername).log"

Function LogWrite
{
   Param ([string]$logstring)

   Add-content $Logfile -value $logstring
}

# CLASSES
class DeviceChannelCounter{
  [string] $device;
  [string] $computingType;
  [bool] $valid;
}

class FileTypeClass{
   [string] $lastWrite;
   [int] $noOfLines;
}

class MeterPendingCalculations{
# [System.DateTime]$lastTimePreviousInterval;
 #[double] $lastValuePreviousInterval;
 [double] $carryValue
 [System.DateTime]$firstTimeThisInterval;
 [double] $firstValueThisInterval;
 [System.DateTime]$lastTimeThisInterval;
 [double] $lastValueThisInterval;  
 [System.DateTime] $intervalBegin;
 [System.DateTime] $intervalEnd;
 [System.Collections.Generic.List[double]] $intervalValues;
 [bool] $valid;
}

class LinesToWrite {
 [System.Collections.ArrayList] $lines
}


Function Update-ProcessedFiles {
  #    Update LIST OF NON PROCESSED AND PARTALLY PROCESSED FILES AND NUMBER OF LINES TO SKIP (PARTIALLY PROCESSED - PROCESSED LINES, NON-PROCESSED = 0)
  param([string] $Path,[string[]] $datFiles) # name of file with file names and processed line counts
     #update data about processed files in status file  () name is defined in configuration)
     $filesStat = $Path + "\FilesStat.json"

     foreach($datFile in $datFiles){
      try{ 
        $fileAtt = Get-Item -Path $datFile 
        Write-Host "Update status $datFile ...."
        $fileBuff= Get-Content -Path $datFile
        $hFilesLines[$datFile].lastWrite = [string]$fileAtt.LastWriteTime
        $hFilesLines[$datFile].noOfLines = $fileBuff.Length
      }
      catch{
        Write-Error "Error updating files processing status "
        exit
      }
     }
     #write files processinfg status list to .json file 
     ConvertTo-Json -InputObject $hFilesLines | Out-File -FilePath $filesStat -Force
}
Function Get-NonProcessedFiles {
  #  PREPARE LIST OF NON PROCESSED AND PARTALLY PROCESSED FILES AND NUMBER OF LINES TO SKIP (PARTIALLY PROCESSED - PROCESSED LINES, NON-PROCESSED = 0)
  param([string] $Path,[string] $dataMUCFilter) # name of file with file names and processed line counts
       
     #read data about processed files from file whose name is constant 
     $filesStat = $Path + "\FilesStat.json"
     $filesList=Get-ChildItem -Path $Path -Filter $dataMUCFilter
     #test if processed files file exists - if not, make it!!!
     if (Test-Path -Path $filesStat){
     # test files for changes and process only changed files    
      $rawJson =  Get-Content -Path $filesStat -Raw
      [hashtable] $hFilesLines = ConvertFrom-Json $rawJson -AsHashtable
      Write-Host "File processing status converted from $filesStat "
  
      # find un-processed files or parts of files 
      foreach ($fileInfo in $filesList){
      #try to find file in hash table of processed files 
        if($hFilesLines.ContainsKey($fileInfo.FullName)){
       #   Write-Host "File $($fileInfo.FullName) processed "
          $sLW = $hFilesLines[$fileInfo.FullName].lastWrite
          #compare dates of write
          $sLW2 = [string]$fileInfo.LastWriteTime
          #un-processed part of file
          if ($sLW -ne $sLW2){
            Write-Host "File$($fileInfo.FullName) LastWrite dates not equal $sLW, $slW2, part of file  not processed"
            #calculate number of lines to skip
            $datFiles.Add($fileInfo.FullName)
          }
          else{
            #Write-Output "Date equal $sLW  "
          }
        }
        else{
         Write-Host "File $($fileInfo.Name) not processed "
         $fileObject = New-Object -TypeName FileTypeClass
         $hFilesLines.Add($fileInfo.FullName,$fileObject)
         $datFiles.Add($fileInfo.FullName)
        }
      } #end for each
    } 
   else { #non-existing file of processed files, create files hash table
      [hashtable] $hFilesLines = @{}
      foreach($fileData in $filesList){
           $fileObject = New-Object -TypeName FileTypeClass
           $hFilesLines.Add($fileData.FullName,$fileObject)
           $datFiles.Add($fileData.FullName)
      }
   } 
   return($hFilesLines)
 } # END function Get-NonProcessed files 
  
Function Convert-OpenMUCDat {
  [CmdletBinding()]
  param ([string] $datFile,  [double] $linesToSkip,[int] $intervalLP = 0)
    #read .dat file and create .csv output file for each device, 
    # update summary file for the channel
  $startDate = Get-Date
  Write-Output "Processing: $datFile, start: $(Get-Date), lines to skip: $linesToSkip" 
  $datArrayRow = Get-Content $datFile
  $datArray = New-Object 'Collections.Generic.List[string]'

  #skip comments
  $iTrailedLines = 0
  foreach($dat in $datArrayRow) {
      if ($dat.Substring(0,1) -eq "#"){
        $iTrailedLines++
      }
      else {
        $il =$datArray.Add($dat)
      }
  }

  #process header of the .dat file
  $headerArr = $datArray[0].Split(";")
  if ($headerArr.Length -le 3){
     throw "$datArray[0] is not a valid header"
  }

  # if there is more that one channel, do it in loop
  
  for ($i=3; $i -lt $headerArr.Length;$i++){
    # find device for this channel
    $hS =  $headerArr[$i].Trim().Trim('"')
    $deviceObject = $hChannels[$hS]
    
    if (-not $deviceObject){
        # did not find channel in configuration
      throw "$hs channel can not be found, defined in file $datFile"
    } 
             
    #check if directory for this device reading exists, if not create it
    $dirDevData = $mUCPath + "\"+ $deviceObject.device
    if(-not (Test-Path -Path $dirDevData)){
             New-Item -ItemType Directory -Force -Path $dirDevData
             Write-Host "Device data directory $dirDevData created"
    }

   #now check if output file for this device and channel exists
   #TO DO - 
   #$ind =  $datFile.IndexOf('_')
   # $fileDatePart = $datFile.Name.Substring(0,$ind)
  
  <# create path for device and channel data file, 
   # so data are organized in such manner:
   #   device (meter)       -  directory
   #   channel (measurment) -  file inside device directory
   #>


   #   first, define directory
  # $pathDevData = $dirDevData +"\" +  $fileDatePart + "_"+ $headerArr[$i].Trim() +".csv"
   $pathDevData = $dirDevData +"\" + $headerArr[$i].Trim() +".csv"
  # $pathDevData
   if(-not (Test-Path -Path $pathDevData)){
           #create header
           $hEader = " Device;MeasurementTime;Meas_Value; Comp_Value"
           # $datArray[0] | Out-File -FilePath $pathDevData
           $hEader | Out-File -FilePath $pathDevData
           Write-Output "Device data file $pathDevData created, with header $hEader"
   }
  }

  #if there are lines to skip (part of the file is already processed), calculate the start 
  if ($linesToSkip -gt 0){
      Write-Output "Skip $linesToSkip in $datFile"
      $iStart = $linesToSkip - $iTrailedLines
  } 
  else {
      $iStart = 1

  }
  for($iDat=$iStart; $iDat -lt $datArray.Count;$iDat++){
       #
        #convert to timestamp  yyyy-MM-dd HH:mm:ss    
        $datArraySplit = $datArray[$iDat].Split(";")     
        $datArraySplit = $datArraySplit.Trim()
        #create time stamp
        #with date and time 
        $sDate = $datArraySplit[0].Substring(0,4)+"-" + $datArraySplit[0].Substring(4,2)   +"-" + $datArraySplit[0].Substring(6,2)
        $sTime =$datArraySplit[1].Substring(0,2) + ":" + $datArraySplit[1].Substring(2,2)
        $datRow = $sDate + " " + $sTime + ";"
        # without date  
        # $datRow =   $datArraySplit[1].Substring(0,2) + ":" + $datArraySplit[1].Substring(2,2) + ":" + $datArraySplit[1].Substring(4,2)+";"
        for ($iCh=3; $iCh -lt $headerArr.Length;$iCh++){
          #find correct device
          $hS =  $headerArr[$iCh].Trim()
          $deviceObject = $hChannels[$hS]
          $mPeCaKey = $deviceObject.device+"-"+$hS
         # $dev = $deviceObject.device
          $valS =$datArraySplit[$iCh]
         # Write-Output "Value for $dev is $val "
          if (!$deviceObject.valid){
             # first data for this device
             # check if data in row are numeric
             if ($valS -match "^[-+]?([0-9]*\.[0-9]+|[0-9]+\.?)$" ){
             # Write-Output "Value for $dev is numeric ($val) "
              # $dateReading =$hMeterPeCa[$mPeCaKey].lastTimeThisInterval
                
                $val = [UInt64]$valS
                #$datRowCh = $datRow + $datArraySplit[$iCh]
                $datRowCh = $datRow + [string]$val 
                $dateReading = [datetime]($sDate + " "+ $sTime)
                $hMeterPeCa[$mPeCaKey].lastTimeThisInterval =  $dateReading
                $hMeterPeCa[$mPeCaKey].lastValueThisInterval = $val
                $hMeterPeCa[$mPeCaKey].firstTimeThisInterval =  $dateReading
                $hMeterPeCa[$mPeCaKey].firstValueThisInterval = $val
                $hMeterPeCa[$mPeCaKey].intervalValues.Add($val)
                $hMeterPeCa[$mPeCaKey].intervalBegin = $dateReading.AddMinutes( - ($dateReading.Minute % $intervalLP))
                $hMeterPeCa[$mPeCaKey].intervalEnd = $hMeterPeCa[$mPeCaKey].intervalBegin.AddMinutes($intervalLP)
                $hMeterPeCa[$mPeCaKey].carryValue= 0
                $hMeterPeCa[$mPeCaKey].valid = $true
                $deviceObject.valid = $true
             }
             else{
             # Write-Output "Value for $deve is not numeric($val) "
             }
             
             $Acc = "0"
          } # end if for first reading for this device
          else {
           # Write-Output "Second itd.. data for $dev"
           if ($valS -match "^[-+]?([0-9]*\.[0-9]+|[0-9]+\.?)$" ){
              # Write-Output "Value for $dev is numeric($val) "
              $val = [uint64] $valS
              #$datRowCh = $datRow + $datArraySplit[$iCh]
              $datRowCh = $datRow + [string]$val 
              $dateStart = $hMeterPeCa[$mPeCaKey].lastTimeThisInterval
              $valStart = $hMeterPeCa[$mPeCaKey].lastValueThisInterval
              $dateReading =[datetime]($sDate + " "+ $sTime)
               #DEBUG 
 <#if ($deviceObject.Device -eq "La10_kluthe"){
  if (($dateReading.Day -eq 8) -and ($dateReading.Month -eq 7) -and ($dateReading.Hour -eq 18)){
     Write-Host "bla bla $start"
  }
  #>
              if ($dateReading -lt $dateStart){
                Write-Output "$dateReading -lt $dateStart for $($deviceObject.device)"
                continue
              }
              #calculate calculated field 
              # $Acc = [string](([double]$val -[double] $valStart) * $ServiceMultiplier)

              # **** to do 
              $currVal =[double]$val 
              $prevVal=[double] $valStart
              if ($currVal -lt $prevVal){
              #current counter lower than previous  counter 
                $result = 0
               # Write-Warning "Counter $currVal less than $prevVal for $($deviceObject.device), measure $hS, result truncated to 0"
              }
              else {
                $result = $currVal - $prevVal #
              }
              
              $Acc = [string]$result #
              $timeSpan= New-TimeSpan -Start $dateStart -End $dateReading 
          
              try{
                  $mPeCaKey = $deviceObject.device+"-"+$hS
                  $iB = $hMeterPeCa[$mPeCaKey].intervalBegin
                  $iE = $hMeterPeCa[$mPeCaKey].intervalEnd
                #check if reading datetime is in interval
                # TODO:  
              #  if ($dateReading -gt $iE) { 
                 if ($dateReading -ge $iE) { 
                  #write ouput for the last interval and start new interval 
                  #create new interval
                 # find diffrence between 
                  
                  $diffValAll =  $val - $valStart
                  if ($diffValAll -lt 0){
                    #Write-Output "$val less than $valStart for key $mPeCaKey"
                    #[Console]::ReadKey()
                    $diffValAll = 0
                  }

                  $diffTimeAll =  New-TimeSpan  -Start $dateStart -End $dateReading
                  $diffTimeNew =  New-TimeSpan  -End $dateReading -Start $iE
                  $diffTimePrev =  New-TimeSpan  -Start $dateStart -End $iE

                  
                  $timeShare =[double]([double]$diffValAll /[double] $diffTimeAll.TotalMinutes)   
                  $diffPrev = [double]( [double]$timeShare * [double] $diffTimePrev.TotalMinutes)
                  $diffNew = [double]([double]$timeShare * [double]$diffTimeNew.TotalMinutes)
              
                  # calculate diffference for the previous interval
                  $diffSum = [math]::Round(([double]$diffPrev + [double]$hMeterPeCa[$mPeCaKey].carryValue))
                  
                  if ($val -ge $valStart){
                   $accLP =  [string]($diffSum +  [double]$hMeterPeCa[$mPeCaKey].lastValueThisInterval - [double]$hMeterPeCa[$mPeCaKey].firstValueThisInterval)
                  }
                  else {
                    $accLP = "0"
                  }
                  $datRowLP = $deviceObject.device +";" + $iE.ToString('yyyy-MM-dd HH:mm:ss') + ";" +[string]$val+ ";" + $accLP
                 
                  try{
                    $imKey=$headerArr[$iCh].Trim()
                    $indName = $imKey.IndexOf("_")
                    $imKey = $imKey.Substring($indName + 1, $imKey.Length - $indName -1 ) 
                    $il = $inMemoryLinesLP[$imKey].lines.Add($datRowLP)
                  }
                 catch{
                   Write-Error "Error writing to in-memory LP for $imKey $error"
                   exit
                }
                  #define new interval
                  $hMeterPeCa[$mPeCaKey].intervalBegin = $dateReading.AddMinutes( - ($dateReading.Minute % $intervalLP))
                  $iBnew =$hMeterPeCa[$mPeCaKey].intervalBegin
                  $hMeterPeCa[$mPeCaKey].intervalEnd =  $iBnew.AddMinutes($intervalLP)
                
         <#     DEBUG
          if($mPeCaKey -eq "1076-1076_Total_Active_Energy_Import_A"){
            Write-Host "Date: $dateReading, Key: $mPeCaKey value: $val last.val $valLast"
            Write-Host $datRowLP
            $hMeterPeCa[$mPeCaKey]
          }
          end DEBUG  #>    

          
                $hMeterPeCa[$mPeCaKey].firstTimeThisInterval = $dateReading #?????
                $hMeterPeCa[$mPeCaKey].firstValueThisInterval = $val
                $hMeterPeCa[$mPeCaKey].carryValue = $diffNew
                $hMeterPeCa[$mPeCaKey].intervalValues.Clear()
                  	
                }
                else{
                }
                 #add value to values in this interval
                $hMeterPeCa[$mPeCaKey].lastTimeThisInterval = $dateReading  
                $hMeterPeCa[$mPeCaKey].intervalValues.Add($val)
                $hMeterPeCa[$mPeCaKey].lastValueThisInterval = $val
               # Write-Output "Date: $dateStart, value:$($deviceObject.LastValue) ,start interval $($hMeterPeCa[$mPeCaKey].intervalBegin), end interval: $($hMeterPeCa[$mPeCaKey].intervalEnd)"
              }
              catch{
                Write-Error "Error in pending calculations, $($Error[0]) "
                exit
              }

              if($timeSpan.Minutes -gt $intervalLP){
                    Write-Output "Time span $timeSpan.Minutes start: $dateStart, ends: $dateReading for $($deviceObject.device), measure $hS"
              }
             }
             else{
             # Write-Output "Value for $dev is not numeric($val) "
               $Acc = ""           
              }

          }  # end of the processing other readings 
          
         # if ($Acc)
          
          $datRowCh = $datRowCh + ";" + $Acc
          if ($Acc -ne "") {
               #create name of the correct file 
              $dirDevData = $mUCPath + "\"+ $deviceObject.device
              # $pathDevData = $dirDevData +"\" +  $fileDatePart + "_"+ $headerArr[$iCh].Trim() +".csv"
               #file without 
               $pathDevData = $dirDevData +"\" +  $headerArr[$iCh].Trim() +".csv"
               $datRowOut =  $deviceObject.device+ ";" + $datRowCh
               $datRowCh = ""
               #write to in-memory lines 
               try{
                  $imKey = $deviceObject.device+"-"+$headerArr[$iCh].Trim()
                  $il = $inMemoryLines[$imKey].lines.Add($datRowOut)
                }
               catch{
                 Write-Error "Error writing to in-memory for $imKey $error"
                 exit
              }
              #######write directly to file 
              # $datRowOut | Add-Content -Path $pathDevData
          } #end if Acc -ne ""
        } # end for channels
  } #end for datArray (complete data processed - first phase)


  #write from in-memory lines to respective files for each device
  #and summary files
  foreach($imKey in $inMemoryLines.Keys){
    # create name of file 
    $sKeys=$imKey.Split("-")
    $pathDevData = $mUCPath +"\" +  $sKeys[0]+"\" +$sKeys[1]+ ".csv"
    try{
      $indName = $sKeys[1].IndexOf("_")
   }
   catch{
       Write-Error "Not correct channel name $sKeys[1]"
       exit
   }
    $sumName =  $sKeys[1].Substring($indName + 1, $sKeys[1].Length - $indName -1 )
    $sumFileName = $sumName + ".csv"
    $lpSumFileName ="LP_" + $sumFileName
    $sumFilePath = $summaryFilesPath +"\" +$sumFileName
    $lpsumFilePath = $summaryFilesPath +"\" +$lpsumFileName
    try{
     $inMemoryLines[$imKey].lines | Add-Content -Path $pathDevData
     $inMemoryLines[$imKey].lines | Add-Content -Path $sumFilePath

    # $indName = $imKey.IndexOf("_")
    # $imKeyLP = $imKey.Substring($indName+1, $imKey.Length - $indName -1)
     $imKeyLP = $sumName
     # do not make LP_file 
     $inMemoryLinesLP[$imKeyLP].lines |  Add-Content -Path $lpsumFilePath
     # 
     $dateSufix= Get-Date -Format "yyyyMMddHHmmss"
     $datIncFile = $sumFileName + $dateSufix
     $inMemoryLines[$imKey].lines.Clear()
     $inMemoryLinesLP[$imKeyLP].lines.Clear()
    }
    catch{
      Write-Error "Error writting $pathDevdata, $Error"
      # to do clear resources
      exit
    }
  }
  #  remember current status of channels
  $devStatFile = $mUCPath + "\devStatFile.json"
  ConvertTo-Json -InputObject $hChannels | Out-File -FilePath $devStatFile -Force
  $metStatFile = $mUCPath + "\meterPendingCalculationsFile.json"
  ConvertTo-Json -InputObject $hMeterPeCa | Out-File -FilePath $metStatFile -Force
  $endDate = Get-Date
  Write-Output "File:$datFile processed -  Start: $startDate , End: $endDate "      

}  #End function Convert-OpenMucDat

Function Get-TimeHolesAndDuplicates {
  param  ([string]$Path, [string]$Interval = "15")
<# This function finds interval "holes" between reading (measuremets) that are bigger than $Interval (temporary in minutes)
  It
  #>

  Write-Output "Checking $Path for reading intervals greater than $Interval"
  $iInterval = [int]$Interval
  try{ 
    Get-ChildItem -Path $Path | 
    ForEach-Object{
      try{ 
       #region 
       
       #endregion Write-Output "Is empty? $($_.FullName) time $(Get-Date)"
     <#  if ([String]::IsNullOrWhiteSpace((Get-content $_.FullName)))  {
           #empty file
          # Write-Output "$($_.FullName) empty"
        }#>
 #       else {#
         # Write-Output "Try to processing $($_.FullName)"
         # To do - define testing algorithm acccording to type of service (filename)
  
        # to do - file filter 
       # if (([System.IO.Path]::GetFileNameWithoutExtension($_.FullName) -eq "Total_Active_Energy_Import_A") -or
        if   ([System.IO.Path]::GetFileNameWithoutExtension($_.FullName) -eq "LP_Total_Active_Energy_Import_A"){
        #    ){
          Write-Output "Sorting $($_.FullName) time $(Get-Date)"
          $currFile = $_.FullName    
          $outFileName =   [System.IO.Path]::GetDirectoryName($_.FullName)+"\S"+ [System.IO.Path]::GetFileName($_.FullName)
        #  Write-Output "Processing $($_.FullName)"
          Import-Csv -Path $_.FullName -Delimiter ";" | Sort-Object  -Property  Device , MeasurementTime -Unique | Export-Csv -Path $outFileName -NoTypeInformation -Delimiter ";" -Force
          Move-Item -Path $outFileName -Destination $_.FullName -Force
          $uniqueCsv =  Import-Csv -Path $_.FullName -Delimiter ";"
          if ($uniqueCsv.Count -gt 0){ 
          Write-Output "Processing $($_.FullName) time $(Get-Date)"
         # $uniqueCsv | Get-Member 
         # Get-Member -InputObject $uniqueCsv[0]
         #  $iRow = 0
           for ($i=0; $i -lt $uniqueCsv.Count - 1; $i++) {
             # Get-Member -InputObject $uniqueCsv[$i]
            if ($uniqueCsv[$i].Device -eq $uniqueCsv[$i +1].Device) {
                $device = $uniqueCsv[$i].Device
                $start = [datetime]$uniqueCsv[$i].MeasurementTime
                $start2 = [datetime]$uniqueCsv[$i+1].MeasurementTime
                
              #DEBUG section for specific device and date
              <# if ($uniqueCsv[$i].Device -eq "La10_kluthe"){
                  #  if (($start.Day -eq 8) -and ($start.Hour -eq 17 ) -and ($start.Month -eq 7)){
                       Write-Host "bla bla $start"
                    }

                }
              #>

                $diff = New-TimeSpan  -Start $start -End $start2
                #Write-Output $diff.TotalMinutes
                if( $diff.TotalMinutes -gt $iInterval) {
                     #log this 
                     $logString ="Interval $start - $start2 greater than $Interval, = $($diff.TotalMinutes) for $device" 
                   #   Write-Output $logString
                      LogWrite $logString

                }

                if ($start -eq $start2){
                  Write-Output "Interval $start = $start2 are the same for $device ?"
                }  

                if ($start -gt $start2){
                  Write-Output "Staring time $start > $start2 greater that ending time for $device"
                }   

                $mValStart = [int]$uniqueCsv[$i]."Meas_Value"
                $mValEnd = [int]$uniqueCsv[$i+1]."Meas_Value"
                if ($mValStart -gt $mValEnd){
                 Write-Host "Staring value $mValStart > $mValEnd greater that ending value for $device at $start to $start2"
                }

                $cValStart = [int]$uniqueCsv[$i]."Comp_Value"
                if ($cValStart -lt 0){
                  Write-Host "Computed value $cValStart negative for $device at $start"
                 }
                
            } # end if for the device

              <#
               
              if ($uniqueCsv[$i].Device -eq "Adj_Sklad"){
                  $startT = Get-Date ("21/3/2021 05:00")
                  $endT = $uniqueCsv[$i].MeasurementTime
                  $diffT = New-TimeSpan  -Start $startT -End $endT
                  if ($diffT.Days -eq 0){
                  Write-Output "Time: $($uniqueCsv[$i].MeasurementTime), Device: $($uniqueCsv[$i].Device, $mValStart, $cValStart )"
                  }
              }
              #>
          } # end for 
        }
      } # end file filter (TO DO)
    # } #end else if file not empty
    } # end try
    catch{
        Write-Output "Error checking $currFile ,row: $($i+2) Error: $_"
        exit
      }
    } #end ForEach-Object
  } # end outer try
  catch{
    Write-Error "Unrecoverable eror , Error: $_"
    exit
  }

 
}  #End function Get-TimeHolesAndDuplicates


<#                                  MAIN PROGRAM                             #>

#Clear-Host
if (Test-Path -Path $mUCPath ){
  #OK
}  
else{
   Write-Error "OpenMUC path $mUCPath not found, exit..."
   exit
}

$configMUCPath = $mUCPath + "\channels.xml"

$startProgram = Get-Date
Write-Output "Service Multiplier: $ServiceMultiplier "

#write lines first in memory and then in file for each device and channel
[hashtable]$inMemoryLines = @{}
#write summary LoadProfiles lines first in memory and then in summary files for all profiles
[hashtable]$inMemoryLinesLP = @{}

#table to keep pending calculations (Load Profiles) between data collection cycles 
[hashtable]$hMeterPeCa =@{} 

$datFiles = New-Object 'Collections.Generic.List[string]'

# if file of devices (meters) exists, read devices status from it
$devStatFile = $mUCPath + "\devStatFile.json"
if (Test-Path -Path $devStatFile){
   $rawJson =  Get-Content -Path $devStatFile -Raw
   [hashtable] $hChannels = ConvertFrom-Json $rawJson -AsHashtable
}
else {
  # Read contentent of the config file in XML variable
  try {
    [xml]$configXMl = Get-Content $configMUCPath
  }
  catch{
    Write-Error "Can not find MUC config file $configMUCPath"
    Exit
  }
   #read devices from MUC config file, save it to hashtable, create header to summary files 
   #$headerSum = "DateTime;"
   [hashtable] $hChannels = @{} 
   foreach ($device in $configXMl.configuration.driver.device ){
     foreach($channel in  $device.channel) {
       # read devices and channels
       $deviceObject = [DeviceChannelCounter]::new()
       $deviceObject.device = $device.id
       $deviceObject.computingType= $channel.computingType
       $hChannels.Add($channel.id, $deviceObject)
      
       #create pending calculation table
       $keyPeCa = $deviceObject.device+"-"+$channel.id
       $oPendingCalculations = [MeterPendingCalculations]::new()
       $oPendingCalculations.intervalValues = @()
       $hMeterPeCa.Add($keyPeCa,$oPendingCalculations)
     }
      <#create device name in header
       if($headerSum.IndexOf($device.id) -eq -1){
        # add device in header
        $headerSum  += $device.id
        $headerSum +=";"
       }#>
   } #end foreach device
} # end else 


#Write-Output "Create $headerSum header"
# check if summary files directory and files exist. If not create them
if (Test-Path -Path $summaryFilesPath ){
    #OK
}  
else{
  Write-Output "Create $summaryFilesPath"
  New-Item -Path $summaryFilesPath  -ItemType "directory"
}

#check reading summary files 
foreach($channelid in $hChannels.Keys){
  try{
     $indName = $channelid.IndexOf("_")
  }
  catch{
      Write-Error "Not correct channel name $channelid"
      exit
  }
  
  $sumFileName = $summaryFilesPath +"\" + $channelid.Substring($indName + 1, $channelid.Length - $indName -1 ) + ".csv"
  if (Test-Path -Path $sumFileName  ){
    #OK
   } 
   else{
     #create summary file for this channel for all devices (meters) 
     Write-Output "Create $sumFileName"
     # TO DO
     # New-Item -Path $sumFileName  -ItemType "file" -Value $headerSum
     $hEader = " Device;    MeasurementTime;  Meas_Value; Comp_Value`n"
     New-Item -Path $sumFileName  -ItemType "file" -Value $hEader

   }
   
   $LPsumFileName = $summaryFilesPath +"\" + "LP_" + $channelid.Substring($indName + 1, $channelid.Length - $indName -1 ) + ".csv"
   if (Test-Path -Path $LPsumFileName  ){
     #OK
    } 
    else{
      #create summary file for this channel for all devices (meters) 
      Write-Output "Create $LPsumFileName"
      # TO DO
      # New-Item -Path $sumFileName  -ItemType "file" -Value $headerSum
      $hEader = " Device;MeasurementTime;  Meas_Value; Comp_Value`n"
      New-Item -Path $LPsumFileName  -ItemType "file" -Value $hEader
 
    }
}


#read calculation pending status file into pending table, if file does not exist
#create hash table
$metStatFile = $mUCPath + "\meterPendingCalculationsFile.json"
if (Test-Path -Path $metStatFile ){
  #OK
  $rawJson =  Get-Content -Path $metStatFile -Raw
  [hashtable] $hMeterPeCa_Temp = ConvertFrom-Json $rawJson -AsHashtable
  # change array to Generic.List 
  foreach($kEy in $hMeterPeca_Temp.Keys){
   try{
    $vTemp = $hMeterPeCa_Temp[$kEy]
    $oPendingCalculations = [MeterPendingCalculations]::new()
    # copy object 
    $oPendingCalculations.intervalBegin = $vTemp.intervalBegin
    $oPendingCalculations.intervalEnd = $vTemp.intervalEnd
    $oPendingCalculations.lastTimeThisInterval = $vTemp.lastTimeThisInterval
    $oPendingCalculations.firstTimeThisInterval = $vTemp.firstTimeThisInterval
    $oPendingCalculations.lastValueThisInterval = $vTemp.lastValueThisInterval
    $oPendingCalculations.firstValueThisInterval = $vTemp.firstValueThisInterval
    $oPendingCalculations.carryValue = $vTemp.carryValue
    $oPendingCalculations.intervalValues=@()
    foreach($val in $vTemp.intervalValues){
        $oPendingCalculations.intervalValues.Add($val)
    }
    $hMeterPeCa.Add($kEy,$oPendingCalculations)
   }
   catch{
        Write-Error "Error in intitialization $kEy $Error[0]"
        exit
   }
  }

}  
else{#TO DO: 

   Write-Output "Create $metStatFile "
   New-Item -Path $metStatFile  -ItemType "file"
}
  #Try to read data about channels from json file 
$channelsFile = $mUCPath + "\channelsFile.json"
if (Test-Path -Path $channelsFile ){
   #OK
  $rawJson =  Get-Content -Path $channelsFile -Raw
  [hashtable] $channelsHash = ConvertFrom-Json $rawJson -AsHashtable
}
else{
    #empty table 
  [hashtable] $channelsHash = @{}
}



#create in-memory lines for readings and interval readings (Load Profile), 
# also, check channelsFile 
foreach($channelid in $hChannels.Keys){
       $deviceObject = $hChannels[$channelid]
       $linesToWrite = [LinesToWrite]::new()
       $keyLines = $deviceObject.device+"-"+$channelid
      # Write-Output "I will add $keyLInes"
       $inMemoryLines.Add($keyLines,$linesToWrite)
       $inMemoryLines[$keyLines].lines = @{}
       try{
        $indName = $channelid.IndexOf("_")
       }
       catch{
         Write-Error "Not correct channel name $channelid"
         exit
       }


    
       # LP lines 
       $keyLinesLP = $channelid.Substring($indName + 1, $channelid.Length - $indName -1 )
       if($inMemoryLinesLP.ContainsKey($keyLinesLP)){
       }
       else{
         Write-Output "Add $keyLinesLP channel to interval readings"
         $linesToWriteLP = [LinesToWrite]::new()
         $inMemoryLinesLP.Add($keyLinesLP,$linesToWriteLP)
         $inMemoryLinesLP[$keyLinesLP].lines = @{}
       }
       
       #channels file - hash table 
       if($channelsHash.ContainsKey($keyLinesLP)){
             
       }
       else{
          
       }

  }
  


#$hFilesLines = Get-NonProcessedFiles -filesStat $filesStat -Path $mUCPath
$hFilesLines = Get-NonProcessedFiles -Path $mUCPath -dataMUCFilter "*.dat" 
# sort file names
Sort-Object $datFiles

#First loop for processing non-processed or partially processed files  
#*****   LOOP FOR RAW files *******
foreach($datFile in $datFiles){
     Write-Host "Process $datFile"
     try{
       #process file  
      # $fileAtt = Get-Item -Path $datFile
      # Convert-OpenMUCDat -datFile $datFile -linesToSkip  $hFilesLines[$fileAtt.FullName].noOfLines -intervalLP 15
      Convert-OpenMUCDat -datFile $datFile -linesToSkip  $hFilesLines[$datFile].noOfLines -intervalLP 15
      }
     catch{
        Write-Host "Error processing $datFile, error: $Error "
        exit
     }
} #end for processing of files 

#update hash table and file processing status file
Update-ProcessedFiles -Path $mUCPath -datFiles $datFiles


# 

$datFiles = New-Object 'Collections.Generic.List[string]'

# to do - update procesesd files                  
$hFilesLines = Get-NonProcessedFiles -Path $summaryFilesPath -dataMUCFilter "*.csv" 
#$datFiles 

#Second loop for processing non-processed or partially processed files  
#*****   LOOP FOR SUMMARY files *******
if (Test-Path $summaryFilesPath){
  Get-TimeHolesAndDuplicates -Path $summaryFilesPath -Interval 15
}
else{
  Write-Error "$summaryFilePath does not exist" 
  exit
}

Write-Host "End of program start: $startProgram , end:$(Get-Date) "



           
      
     

















