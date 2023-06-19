<# this script process Profiles as created in OpenMUC
 Profiles in OpenMUC are recorded in data files as HEX strings. 
 

#>
param($mUCPath = "C:\OpenMUCLinux",$dataMUCFilter = "*.dat")

# this is class for file status 
class FileStatusClass{
  [string] $lastWrite;
  [int] $noOfLines;
}

class DeviceChannelCounter{
  [string] $device;
  [string] $computingType;
  [bool] $valid;
}

Function Convert-HEXProfile {

    [CmdletBinding()]
    param ([string] $HexString,  [System.Collections.Generic.List[string]] $ProfileList, [string]$Device)
   
#  define services (measures) in this profile. it should be read from the meters
#$cServices = @( "CLOCK;",	"STATUS;", 	"Active Energy Import (+A) (value);",	"Active Energy Import (+A) Rate 1 (value);",	"Active Energy Import (+A) Rate 1 (value);",	"Active Energy Import (+A) Rate 2 (value);",	"Active Energy Import (+A) Rate 3 (value);",	"Active Energy Import (+A) Rate 4 (value);",	"Reactive Energy Import (+R) (QI+QII) (value)")
    
    $cServices = @( "Date;",	"MeasurementTime;","(empty);", 	"Total_Active_Energy_Import_A;",	"Active Energy Import (+A) Rate 1 (value);",	"Active Energy Import (+A) Rate 2 (value);",	"Active Energy Import (+A) Rate 3 (value);",	"Active Energy Import (+A) Rate 4 (value);",	"Total_Reactive_Energy_Import_R_QI_QII;",
,   "Total_Reactive_Energy_Import_R_QI_QII;", "Instantaneous reactive import power(+R) (value);", "Average apparent power")

    $asciiLines= [System.Collections.Generic.List[string]]@()
    $Header = $cServices[0]+ $cServices[1]+ $cServices[2]
    Write-Output "Processing $Device"
   #HARD CODED 
   # for($i =3; $i -lt 4;$i++){
    # $iArr = 3,4,5,6,7,8,9,10,11
    if (($Device -eq (1037)) -or ($Device -eq "1036")){
      $iArr = 3,9
    }
    else{
     $iArr = 3,8
    }
    if ($Device -eq "2094967"){
      $iArr =,3
    }


    foreach($i in $iArr) {
      $sService = $Device + "_"+ $($cServices[$i]) 
      $Header +=$sService
    }
    $Header = $Header.Trim(";")
   #END HARDCODED

   # Write-Output "Header $Header"
   # $Header =  -join $cServices
   
    $AsciiString = ""
    $iL = $ProfileList.Add($Header)
    #Write-Host "Processing $($HexString.Length) hex characters to ascii for meter $Device, starting with $($HexString.Substring(0,20))"
    
    for($i=0;$i -lt $HexString.Length;$i=$i+2){
       if (($i % 1000000) -eq 0){
          Write-Output "$i chars processed $(Get-Date))"
        }
        $hexSubS = $HexString.Substring($i,2)
        try{
          $asciiC = [char][convert]::toint16($hexSubS,16) #is it new line???
          if ($asciiC -eq "`n"){
             # $AsciiString
              $iA = $asciiLines.Add($AsciiString)
              $AsciiString = ""
          }
          else {
            $AsciiString += $asciiC
          }
        }
        catch{
        # Write-Output "Error parsing char $i : $hexSubS  in $datFile for $Device"
        # $hexSubS | Format-Hex
         #exit
         break
        }
    } # end for hex string processing
    
       
    #process second line of the profile - number of lines in each entry
    try{
   #     $numEntriesProfile = [int32] $asciiLines[0].Split(" ")[2]
        $numEntriesEntry = [int32] $asciiLines[1].Trim().Split(" ")[3]
        Write-Host "Meter $Device - lines for entry:  $numEntriesEntry"
        #$entry   
    }
    catch{
         Write-Host "Error parsing  in $datFile "
         exit
    }
    
     #calculate how many lines to skip for each entry
    $numOfLinesToSkip = $numEntriesEntry + 2
    #process ascii lines -    
    for($i=1; $i-lt $asciiLines.Count;$i+= $numOfLinesToSkip){
           if ($asciiLines[$i].Length -gt 0){
              # extract date time for the entry, it is in the 2nd ascii line for the entry
              #convert to timestamp  yyyy-MM-dd HH:mm:ss 
              $dateLineParsed = $asciiLines[$i+1].Trim().Split(" ")
              $dateYYYY = $dateLineParsed[3] + $dateLineParsed[4]
              $dateMM = [string][convert]::toint16($dateLineParsed[5],16)
              if($dateMM.Length -eq 1){
                  $dateMM = "0" + $dateMM
              }
              $dateDD = [string][convert]::toint16($dateLineParsed[6],16)
              if($dateDD.Length -eq 1){
                $dateDD = "0" + $dateDD
              }
             
              #time 
              $timeHH = [string][convert]::toint16($dateLineParsed[8],16)
              if($timeHH.Length -eq 1){
                $timeHH = "0" + $timeHH
              }
              $timeMM = [string][convert]::toint16($dateLineParsed[9],16)
              if($timeMM.Length -eq 1){
                $timeMM= "0" + $timeMM
              }
              $timeSS = [string][convert]::toint16($dateLineParsed[10],16)
              if($timeSS.Length -eq 1){
                $timeSS= "0" + $timeSS
              }
              $dateYYYY = [string][convert]::toint16($dateYYYY,16)
              $dateOldS = $dateDD +"."+$dateMM +"."+ $dateYYYY + " " + $timeHH + ":"+ $timeMM + ":" + $timeSS
              $dateOld = Get-Date $dateOldS
             
              # HARD coded correction of the clock
              if ($Device -eq "1037" ){
                $dateNew = $dateOld.AddDays(-30)
                $dateNew = $dateNew.AddHours(-7)
                
              }
              else {
                   $dateNew = $dateOld.AddHours(-1)
              } 
              #end HARD coded 
              $profileLine = $dateNew.ToString("yyyyMMdd") + ";         " + $dateNew.ToString("HHmmss") +"; "
      #        $profileLine =  $dateYYYY + $dateMM  +  $dateDD +";        "
      #        $profileLine += $timeHH + $timeMM +  $timeSS + ";      "

            # $profileLine
              # HARD CODED TO GET ONLY ONE PROFILE ENTRY
            #  for($j=3; $j -lt $numEntriesEntry +1;$j++){
           # $jArr = 3,4,5,6,7,8,9,10,11
           if (($Device -eq (1037)) -or ($Device -eq "1036")){
            $jArr = 3,9
          }
          else{
           $jArr = 3,8
          }
          if ($Device -eq "20949677"){
            $jArr =,3
          }
            
            #for($j=3; $j -lt 4 ;$j++){
            # $asciiLines[$i+$j]
             foreach($j in $jArr) {
                   $serviceLineParsed = $asciiLines[$i+$j].Trim().Split(" ")
                   $serviceLine = ";  " + $serviceLineParsed[3]
                  # $serviceLine
                  $profileLine += $serviceLine
              }
              #end HARD CODED
              #add entry to profile lines
              $iL = $ProfileList.Add($profileLine)
          }
           else {
              # Write-Output "Empty line"
           }
     }  

} #end of the function Convert-HEXProfile

Function Get-NonProcessedFiles {
  #                       PREPARE LIST OF NON PROCESSED AND PARTALLY PROCESSED FILES AND NUMBER OF LINES TO SKIP (PARTIALLY PROCESSED - PROCESSED LINES, NON-PROCESSED = 0)
  
    param([string]$filesStat, [string]$dataPath)
     
     #read data about processed files 
     
     $filesList=Get-ChildItem -Path $dataPath -Filter $dataMUCFilter
     if (Test-Path -Path $filesStat){
     # test files for changes and process only changed files    
      $rawJson =  Get-Content -Path $filesStat -Raw
      [hashtable] $hFilesLines = ConvertFrom-Json $rawJson -AsHashtable
      Write-Host "File processing status converted from $filesStat "
  
      # find un-processed files or parts of files 
      foreach ($fileInfo in $filesList){
      #try to find file in hash table of processed files 
        if($hFilesLines.ContainsKey($fileInfo.FullName)){
          Write-Host "File $($fileInfo.FullName) processed "
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
         $fileObject = New-Object -TypeName FileStatusClass
         $hFilesLines.Add($fileInfo.FullName,$fileObject)
         $datFiles.Add($fileInfo.FullName)
        }
      } #end for each
    } 
   else { #non-existing file of processed files, create files hash table
      [hashtable] $hFilesLines = @{}
      foreach($fileData in $filesList){
           $fileObject = New-Object -TypeName FileStatusClass
           $hFilesLines.Add($fileData.FullName,$fileObject)
           $datFiles.Add($fileData.FullName)
      }
   } 
   return($hFilesLines)
} # end of function Get-NonProcessedFiles

Function Get-OpenMucData{
#this function gets data for requested meter/usage point for requested interval in requested diffrent format from source
# parameters 
[CmdletBinding()]
param ([datetime] $StartTime, [datetime] $EndTime , [string]$Device, [PSObject]$Format, [string]$Source = "C:\OpenMUCLinux\Test_Summary")

#Write-Output "I am looking for data for $Device from $StartTime to $EndTime from source $Source"
$o = Get-Member -InputObject $Format
#$o.ToString()
# create hash table with files to read 
[hashtable]$dataTable = @{}

foreach($oO in $o){
    $sTemp = $oO.ToString()
  #  $sTemp
    if ($sTemp.Contains("=")) {
      $arrN = $sTemp.Split(" ")
      $arrN = $arrN[1].Split("=")
      $ind = $arrN[0].IndexOf("_")
       
      #find name of general channel
      if ($ind -gt 0) {
        $channelInst = $arrN[0] #instance of channel for device
        $channelGen = $arrN[0].Substring($ind+1,$arrN[0].Length - $ind -1) #
        $fileToSearch =  "LP_"+ $channelGen+".csv"
        $f= Get-ChildItem -Path $Source -Filter $fileToSearch -Recurse -ErrorAction SilentlyContinue -Force
        if ($f){
          #check if file is already in memory 
       #  $channel=[io.path]::GetFileNameWithoutExtension($f.FullName) 
         $usagePoint = $hChannels[$channelInst].device
        # $usagePoint
         if(-not $dataTable.ContainsKey($f.FullName)) {
       #    Write-Output "I would read and add $f"
           $lines = Import-Csv -Path $f -Delimiter ";"
           $dataTable.Add($f.FullName,$lines)
           #find data for start date
           $countFound = 0
           $lineNo = 0
           $isFirst = $true
           foreach($line in $lines){
            if($line.Device -eq $usagePoint){ 
             try{
             $mTime = [Datetime]::ParseExact($line.MeasurementTime, 'yyyy-MM-dd HH:mm:ss', $null)
            # Write-Output "Comparing $mTime to $StartTime"
              if (($mTime -gt $StartTime) -and ($mTime -lt $EndTime)) {
               #  Write-Output "Measurement time  $mTime greater than $StartTime and less than $EndTime"
              # break;
              if ($isFirst)
              {
                  $isFirst = $false
                  Write-Output "Found $mTime for meter $Device at UsagePoint $usagePoint  for $StartTime and $EndTime in $($f.FullName) "
              }
              $countFound++
              }
             } # end try
             catch{
              Write-Error "Eror in line $lineNo in file $f.FullName - $($line.MeasurementTime) = not correct , Error: $_"
             # exit
            }
            $lineNo++
            } #if line.Device = usagePoint
           } # for each line
         Write-Output "Found $countFound for $Device in $StartTime and $EndTime interval to merge"    
         } #end if file not in memory
      }
    } 
    } 
} # end for each object
    


} # end of function Get-OpenMUCdata

<#                                                        MAIN Program                                                          

Main program

#>
Clear-Host
#$mUCPath = "C:\OpenMUCLinux"


#Write-Host "Multiplier = $ServiceMultiplier"
$startDate = Get-Date
$outPath = $mucPath + "\StagingDataTestNew"
$lpPath = $mUCPath + "\LoadProfilesNew"
$filesStat = $lpPath + "\FilesStat.json"
$configMUCPath = $mUCPath + "\channels.xml"

#  # Read contentent of the config file in XML variable
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
   }
 } #end foreach device


if (Test-Path -Path $outPath){
# directory exists
}
else{
  # create staging data directory
  Write-Output "Create data staging directory $outPath"
  New-Item -Path $outPath  -ItemType "directory" 
}


$datFiles = New-Object 'Collections.Generic.List[string]'
$hFilesLines = Get-NonProcessedFiles -filesStat $filesStat -dataPath $lpPath

<# MAIN LOOP - foreach file #>
$filesList=Get-ChildItem -Path $lpPath -Filter $dataMUCFilter

foreach($datFile in $datFiles) {
  $startDateFile = Get-Date
  Write-Output "Processing: $datFile, start $startDateFile" 
  $datArrayRow = Get-Content $datFile
  $datArray = [System.Collections.Generic.List[string]]@()
  
  #filter leading comments in file 
  $iRow = 0
  foreach($dat in $datArrayRow) {
     if ($dat.Substring(0,1) -eq "#"){
     $iRow++
     }
     else {
     $il =$datArray.Add($dat)
    }
  } # end filtering comments
  
  # process each line (set of LoadProfiles) in file. 
  # first line is header
  $goodLines = New-Object Collections.Generic.List[Int]
  Write-Output "There are $($datArray.Count) lines "
  for($i = 1;$i -lt $datArray.Count;$i++){
     Write-Output "Will process $($datArray[$i].Length) characters in line $i"
     if($datArray[$i].Length -gt 0){
      $i0x = $datArray[$i].IndexOf("0x")
      if ($i0X -eq -1){
        Write-Error "Bad format in hex data for line $i, lenght of line "
        $datArray[$i].Substring(0,150) | Format-Hex
        #break
      }
      else{
      Write-Output "Good format for line $i "
      $datArray[$i].Substring(0,150) | Format-Hex
      $goodLines.Add($i)
      }
    }
   else{
         Write-Output "Lenght 0 for line $i"
      }
  }
  
 # Write-Output "Header profile in $($datArray[0])"
  $headerArr = $datArray[0].Split(";")

  #$profileArr = $datArray[$datArray.Count -2].Split(";")
 
 # Write-Output "There are  elements $($headerArr.Count) in header, $($profileArr.Count) in data" 
  $profileLines = [System.Collections.Generic.List[string]]@()

  <# main loop #>
  ForEach ($iGood in $goodLines) {
    $profileArr = $datArray[$iGood].Split(";")
     for($i=3; $i -lt $profileArr.Count;$i++){
      $deviceName = $headerArr[$i].Trim().Substring(0,$headerArr[$i].Trim().IndexOf("_"))
      $i0X = $profileArr[$i].IndexOf("0x")
      if ($i0X -eq -1){
          Write-Error "Not expected string in hex data for $deviceName entry, length $($profileArr[$i].Length)"
          $profileArr[$i].Substring(0,100) | Format-Hex
          continue
      }
  #   Write-Output "Processing $($profileArr[$i].Substring(0,3))"
     $hs = $profileArr[$i].Substring($i0X+2,$profileArr[$i].Length -$i0X-2)
    #extract name of the device 
    Convert-HEXProfile -HexString $hs -ProfileList $profileLines -Device $deviceName
   #write to respective file, 
  #  Write-Host "I would write to $deviceName dir/file"
    
    $endDate = Get-Date
    $noLines = $profileLines.Count
    Write-Output "Start: $startDate , End: $endDate, Lines: $noLines  " 
    $csvFile = $outPath + "\" +$deviceName +".dat"
    if (Test-Path -Path $csvFile){
        Out-File -FilePath $csvFile -InputObject $profileLines.GetRange(1, $noLines -1) -Append
    }
    else{
        Out-File -FilePath $csvFile -InputObject $profileLines -Force
    }
    $profileLines.Clear()
     } # end 
  }



  $endDateFile = Get-Date
  Write-Output "File $datfile processed, start $startDateFile, end: $endDateFile"
  #remember data about file processed
  $fileBuff= Get-Content -Path $datFile
  $fileAtt = Get-Item -Path $datFile
  $hFilesLines[$fileAtt.FullName].lastWrite = [string]$fileAtt.LastWriteTime
  $hFilesLines[$fileAtt.FullName].noOfLines = $fileBuff.Length
} # End foreach datfile

# memorize processed files
#write file list to .json file 
ConvertTo-Json -InputObject $hFilesLines | Out-File -FilePath $filesStat -Force

$startDate = Get-Date

#logical control of staging data

Get-ChildItem -Path $outPath -Filter "*.dat"| 
ForEach-Object{
  
  Write-Output $_.FullName
  $sortCsv = Import-Csv -Path $_.FullName -Delimiter ";" | Sort-Object -Property "Date", "MeasurementTime" -Unique 
  $sortCsv | Export-Csv -Path $_.FullName -NoTypeInformation -Delimiter ";"
  $data = Get-Content $_.FullName
  $data | ForEach-Object {$_.replace('";"',';').TrimStart('"').TrimEnd('"')} | Set-Content $_.FullName
 
  $firstItem = $sortCsv[0]
 # $firstItem
  $startInt = [datetime]::ParseExact($firstItem.Date+$firstItem.MeasurementTime,'yyyyMMddHHmmss',$null)
 # Write-Output "start: $startInt"
  ForEach ($csvItem in ($sortCsv | Select-Object -skip 1)){
    $endInt = [datetime]::ParseExact($csvItem.Date+$csvItem.MeasurementTime,'yyyyMMddHHmmss',$null)
   # Write-Output "start: $startInt end: $endInt"
   $diff = New-TimeSpan  -Start $startInt -End $endInt
   #Write-Output $diff.TotalMinutes
   if( $diff.TotalMinutes -gt 15) {
        # Write-Output "Interval $startInt - $endInt greater than 15, = $($diff.TotalMinutes) for $($_.FullName) "
        $deviceName = [IO.Path]::GetFileNameWithoutExtension( $_.FullName )
        Get-OpenMucData -StartTime $startInt -EndTime $endInt -Device $deviceName -Format $firstItem
   }

   if ($startInt -eq $endInt){
     Write-Output "Interval $start = $start2 are the same for $($_.FullName) ?"
   }  
    $startInt = $endInt
  }
}

# & $PSScriptRoot\OpenMUCDataScript.ps1



