<# 
.Synopsis
This is script that controls metering data. Data are controlled by checking duplicates, time holes(missing data in time intervals), ....

#>

param([string]$mUCPath = "C:\OpenMUCLinux\Test_StagingData",[string]$summaryFilesPath = "C:\OpenMUCLinux\Test_Summary",[double] $ServiceMultiplier = 1/30)
##### CLASSES
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

    class cTimeGap {
      [System.DateTime] $gapBegin;
      [System.DateTime] $gapEnd;
    }
#####                               FUNCTIONS
Function Get-TimeGapsAndDuplicates {
    param  ([string]$Path, [string]$Interval = "15",[hashtable] $TableOfGaps)
  <# This function finds interval "holes"- time gaps between reading (measuremets) that are bigger than $Interval (temporary in minutes)
    It
    #>
  
    Write-Output "Checking $Path for reading gap greater than $Interval)"
    $iInterval = [int]$Interval
    try{ 
      Get-ChildItem -Path $Path | 
      ForEach-Object{
        try{
        <# if ([String]::IsNullOrWhiteSpace((Get-content $_.FullName)))  {
             #empty file
            Write-Output "$($_.FullName) empty"
          }
          else {#>
           # Write-Output "Try to processing $($_.FullName)"
           # To do - define testing algorithm acccording to type of service (filename)
    
          # to do - file filter 
          # if (([System.IO.Path]::GetFileNameWithoutExtension($_.FullName) -eq "Total_Active_Energy_Import_A") -or
          if   ([System.IO.Path]::GetFileNameWithoutExtension($_.FullName) -eq "LP_Total_Active_Energy_Import_A"){
          #    ){
            $currFile = $_.FullName    
            Write-Output "Checking $($_.FullName)"
            $outFileName =   [System.IO.Path]::GetDirectoryName($_.FullName)+"\S"+ [System.IO.Path]::GetFileName($_.FullName)
          #  Write-Output "Processing $($_.FullName)"
            Import-Csv -Path $_.FullName -Delimiter ";" | Sort-Object  -Property  Device , MeasurementTime -Unique | Export-Csv -Path $outFileName -NoTypeInformation -Delimiter ";" -Force
            Move-Item -Path $outFileName -Destination $_.FullName -Force
            $uniqueCsv =  Import-Csv -Path $_.FullName -Delimiter ";"
            if ($uniqueCsv.Count -gt 0){ 
            Write-Output "Processing $($_.FullName)"
           # $uniqueCsv | Get-Member 
           # Get-Member -InputObject $uniqueCsv[0]
           #  $iRow = 0
             for ($i=0; $i -lt $uniqueCsv.Count - 1; $i++) {
               # Get-Member -InputObject $uniqueCsv[$i]
              if ($uniqueCsv[$i].Device -eq $uniqueCsv[$i +1].Device) {
                  $device = $uniqueCsv[$i].Device
                  $start = [datetime]$uniqueCsv[$i].MeasurementTime
                  $start2 = [datetime]$uniqueCsv[$i+1].MeasurementTime
                  
                  #DEBUG 
                <# if ($uniqueCsv[$i].Device -eq "La10_kluthe"){
                      if (($start.Day -eq 8) -and ($start.Hour -eq 17 ) -and ($start.Month -eq 7)){
                         Write-Host "bla bla $start"
                      }
  
                  }
                 #>
  
                  $diff = New-TimeSpan  -Start $start -End $start2
                  #Write-Output $diff.TotalMinutes
                  if( $diff.TotalMinutes -gt $iInterval) {
                        Write-Output "Interval $start - $start2 greater than $Interval, = $($diff.TotalMinutes) for $device"
                        # add to the hash table of gaps
                        $gapKey = $device +"_" + [System.IO.Path]::GetFileNameWithoutExtension($_.FullName)
                        
                        $oTimeGap = [cTimeGap]::new()
                        $oTimeGap.gapBegin = $start
                        $oTimeGap.gapEnd = $start2
                        #check if key exists in table of gaps
                        if($TableOfGaps.ContainsKey($gapKey)){
                     #     Write-Output "I would update $gapKey with $start and $start2 "
                          #read table of gaps and update element
                          $oListOfGaps = $TableOfGaps[$gapKey]
                          $oListOfGaps.Add($oTimeGap)
                         
                        }
                        else{
                          # create new element and add first gap
                          
                      #    Write-Output "I would add $gapKey with $start and $start2 reading times as new element" 

                          $oListOfGaps = New-Object System.Collections.Generic.List[System.Object]
                          $oListOfGaps.Add($oTimeGap)
                          $TableOfGaps.Add($gapKey,$oListOfGaps)
                        }

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

# Function Get-KPIs calculate Key Performances Indicators and writes them in JSON file that is displayed on the dashborad in real-time


#This function gives KPIs for service specified in KPIService
Function Get-KPIs{
  param  ([string]$KPIService = "Total_Active_Energy_Import_A")
  
 
  <#
  # test if file with KPI displayed timeStmp exists 
 $KPItimeStampFile = $summaryFilesPath + "\KPITimeStampFile.json"
 if (Test-Path -Path $KPItimeStampFile ){
  $rawJson =  Get-Content -Path $KPItimeStampFile -Raw
  [hashtable] $timeStampsHash = ConvertFrom-Json $rawJson -AsHashtable
 }
 else{
  [hashtable]$timeStampsHash = @{}
 }
 #>

 # create or read timestamp hash table for LP files 
 $serviceLPfile = $summaryFilesPath + "\LP_" + $KPIService + ".csv"
 if (Test-Path -Path $serviceLPFile ){
    # OK, there is file
  }
 else {
        Write-Error - Message "Could not find $serviceLPFile"
        Exit
 }

 $LP_KPItimeStampFile = $summaryFilesPath + "\LP_KPITimeStampFile.json"
 if (Test-Path -Path $LP_KPItimeStampFile ){
  $rawJson =  Get-Content -Path $LP_KPItimeStampFile -Raw
  [hashtable] $timeStampsHashLP = ConvertFrom-Json $rawJson -AsHashtable
  # find the oldest time stamp in table 
  $minDate = [datetime]"9999-12-31"
  foreach($kEy in $timeStampsHashLP.Keys){
       if($timeStampsHashLP[$kEy] -lt $minDate){
          $minDate = [datetime]$timeStampsHashLP[$kEy]
           Write-Output "minDate = $minDate"

       }
  }  

 }
 else{
   # if 
  [hashtable]$timeStampsHashLP = @{}
 <#region  $todayString = Get-Date -Format "yyyy-MM-dd"
  $minDate = [datetime] $todayString
  #>
  $minDate = (Get-Date).AddDays(-5)
  Write-Host("start  new stream file from $minDate")
    # reset dataset - log in
    $password = "LukaNino5559*pa" | ConvertTo-SecureString -asPlainText -Force
    $username = "mirko.jelcic@matrix99.net"
    $credential = New-Object System.Management.Automation.PSCredential($username, $password)
    Connect-PowerBIServiceAccount -Credential $credential
    # data set created by powerbi service
    $dataSetID = "45066da8-e46e-4036-beae-8f534eb70bac"
    $tableName = "RealTimeData"
    #The below URL should be used for My Workspace datasets
    $myWorkSpaceUrl = 'https://api.powerbi.com/v1.0/myorg/datasets/' + $datasetId + '/tables/' + $tableName + '/rows'
    Invoke-PowerBIRestMethod -Url "$myWorkSpaceUrl" -Method Delete
    Disconnect-PowerBIServiceAccount

 }

 
 $payloadArr = @()
 $lPService = Import-Csv -Path $serviceLPfile -Delimiter ";"  | Where-Object {[datetime]$_.MeasurementTime -GT $minDate} #  -Contains $todayString
 Write-Output "Filtered $($lpService.Count) from $minDate"
 #$lpDevice=""
 for ($i = 0;$i -lt $lpService.Count;$i++){
   <#  if($lpService[$i].Device -ne $lpDevice){
       $lpDevice = $lpService[$i].Device
       $lpLastTime = [datetime]$lpService[$i].MeasurementTime 
       Write-Output "New device $lpDevice with timestamp $lpLastTime"
     }
     else {
      
     } #>
     $kEy = $lpService[$i].Device  +"-" +$KPIService 
     if ($timeStampsHashLP.ContainsKey($kEy)){
      $lpLastTime = [datetime]$lpService[$i].MeasurementTime   
    #  Write-Output "There is time stamp for $kEy, it is $($timeStampsHashLP[$kEy])   "
     if ($timeStampsHashLP[$kEy] -eq $lpLastTime){
       #Write-Output "New time stamp is the same..."
     }
     else {
    
      $timeStampsHashLP[$kEy] = [datetime]$lpService[$i].MeasurementTime 
      
      Write-Output "New time stamp for $kEy is $lpLastTime, write to push file ...."
      $payload = @{
        "usagePoint" = $lpService[$i].Device
        "currValue" = $lpService[$i].Comp_Value / 7
        "maxValue" = 200
        "targetValue" =120
        "timeStamp" = $lpService[$i].MeasurementTime
        }
        $payloadArr += $payload

     }
   }
   else {
     Write-Output "No time stamp for $kEy, time stamp $lpLastTime added "
     $timeStampsHashLP.Add($kEy,$lpLastTime)
   }

 }
# write to push file 
#$payloadArr
#$timeStampsHashLP
# write time stamps table to file
 

 <#            TO DO:  process pending calculations to refine real time display  
  #table to keep pending calculations (Load Profiles) between data collection cycles 
  # reading 

[hashtable]$hMeterPeCa =@{} 
$payloadArr = @()
#  try to read meters status file 
$metStatFile = $mUCPath + "\meterPendingCalculationsFile.json"
 if (Test-Path -Path $metStatFile ){
  #OK, there is mter pending calculations file
  $rawJson =  Get-Content -Path $metStatFile -Raw
  [hashtable] $hMeterPeCa_Temp = ConvertFrom-Json $rawJson -AsHashtable
  # change array to Generic.List 
  $summ = 0
 #Write-Output "Creating KPI table"
 
  
 
  
  foreach($kEy in $hMeterPeca_Temp.Keys){
    # is this our service ???
   # $kEy
    $arr = $kEy.Split("-")
    $usagePoint = $arr[0]
    $arr =$arr[1].Split("_",2)
    $device = $arr[0]
    $service = $arr[1]
   
    if ($service -eq $KPIService){ 
     # Write-Output "Usage Point:  $usagePoint  Device:  $device service:  $service "
      try{
        #Write-Host ($kEy, "Start: ",$vTemp.intervalBegin," End: ",  $vTemp.firstTimeThisInterval)
        
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
        #calculate KPI -
        $currValue = $vTemp.lastValueThisInterval  -  $vTemp.firstValueThisInterval + $vTemp.carryValue
        $summ += $currValue
        
        #Write-Output " $device $usagePoint $($vTemp.lastValueThisInterval)    $($vTemp.firstValueThisInterval)  $($vTemp.carryValue) "
        #$summ
        $oPendingCalculations.intervalValues=@()
        foreach($val in $vTemp.intervalValues){
            $oPendingCalculations.intervalValues.Add($val)
        }
       
        #TO DO: ------- read from profile files and filter data to be writen to the push file
        #write data to hash table - prepare push file
        $hMeterPeCa.Add($kEy,$oPendingCalculations)
        # write new time stamp to hash table 
        if ($timeStampsHash.ContainsKey($kEy)){
           #Write-Output "There is time stamp for $kEy, it is $($timeStampsHash[$kEy])   "
          if ($timeStampsHash[$kEy] -eq $vTemp.intervalEnd){
       #     Write-Output "New time stamp is the same..."
          }
          else {
          # Write-Output "New time stamp for $kEy is $($vTemp.intervalEnd)"
           $timeStampsHash[$kEy] = $vTemp.intervalEnd
          }
        }
        else {
          Write-Output "No time stamp for $kEy, time stamp $($vTemp.intervalEnd) added "
          $timeStampsHash.Add($kEy,$vTemp.intervalEnd)
        }

        $payload = @{
          "usagePoint" = $usagePoint
          "currValue" = $currValue
          "maxValue" = 200
          "targetValue" =120
          "timeStamp" = $vTemp.intervalEnd
          }
          $payloadArr += $payload
        #$oPendingCalculations
      }
      catch{
          Write-Error "Error in intitialization $kEy $Error[0]"
          exit
      } #try catch
    } #end if
   } # for each
  # write time stamps table to file
  ConvertTo-Json -InputObject $timeStampsHash | Out-File -FilePath $KPItimeStampFile -Force

 }
 else{
  Write-Error "No Meter Pending Calculation file, exit...."
  exit
 } # no status (pending calculations) file - error 

#refresh dashboard

# while (1) { #TEST LOOP
#>
try{
  # write to dataset
  $endpoint = "https://api.powerbi.com/beta/fbd26fd4-e0f5-4bf0-b237-9be32fc209bc/datasets/45066da8-e46e-4036-beae-8f534eb70bac/rows?key=DNVMgWCFpboMjrxsHmbNjnmwOHde2Fopf8iGF30y3ju5VMiT9fiyry2BTT2CKEd4zL02laTW%2FShloHYbf%2BWPlg%3D%3D"
  Invoke-RestMethod -Method Post -Uri "$endpoint" -Body (ConvertTo-Json @($payloadArr))
  ConvertTo-Json -InputObject $timeStampsHashLP | Out-File -FilePath $LP_KPItimeStampFile -Force
}
catch{
  Write-Error "Push not sent..."
  

}
Start-Sleep -Seconds 10
  #} #end TEST LOOP
} #end function Get-KPIs





#########                           MAIN
#Get (calculate) KPIs
Get-KPIs

###### Time Gaps 
#is there a time gaps file
$timeGapsFile = $mUCPath + "\timeGapsFile.json"
if (Test-Path -Path $timeGapsFile ){ #???
  $rawJson =  Get-Content -Path $timeGapsFile -Raw
  [hashtable] $hTimeGaps = ConvertFrom-Json $rawJson -AsHashtable
  Write-Output "Read $($hTimeGaps.Count) element of time gaps"
}
else { # create it 
     Write-Output "Creating Time Gaps data, this will take a while..."
     [hashtable] $hTimeGaps = @{} 
     Get-TimeGapsAndDuplicates -Path $summaryFilesPath -TableOfGaps $hTimeGaps
}
# remember gaps
ConvertTo-Json -InputObject $hTimeGaps | Out-File -FilePath $timeGapsFile -Force