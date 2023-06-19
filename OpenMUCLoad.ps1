<#
   .Synopsis
This is script loads transfered and controlled data to Power BI data sets - push and history (standard)

#>
param([string]$mUCPath,[string]$summaryFilesPath,[string]$user,[string]$pass,[string]$pushRESTApiString,[double] $ServiceMultiplier = 1/30)
Write-Host "Matrix99 OpenMUCMLoad - mUCPath path $mUCPath"
Write-Host "Matrix99 OpenMUCLoad - path path $SummaryFilesPath"

#This function gives KPIs for service specified in KPIService
Function Get-KPIs{
    param  ([string]$KPIService = "Total_Active_Energy_Import_A")
    
    # test if file with KPI displayed timeStmp exists 
    
   $KPItimeStampFile = $summaryFilesPath + "\KPITimeStampFile.json"
   if (Test-Path -Path $KPItimeStampFile ){
    $rawJson =  Get-Content -Path $KPItimeStampFile -Raw
    [hashtable] $timeStampsHash = ConvertFrom-Json $rawJson -AsHashtable
   }
   else{
    [hashtable]$timeStampsHash = @{}
   }
   
  
   # create timestamp hash table for LP files 
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
    [hashtable]$timeStampsHashLP = @{}
    $todayString = Get-Date -Format "yyyy-MM-dd"
    $minDate = [datetime] $todayString
    $minDate
      # reset dataset - log in
      $password = "LukaNino5559*ph" | ConvertTo-SecureString -asPlainText -Force
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
   Write-Output "Filtered $($lpService.Count)"
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
          "currValue" = $lpService[$i].Comp_Value
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

#Get-KPIs
