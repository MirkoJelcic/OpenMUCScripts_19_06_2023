<#      This script that reads meters starting data 
        from excel file that contains meters, usage points and instateniuous values 
        of counters 

#>



Clear-Host
$mUCPath = "C:\OpenMUCData\TestExcel"
#$configMUCPath = $mUCPath + "\channels.xml"
$excelPath = $mUCPath+"\Odečty elektroměry.xlsx"
# add type for device metering value
$Source = @"
  public class DeviceChannelCounter{
      public string device;
      public double lastValue = 0;
      public bool valid = false;
  }
"@

Add-Type -TypeDefinition $Source
# import table of counter values 
$tableCounter = Import-Excel -Path $excelPath -WorksheetName "Odečty elektroměrů" -StartRow 5 -EndRow 26

# import table of acc values 
#$tableAcc = Import-Excel -Path $excelPath -WorksheetName "Odečty elektroměrů" -StartRow 30 -EndRow 51

#$table of the Czech months
$cMonths = @( "leden",	"únor", 	"březen",	"duben",	"květen",	"červen",	"červenec",	"srpen",	"září", 	"říjen", "listopad", "prosinec")
# if file of devices exists, read devices from it
$devStatFile = $mUCPath + "\devStatFile.json"
if (Test-Path -Path $devStatFile){
   $rawJson =  Get-Content -Path $devStatFile -Raw
   [hashtable] $hChannels = ConvertFrom-Json $rawJson -AsHashtable
}
else{
   #else, read devices from excel table 
   #and save it to hashtable
   Write-Output "File $devStatFile does not exist"
     [hashtable] $hChannels = @{} 
   # $device
   foreach($dMeter in  $tableCounter) {
       # $
       $deviceObject = New-Object DeviceChannelCounter
       $channelName = [string]$dMeter.Meter + "-Active Energy"
       $deviceObject.device = $dMeter.Usage_Point
       $hChannels.Add($channelName, $deviceObject)
       #Write-Output "$($dMeter.Meter) added...."
    }
  }
  
  #define data .csv
  $pathDevData = $mUCPath+"\InitData" +".csv"
  # $pathDevData
   if(-not (Test-Path -Path $pathDevData)){
           #create header
           $hEader = "Meter; Usage_Point;  DateTime  ;  Value; Comp. Value"
           # $datArray[0] | Out-File -FilePath $pathDevData
           $hEader | Out-File -FilePath $pathDevData
           Write-Output "Device data file $pathDevData created, with header $hEader"
   }
 
# start date of measurements 
$dateMeter = [datetime]"12/31/2019 23:59:59"
#$dMeter
foreach($row in $tableCounter){
     #  Write-Output "================="
     for ($i = 0; $i -lt $cMonths.Length-2; $i++){
     # $cMonths[$i]
     # $row.($cMonths[$i])
     # calculate Quantity (acc. value)
     # first find device and channel
     $channelName = [string]$row.Meter + "-Active Energy"
     $deviceObject = $hChannels[$channelName]
     # check if counter is valid
     if ($deviceObject.valid){
         $value = [string] ([double]$row.($cMonths[$i])-[double]$deviceObject.lastValue)
         $deviceObject.lastValue = [double]$row.($cMonths[$i])
     }
     else {#not valid, make it valid
      $deviceObject.valid = $true
      $value = 0
      $deviceObject.lastValue = [double]$row.($cMonths[$i])
     } 
    #update device(meter) object
      $hChannels[$channelName] = $deviceObject 
      $newDate =  $dateMeter.AddMonths($i+1)
       #convert to timestamp  yyyy-MM-dd HH:mm:ss
      $newDateString = $newDate.ToString("yyyy-MM-dd hh:mm:ss")
    # make row
     $outStr = [string]$row.Meter + ";" + $row.Usage_Point+";"+ $newDateString +";"+  $row.($cMonths[$i])+";"+ $value
    # $outStr 
     if ($row.($cMonths[$i])){
        # Write-Output "Not empty"

      }
      else 
      {
            Write-Output "empty $i for $($row.Meter)" 
           # $outStr Grt
      }
     $outStr | Add-Content -Path $pathDevData
    }
}
$dataObject = Import-Csv  $pathDevData -Delimiter ";"
#  export 
#convert devices (meter) file to json
#cls
ConvertTo-Json -InputObject $hChannels | Out-File -FilePath $devStatFile -Force
$excelFile = $mUCPath + "\OpenMUCExcel.xlsx"
#now, make excel output
Export-Excel -Path $excelFile -InputObject $dataObject -TableName "Active_Energy"  -Append -Show