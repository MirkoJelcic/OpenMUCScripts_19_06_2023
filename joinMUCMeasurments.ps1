param([string]$File1 = "C:\OpenMUCLinux\StagingDataNew\20949677.dat" ,  [string]$File2 ="C:\OpenMUCLinux\StagingDataNew\1075.dat"
 , [datetime]$JoinTime="2021-01-21 10:15",[string]$inJoinDevice ="20949677", [string]$outJoinDeviceID ="1075" )

Function Clean-File {
param([string]$File)
$data = Get-Content $File
$data | ForEach-Object {$_.replace('";"',';').TrimStart('"').TrimEnd('"')} | Set-Content $File
$data = Get-Content $File
$data | ForEach-Object {$_.replace('";',';')} | Set-Content $File
}

Write-Output "Join $File1 with $File2 at $JoinTime" 
try{

  Clean-File -File $File1    
  Clean-File -File $File2

  $file1Csv = Import-Csv -Path $File1 -Delimiter ";" | Sort-Object -Property "Date", "MeasurementTime" -Unique 
  #Get-Member -InputObject $file1Csv
  $index1 = 0
  $index2 = 0
  ForEach ($objCsv1 in $file1Csv){
  
    $s1 = $objCsv1.Date + " " + $objCsv1.MeasurementTime
    $t1= [datetime]::parseexact($s1, 'yyyyMMdd HHmmss', $null)
    if( $t1 -lt $JoinTime){
      #  Write-Output " $t1 OK "
        $index1++
        #$outCsv.Add($objCsv1)
    }
    else {
        Write-Output " $t1 OK "
        break;
    }
  }

  $file2Csv = Import-Csv -Path $File2 -Delimiter ";" | Sort-Object -Property "Date", "MeasurementTime" -Unique 

  ForEach ($objCsv2 in $file2Csv){
    $s2 = $objCsv2.Date + " " + $objCsv2.MeasurementTime
    $t2= [datetime]::parseexact($s2, 'yyyyMMdd HHmmss', $null)
    if( $t2 -ge $JoinTime){
        Write-Output " $t2 OK "
        break;
        #$outCsv.Add($objCsv1)
    }
    else {
        $index2++
    } 
  }
 Write-Output ("Index1 = $index1, Index2 = $index2")
 # write output file, combining two lists 
 $outPath = Split-Path $File1 -Parent
 $fileOut1 = $outPath+"\"+ $outJoinDeviceID +"_1.dat"
 $fileOut2 = $outPath+"\"+ $outJoinDeviceID +"_2.dat"
 Write-Output "i wil join $fileOut1 and $fileOut2"
 $o1 = $file1Csv[0..$index1]
 $end = $file2Csv.Count - 1
 $o2 = $file2Csv[$index2..$end]
 $o1.Count
 #Get-Member -InputObject $o1[0]
# change last element in 
#$o1[-1]
 
 $o1 | Export-Csv -Path $fileOut1 -Force -NoTypeInformation -Delimiter ";"
 $o2 | Export-Csv -Path $fileOut2 -Force -NoTypeInformation -Delimiter ";"
 Clean-File -File $fileOut1    
 Clean-File -File $fileOut2
 #change header of the file
$data1 = Get-Content $fileOut1
$s1 = $data1[0].Replace($inJoinDevice,$outJoinDeviceID)
$data1.Item(0)=$s1
$data1[0]
$data1| Set-Content -Path $fileOut1 -Force
}
catch{
    Write-Error "Error $File1 with $File2 $error[0] "
}