<# 
 .Synopsis
This script monitors changes that are done on the specified path and 
copy them to the destination directory and calls script with destionation 
directory as a parameter.

"C:\Users\Mirko Jelcic\OneDrive\openmuc\data1\ascii", "C:\Users\Mirko Jelcic\OneDrive\openmuc\data2\ascii","C:\Users\Mirko Jelcic\OneDrive\openmuc\data3\ascii")

#>

#param([String[]]$Paths =@("C:\Users\Mirko Jelcic\OneDrive\openmuc\data1\ascii", "C:\Users\Mirko Jelcic\OneDrive\openmuc\data2\ascii","C:\Users\Mirko Jelcic\OneDrive\openmuc\data3\ascii"), $Script = "", $DestinationPath = "C:\Users\Mirko Jelcic\OneDrive\LachNer\Test_StagingData", $OffLineStartFile="20210401", $FileFilter ="202*.dat")
param([String[]]$Paths =@("C:\Users\mirko\OneDrive\openmuc\data1\ascii", "C:\Users\mirko\OneDrive\openmuc\data2\ascii","C:\Users\mirko\OneDrive\openmuc\data3\ascii"), $Script = "", $DestinationPath = "C:\LachNerDir\Test_StagingData", $OffLineStartFile="20210401", $FileFilter ="202*.dat")

 #check if destination path exist, if not create it!!!

 if (Test-Path -Path $DestinationPath ){
  #OK 
  Write-Output "Write to $DestinationPath..."

}  
else{
  # Write-Error "Destination path $Destinationpath not found, exit..."
  Write-Output "Create $DestinationPath..."
  New-Item -Path $DestinationPath  -ItemType "directory"
  exit
}
 #restart one drive
Stop-Process -Name OneDrive -Force
#    ************************** TO DO - check type of ondrive installation and start existng exe

#Start-Process $env:LOCALAPPDATA\Microsoft\OneDrive\OneDrive.exe
#    ************************** /allusers

Start-Process "C:\Users\mirko\AppData\Local\Microsoft\OneDrive\OneDrive.exe"


$now = Get-Date

foreach($Path in $Paths){
  Write-Host "Matrix99 OpenMUCFileWatcher - input path $Path"
}

Write-Host "Matrix99 OpenMUCFileWatcher - output path $DestinationPath"


Write-Host "Start at $now, waiting to sychronize OneDrive"

 Write-Host "Offline Start File " $($offLineStartFile)
#wait for three minutes
Start-Sleep -Seconds 180



try {
  #create hash table for each monitored directory with its sufix
  $pathsSufixes = @{}
  for($i=0; $i -lt $Paths.Count; $i++){ 
      $s = "_" + [string]$i
      $pathsSufixes.Add($Paths[$i],$s) 
      Write-Host "Create sufix $s for $($Paths[$i])"
  }
 

  if ($OffLineStartFile -ne ""){
      #get items from source directory
    foreach($Path in $Paths){ 
      Write-Host "Synchronizing files changed off-line for $Path"
      Get-ChildItem -Path $Path -Filter $FileFilter|
      ForEach-Object {
          $sourceFileName = [System.IO.Path]::GetFileNameWithoutExtension($_.FullName) 
          git 
          #Write-Output "Examining $sourceFileName, comparing $shortName with $OffLineStartFile  "
          if ( $shortName -gt $OffLineStartFile) {
            $destinationFile = $sourceFileName + $pathsSufixes[$Path]  + [System.IO.Path]::GetExtension($_.FullName) 
            $destinationFilePath = $DestinationPath + "\" +$destinationFile
            if (Test-Path  $destinationFilePath){
              $dp =Get-Item -Path $destinationFilePath
            #  Write-Host "Destination: $($dp.FullName) $($dp.LastWriteTimeString)"
            #  Write-Host "Source: $($_.FullName) $($_.LastWriteTimeString)"
             # continue
             if ($_.LastWriteTime -ne $dp.LastWriteTime){
              Write-Output "Re-write $destinationFile to $destinationFilePath"
              Copy-Item -Path $_.FullName -Destination $destinationFilePath -Force
             }
            }
            else{
             Write-Output "Write $destinationFile to $destinationFilePath"
             Copy-Item -Path $_.FullName -Destination $destinationFilePath -Force
            }
        }
      }
    }
  }
} 
catch{
   Write-Error "Error $error[0] synchronizing files changed off-line"
   exit
}

