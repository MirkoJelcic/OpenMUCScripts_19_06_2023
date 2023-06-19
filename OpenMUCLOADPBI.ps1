#1. Authenticate to Power BI with Power BI Service Admin account

$User = "mirko.jelcic@matrix99.net"
$PW = "LukaNino5559*pa"

$SecPasswd = ConvertTo-SecureString $PW -AsPlainText -Force
$myCred = New-Object System.Management.Automation.PSCredential($User,$SecPasswd)

Connect-PowerBIServiceAccount -Credential $myCred
#2. Get the workspace ID of the dataset(s) to be refreshed

$WSIDAdmin =  Get-PowerBIWorkspace -Scope Organization -Name 'LachNerWS' | 
   Where-Object {$_.Type -eq "Workspace"}  | ForEach-Object {$_.Id}
$WSIDAdmin

#3. Get the dataset ID(s) of the datasets to be refreshed 

#Refresh History dataset
$DSIDRefresh = Get-PowerBIDataset -Scope Organization -WorkspaceId $WSIDAdmin  |  
  Where-Object {$_.Name -eq "Boki"} | ForEach-Object {$_.Id}

$DSIDRefresh   


<# second time
$WSIDAdmin =  Get-PowerBIWorkspace -Scope Organization -Name 'DataFlowTestWS' | 
   Where-Object {$_.Type -eq "Workspace"}  | ForEach-Object {$_.Id}
$WSIDAdmin

#3. Get the dataset ID(s) of the datasets to be refreshed 

#Refresh History dataset
$DSIDRefresh = Get-PowerBIDataset -Scope Organization -WorkspaceId $WSIDAdmin  |  
  Where-Object {$_.Name -eq "CEZDistribution"} #| ForEach-Object {$_.Id}

$DSIDRefresh   
#>

  #4. Build dataset refresh URLs

$RefreshDSURL = 'groups/' + $WSIDAdmin + '/datasets/' + $DSIDRefresh + '/refreshes'

$RefreshDSURL

#5. Execute refreshes with mail on failure

$MailFailureNotify = @{"notifyOption"="MailOnFailure"}

Invoke-PowerBIRestMethod -Url $RefreshDSURL -Method Post -Verbose -Body  $MailFailureNotify


