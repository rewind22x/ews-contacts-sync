#Load file containing ids
. "ids.ps1"
# Functions
function SyncContacts ($emailAddress){

    function WriteToLogFile ($logMessage){
        Add-content $logFile -value $logMessage
    }

    # Create a log file
    $logFolderPath = $LogPath
    $logDate = Get-Date -Format "yyyMMdd_HHmm"
    $logFileName = $emailAddress+"_"+$logDate+".log"
    $logFile = New-Item -itemType File -Path $logFolderPath -Name $logFileName

    WriteToLogFile("Accessing "+$emailAddress)

    # use credentials to access mailbox of user 
    $Service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $emailAddress)

    $userContactsFolder = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Contacts)

    $publicContactsFolder = $FolderToContactsSource

    $publicRootFolder = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::PublicFoldersRoot)

    $publicFoldersSource = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($Service, $publicRootFolder)

    $publicFoldersPath = $publicContactsFolder.Split("\")

    # get the exact path for the folder spcified in $spuserContactsFolder
    for ($i = 1; $i -lt $publicFoldersPath.length; $i++)
    {
        $folderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(1)
        $searchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName, $publicFoldersPath[$i])
        $findFolderResults = $Service.FindFolders($publicFoldersSource.Id, $searchFilter, $folderView)
        if ($findFolderResults.TotalCount -gt 0)
        {
            $publicFoldersSource = $findFolderResults.Folders[0]
            WriteToLogFile("Found  "+$publicContactsFolder)
        }
        else
        {
            WriteToLogFile($publicContactsFolder+" Not Found")
            exit
        }
    }

    $userItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)
    $userFindItemResults = $Service.FindItems($userContactsFolder, $userItemView)

    # delete all contacts from the mailbox that match the parameters
    WriteToLogfile("Deleting $Description1 contacts...")
    $count = 1
    foreach ($userItem in $userFindItemResults.Items | Where-Object {$_.CompanyName -eq $Description1})
    {
        WriteToLogFile($userItem.DisplayName+" "+$count)
        $userItem.Delete([Microsoft.Exchange.WebServices.Data.DeleteMode]::HardDelete)
        $count++
    }
    WriteToLogfile("Deleting $Description2 contacts...")
    $count = 1
    foreach ($userItem in $userFindItemResults.Items | Where-Object {$_.CompanyName -eq $Description2})
    {
        WriteToLogFile($userItem.DisplayName+" "+$count)
        $userItem.Delete([Microsoft.Exchange.WebServices.Data.DeleteMode]::HardDelete)
        $count++
    }

    $pfItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)
    $pfFindItemResults = $Service.FindItems($publicFoldersSource.Id, $pfItemView)

    # copy the contents of of the folder specified in $pfContacts to the mailbox Contacts folder
    WriteToLogfile("Copying contacts...")
    $count = 1
    foreach ($pfItem in $pfFindItemResults.Items)
    {
        WriteToLogFile($pfItem.DisplayName+" "+$count)
        $pfItem.Copy($userContactsFolder)
        $count++
    }
}
Import-Module 'C:\Program Files\Microsoft\Exchange\Web Services\2.0\Microsoft.Exchange.WebServices.dll'

# Acquire the Oauth token
$MsalParams = @{
    ClientId = $AppClientId
    TenantId = $TenantId   
    Scopes   = "https://outlook.office.com/EWS.AccessAsUser.All"   
}

$MsalRenew = @{
    # grant_type = "refresh_token"
    ClientId = $AppClientId
    TenantId = $TenantId   
    Scopes   = "https://outlook.office.com/EWS.AccessAsUser.All" 
}
 
$MsalResponse = Get-MsalToken @MsalParams
$EWSAccessToken  = $MsalResponse.AccessToken

# Create connection to EWS
$Service = [Microsoft.Exchange.WebServices.Data.ExchangeService]::new()
 
# Use Modern Authentication to authenticate
$Service.Credentials = [Microsoft.Exchange.WebServices.Data.OAuthCredentials]$EWSAccessToken
$Service.UseDefaultCredentials = $false
 
# Check EWS connection
$Service.Url = "https://outlook.office365.com/EWS/Exchange.asmx"
# EWS connection is Success if no error returned.

# Single mailbox or distribution list
$userInput = Read-Host -Prompt "Please enter an email address or distribution group name:"

if ($userInput -match '@'){
    SyncContacts($userInput)
}

else {        
    Connect-ExchangeOnline -UserPrincipalName $TokenUser
    $userObjects = Get-DistributionGroupMember -Identity $userInput
    $counter = 1
    foreach ($SMTPAddress in $userObjects.PrimarySMTPAddress){
        SyncContacts($SMTPAddress)

        $MsalResponse = Get-MsalToken @MsalRenew
        $EWSAccessToken  = $MsalResponse.AccessToken
        $Service = [Microsoft.Exchange.WebServices.Data.ExchangeService]::new() 
        # Use Modern Authentication to authenticate
        $Service.Credentials = [Microsoft.Exchange.WebServices.Data.OAuthCredentials]$EWSAccessToken
        $Service.UseDefaultCredentials = $false        
        # Check EWS connection
        $Service.Url = "https://outlook.office365.com/EWS/Exchange.asmx"
        $counter++

    }
    Disconnect-ExchangeOnline -Confirm:$false
}