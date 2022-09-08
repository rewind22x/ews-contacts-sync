#Load file containing ids
. "C:\TmpExchange\ids.ps1"
# Functions

function WriteToLogFile ($logMessage){
    Add-content $logFile -value $logMessage
}

Import-Module 'C:\Program Files\Microsoft\Exchange\Web Services\2.0\Microsoft.Exchange.WebServices.dll'

# Acquire the Oauth token
$MsalParams = @{
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

$userMailbox = Read-Host -Prompt "Enter the email address"


if (!$userMailbox)
{
  Write-Host "Variable is null"
}
else
{
        # Create a log file
        $logFolderPath = "C:\Scripts\logs\"
        $logDate = Get-Date -Format "yyyMMdd_HHmm"
        $logFileName = $userMailbox+"_"+$logDate+".log"
        $logFile = New-Item -itemType File -Path $logFolderPath -Name $logFileName

        WriteToLogFile("Accessing "+$userMailbox)
        #$exchangeOnline = [Microsoft.Exchange.WebServices.Data.ExchangeService]::new()
        #$exchangeOnline.UseDefaultCredentials = $false
        #$exchangeOnline.Credentials = $Service.Credentials
        #$exchangeOnline.Url = "https://outlook.office365.com/EWS/Exchange.asmx"

        # use credentials to access mailbox of user 
        $Service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $userMailbox)

        $contactsFolder = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Contacts)

        $pfContactsFolder = "\CompanyInfo\Associates"

        $pfRootFolder = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::PublicFoldersRoot)

        $pfSource = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($Service, $pfRootFolder)
        
        $pfFolderPath = $pfContactsFolder.Split("\")

        # get the exact path for the folder spcified in $spContactsFolder
        for ($i = 1; $i -lt $pfFolderPath.length; $i++)
        {
            $folderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(1)
            $searchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName, $pfFolderPath[$i])
            $findFolderResults = $Service.FindFolders($pfSource.Id, $searchFilter, $folderView)
            if ($findFolderResults.TotalCount -gt 0)
            {
                $pfSource = $findFolderResults.Folders[0]
                WriteToLogFile("Found  "+$pfContactsFolder)
            }
            else
            {
                WriteToLogFile($pfContactsFolder+" Not Found")
                exit
            }
        }

        $userItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView(10000)
        $userFindItemResults = $Service.FindItems($contactsFolder, $userItemView)

        # delete all contacts from the mailbox that match the parameters
        WriteToLogfile("Deleting AzTec Consultants contacts...")
        $count = 1
        foreach ($userItem in $userFindItemResults.Items | Where-Object {$_.CompanyName -eq "AzTec Consultants, Inc"})
        {
            WriteToLogFile($userItem.DisplayName+" "+$count)
            $userItem.Delete([Microsoft.Exchange.WebServices.Data.DeleteMode]::HardDelete)
            $count++
        }
        WriteToLogfile("Deleting AzTec S&L contacts...")
        $count = 1
        foreach ($userItem in $userFindItemResults.Items | Where-Object {$_.CompanyName -eq "AzTec Surveying & Locating"})
        {
            WriteToLogFile($userItem.DisplayName+" "+$count)
            $userItem.Delete([Microsoft.Exchange.WebServices.Data.DeleteMode]::HardDelete)
            $count++
        }

        $pfItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView(10000)
        $pfFindItemResults = $Service.FindItems($pfSource.Id, $pfItemView)

        # copy the contents of of the folder specified in $pfContacts to the mailbox Contacts folder
        WriteToLogfile("Copying All Associate contacts...")
        $count = 1
        foreach ($pfItem in $pfFindItemResults.Items)
        {
            WriteToLogFile($pfItem.DisplayName+" "+$count)
            $pfItem.Copy($contactsFolder)
            $count++
        }
}


