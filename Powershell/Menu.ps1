##VERSION 2.0  - AP - 14/02/2020
$FormatEnumerationLimit = -1
Set-Location -Path $PSScriptRoot

Function Connect365 {
     Set-ExecutionPolicy Unrestricted -Force 
     Import-Module MSOnline
     Get-PSSession | Remove-PSSession
     #$credential = Get-Credential
     #$credential.Password | ConvertFrom-SecureString | Out-File O365Account.txt
     if (!$AdminName) {$AdminName = Read-Host "Enter your Office 365 Admin email (First.Last@ellisonssolcitiors.com) etc..."}
     Import-Module MSOnline
     $Pass = Get-Content "O365Account.txt" | ConvertTo-SecureString
     $Cred = new-object -typename System.Management.Automation.PSCredential -argumentlist $AdminName, $Pass
     Connect-MsolService -Credential $Cred
     $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid -Credential $cred -Authentication Basic -AllowRedirection
     Import-PSSession $Session -AllowClobber 

}

function Left {
     Clear-Host
     & '.\UserLeft.ps1'
}

function NewUser {
     Clear-Host
     & '.\NewUser.ps1'
}

function FullAccess {
     Connect365
     Clear-Host
     $Requestee = $null
     $Target = $null
     $Requestee = Read-Host "Who do you want to have access?"
     Write-Host
     $Target = Read-Host "Whos Inbox do they need access to?"
     Add-MailboxPermission -Identity $Target -User $Requestee -AccessRights FullAccess -InheritanceType All -Automapping:$true
     Write-Host "Done, This might take 20 minutes."
}

function RemoveAccess {
     Connect365
     Clear-Host
     $Requestee = $null
     $Target = $null
     $Requestee = Read-Host "Who do you want remove from having access?"
     Write-Host
     $Target = Read-Host "Whos Inbox do they need removing from?"
     Remove-MailboxPermission -Identity $Target -User $Requestee -AccessRights FullAccess -InheritanceType All
     Write-Host "Done, This might take 20 minutes."
}

function SendOnBehalf {
     Connect365
     Clear-Host
     $Requestee = $null
     $Target = $null
     Write-Host "Example - Type Andrew Powell if he wants to send on behalf of someone"
     $Requestee = Read-Host "Who do you want to grant send on behalf to?"
     Write-Host
     $Target = Read-Host "Who do they send on behalf of?"
     Set-mailbox $Target -Grantsendonbehalfto @{add = $Requestee }
     Write-Host "Done, This might take 20 minutes."
}

function AccessBehalf {
     Connect365
     Clear-Host
     $Requestee = $null
     $Requestee = Read-Host "Whos mailbox do you want to check permissions on?"
     Get-Mailbox $Requestee | Format-Table Name, grantsendonbehalfto -wrap
}

function syncAD {
     Get-PSSession | Remove-PSSession

     $DomainControllers = Get-ADDomainController -Filter *
     ForEach ($DC in $DomainControllers.Name) {
          Write-Host "Processing for "$DC -ForegroundColor Green
          If ($Mode -eq "ExtraSuper") {
               REPADMIN /kcc $DC
               REPADMIN /syncall /A /e /q $DC
          }
          Else {
               REPADMIN /syncall $DC "DC=Ellisonslegal,DC=com" /d /e /q
          }
     }
}

function DisableOutOfOffice {
     Connect365
     Clear-Host
     $Requestee = $null
     $Requestee = Read-Host "Whos mailbox do you want to remove the Out Of Office for?"
     Set-MailboxAutoReplyConfiguration -Identity $Requestee -AutoReplyState Disabled
}

function Show-Menu { 
     param ( 
          [string]$Title = 'Procedures' 
     ) 
     Clear-Host 
     Write-Host "================ $Title ================" 
     
     Write-Host "1: Press '1' for Full Access." 
     Write-Host "2: Press '2' for Remove Access." 
     Write-Host "3: Press '3' for Send On Behalf."
     Write-Host "4: Press '4' for View Send on Behalf Permissions." 
     Write-Host "5: Press '5' for the new user procedure." 
     Write-Host "6: Press '6' for the user left procedure." 
     Write-Host "7: Press '7' to connect to 365." 
     Write-Host "8: Press '8' to select Disable Out Of Office." 
     Write-Host "8: Press '9' to sync all AD Controllers." 
     Write-Host "Q: Press 'Q' to quit." 
} 
do { 
     Show-Menu 
     $input = Read-Host "Please make a selection" 
     switch ($input) { 
          '1' { 
               Clear-Host
               'You chose the full access procedure' 
               FullAccess
          } '2' { 
               Clear-Host 
               'You chose the Remove access procedure'
               RemoveAccess
          }
          '3' { 
               Clear-Host 
               'You chose the Send On Behalf procedure'
               SendOnBehalf
          }
          '4' { 
               Clear-Host 
               'You chose the view access procedure'
               AccessBehalf
          }
          '5' { 
               Clear-Host 
               'You chose the New user procedure'
               NewUser
          }
          '6' { 
               Clear-Host 
               'You chose Employee Left procedure'
               Left
          }
          '7' { 
               Clear-Host 
               'You chose Office 365 connection'
               Connect365
          } '8' { 
               Clear-Host 
               'You Selected the Out Of Office Procedure Office 365 connection'
               DisableOutOfOffice
          } '9' { 
               Clear-Host 
               'You Selected the sync AD procedure'
               syncAD
           
          } 'q' { 
               'Thank you, come again'
               return 
          } 
     } 
     pause 
} 
until ($input -eq 'q')



# Import-Module AzureAD


# Add additonal users to Send on Behalf permissions for mailbox. add= list if a comma seperate list. Each email address should be in double quoted brackets

# Set-mailbox -MySharedMailbox- -Grantsendonbehalfto @{add="john.smith@domain.com"}

# Confirm that user has been succesfully added to send on behalf permissions for mailbox

# Get-Mailbox 'MySharedMailbox' | ft Name,grantsendonbehalfto -wrap

# Display exit script (to keep window open in order to view the above)

# Read-Host -Prompt "Press Enter to exit"

## Set-MailboxAutoReplyConfiguration "" -AutoReplyState enabled -ExternalAudience all -InternalMessage "Message" -ExternalMessage "Message"
# Set-UnifiedGroup -Identity "IT" -HiddenFromAddressListsEnabled $true

#Remove-MailboxPermission -Identity "" -User "" -AccessRights FullAccess -InheritanceType All


##set-msoluserprincipalname -newuserprincipalname "Emma.Emerson@ellisonssolicitors.com" -userprincipalname "Emma.Closs@ellisonssolicitors.com"

##Set-Mailbox -Identity "Joe Healy" -IssueWarningQuota 24.5gb -ProhibitSendQuota 24.75gb -ProhibitSendReceiveQuota 25gb -UseDatabaseQuotaDefaults $false

##Get-MobileDevice -Resultsize Unlimited | Select-Object Identity, DeviceID, FriendlyName, DeviceImei, DeviceMobileOperator, DeviceOS, DeviceType, UserDisplayName | Export-CSV C:\Powershell\MobileDevices\ActiveSyncDevicesOnCloud.csv

##$DeviceID = "51f3e3ef066235e386ed036ccb27a64a"
##$identity = "Gabriel Sarateanu"
##Set-CASMailbox "Gulcan Dirlik" -ActiveSyncAllowedDeviceIDs @{Add="3F634263D6094C32A2050E4E6D66EDF8"}
##Get-CASMailbox $alias | Select ActiveSyncAllowedDeviceIDs,ActiveSyncBlockedDeviceIDs
##Set-CASMailbox "Gabriel Sarateanu" -EwsBlockList @{Add="Outlook-iOS/*","Outlook-Android/*"}

#Enable - Save a copy of sending mail items in the Shared mailbox sent items folder
#Set-Mailbox -Identity <identity> -MessageCopyForSentAsEnabled $True
#Set-Mailbox -Identity <identity> -MessageCopyForSentAsEnabled $False
#Set-Mailbox -Identity <identity> -MessageCopyForSendOnBehalfEnabled $True
#Set-Mailbox -Identity <identity> -MessageCopyForSendOnBehalfEnabled $False
#Add-RecipientPermission "REQUESTED" -AccessRights SendAs -Trustee "Requestee"

#Get-MsolUser | Where-Object {$_.licenses[0].AccountSku.SkuPartNumber -eq ($acctSKU).Substring($acctSKU.IndexOf(":")+1, $acctSKU.Length-$acctSKU.IndexOf(":")-1) -and $_.IsLicensed -eq $True} | Set-MsolUserLicense -LicenseOptions $x