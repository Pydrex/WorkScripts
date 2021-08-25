#Module

Write-Host "Loading Powershell Ellisons Module" -BackgroundColor Black -ForegroundColor Green
Write-Host "Version 4" -BackgroundColor Black -ForegroundColor Green
Write-Host "Created and Maintained by Andrew Powell" -BackgroundColor Black -ForegroundColor Green
Write-Host "Updated 25/08/2021 - 14:30" -BackgroundColor Black -ForegroundColor Green

#######################################################################
#             Check AzureAD Module - Install If Missing               #
#######################################################################
Set-Location -Path $PSScriptRoot
$AzureAD = "AzureAD"

$Installedmodules = Get-InstalledModule

if ($Installedmodules.name -contains $AzureAD) {

    #Update-Module $AzureAD -Confirm:$False
    "$AzureAD is installed "

}

else {

    Install-Module AzureAD -Confirm:$False -Force -AllowClobber

    "$AzureAD now installed"

}

#######################################################################
#              Check MSOnline Module - Install If Missing             #
#######################################################################

$MSOnline = "MSOnline"

$Installedmodules = Get-InstalledModule

if ($Installedmodules.name -contains $MSOnline) {

    "$MSOnline is installed "
    #Update-module MSOnline -Confirm:$False

}

else {

    Install-Module MSOnline -Confirm:$False -Force

    "$MSOnline now installed"

}

##

$ExchangeOnlineManagement = "ExchangeOnlineManagement"

$Installedmodules = Get-InstalledModule

if ($Installedmodules.name -contains $ExchangeOnlineManagement) {

    #Update-Module -Name ExchangeOnlineManagement -Confirm:$False
    "$ExchangeOnlineManagement is installed "

}

else {

    Install-Module -Name ExchangeOnlineManagement -Confirm:$False -Force

    "$ExchangeOnlineManagement now installed"

}

$SPOService = "Microsoft.Online.SharePoint.PowerShell"

if ($Installedmodules.name -contains $SPOService) {

    #Update-Module -Name ExchangeOnlineManagement -Confirm:$False
    "$SPOService is installed "

}

else {

    Install-Module -Name Microsoft.Online.SharePoint.PowerShell -Confirm:$False -Force

    "$SPOService now installed, You may need to download the MSi from the site"

}

#requires -module ActiveDirectory
#requires -module MSOnline

Set-Location -Path $PSScriptRoot
$Global:currentUser = $env:UserName

function Start-UnlockedADAccounts {
    Import-Module ActiveDirectory
    $Search = (Search-ADAccount -lockedout | Select-Object SamAccountName)
    $LockedUsers = $null
    if($search) {$LockedUsers = $Search | Select-Object -ExpandProperty SamAccountName}
    if($LockedUsers) {Write-Host "Unlocking User(s): $LockedUsers"; Start-SyncAD; $LockedUsers | Unlock-AdAccount } else {Write-Host "Found No Locked Out Accounts, Returning to menu" -ForegroundColor Green}
}

function Start-PWDReset {
    Import-Module ActiveDirectory
    $User = $null
        Do {
                $User = Read-Host "Enter in the USERNAME of the person you wish to reset password expiry on"
                Write-Host
        
                
                Write-Host "Checking if $User is a valid user..." -ForegroundColor:Green
                If ($(Get-ADUser -Filter { SamAccountName -eq $User })) {
                    Write-Host "Found user: Is this correct?" (Get-ADUser $User | Select-Object -ExpandProperty DistinguishedName)
                    Write-Host
        
                    $Proceed = Read-Host "Continue? (y/n)"
                    Write-Host
        
        
                    if ($Proceed -ieq 'y') {
                        Get-ADUser $User | .\pwdset.ps1
                        $Exit = $true
                    }
        
                }
                else {
                    Write-Host "$User was not a valid user" -ForegroundColor:Red
                    Start-Sleep 2
                    $Exit = $false
                    Clear-Host
                }
        
        } until ($Exit -eq $true)
}

Function Enter-Office365 {

    if (!(Get-PSSession | Where { $_.ConfigurationName -eq "Microsoft.Exchange" })) { 
        Get-PSSession | Remove-PSSession
        Import-Module ActiveDirectory
        Import-Module MSOnline
        Import-Module Microsoft.Online.SharePoint.PowerShell
        #Import-Module ExchangeOnlineManagement
        Clear-Host
        Set-CredsUp
        Connect-MsolService -Credential $Global:365Cred
        Connect-SPOService -Url https://ellisonssolicitors-admin.sharepoint.com -Credential $Global:365Cred
        $Session = Connect-ExchangeOnline -UserPrincipalName $Global:365AdminUsername -Credential $Global:365Cred
        #$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid -Credential $Global:365Cred -Authentication Basic -AllowRedirection
    }
    if ($Session) {Import-PSSession $Session -AllowClobber}
}
function Set-CredsUp {

    $Global:365AdminPassword = Get-Content ".\creds\$Global:currentUser-O365Password.txt"
    $Global:365AdminUsername = Get-Content ".\creds\$Global:currentUser-O365AdminName.txt"
    $Global:SRVAdminPassword = Get-Content ".\creds\$Global:currentUser-SRVPassword.txt"
    $Global:SRVAdminUsername = Get-Content ".\creds\$Global:currentUser-SRVAdminName.txt"
    $Global:LocalAdminPassword = Get-Content ".\creds\$Global:currentUser-AdministratorPassword.txt"
    $Global:LocalAdminUsername = Get-Content ".\creds\$Global:currentUser-AdministratorName.txt"

}

function CopyUserGroups {
    Write-host 'WARNING THIS WILL OVERWRITE ALL USER GROUPS OF THAT USER'
    $CopyFROM = Read-Host "Enter the logon name of the user you wish to copy FROM!"
    $CopyTO = Read-Host "Enter the logon name of the user you wish to copy TO!"
    Get-ADUser -Identity $CopyFROM  -Properties memberof | Select-Object -ExpandProperty memberof | Add-ADGroupMember -Members $CopyTO | Write-Output $true
    'Done'
    
}
function Start-365Menu {
    do {
        Clear-Host
        [string]$Title = 'Office 365 Menu' 
        Clear-Host 
        Write-Host "================ $Title ================" 
        Write-Host "1:  Press '1' for Full Access." 
        Write-Host "2:  Press '2' for Remove Access." 
        Write-Host "3:  Press '3' for Send On Behalf."
        Write-Host "4:  Press '4' for View Send on Behalf Permissions." 
        Write-Host "5:  Press '5' to connect to 365." 
        Write-Host "6:  Press '6' to select Disable Out Of Office."
        Write-Host "7:  Press '7' to enable MFA and OWA."
        Write-Host "8:  Press '8' to disable OWA."
        Write-Host "9:  Press '9' to add to Dementia Friends signature."
        Write-Host "10: Press '10' to remove  Dementia Friends signature."
        Write-host "11: Press '11' to run New starter fix for calendar and send as permissions"
        Write-host "12: Press '12' to copy AD Groups from one user to another"
        Write-Host "R:  Press 'R' to return to the previous menu." 
        $input = Read-Host "Please make a selection" 
        switch ($input) { 
             '1' { 
                  Clear-Host
                  'You chose the full access procedure' 
                  Start-FullAccess
             } '2' { 
                  Clear-Host 
                  'You chose the Remove access procedure'
                  Start-RemoveAccess
             } '3' { 
                  Clear-Host 
                  'You chose the Send On Behalf procedure'
                  Start-SendOnBehalf
             } '4' { 
                  Clear-Host 
                  'You chose the view access procedure'
                  Start-AccessBehalf
             } '5' { 
                  Clear-Host 
                  'You chose Office 365 connection'
                  Enter-Office365
             } '6' { 
                  Clear-Host 
                  'You Selected the Out Of Office Procedure Office 365 connection'
                  Start-DisableOutOfOffice
             } '7' { 
                    Clear-Host 
                    'You selected Enable OWA and MFA'
                    Start-EnableOWA  
             } '8' { 
                    Clear-Host 
                    'You selected Disable OWA'
                    Start-DisableOWA 
             } '9' { 
                Clear-Host 
                'You selected add to Dementia Friends'
                addToDF  
             } '10' { 
                Clear-Host 
                'You selected remove from Dementia Friends'
                removeFromDF
             } '11' {
                Clear-Host
                'Running new starter fixes'
                Start-NewStarterFix
             } '12' {
                Clear-Host
                CopyUserGroups
             } 'R' { 
                  return 
             } 
        }
        pause
        Clear-Host
    } 
    until ($input -eq 'R')
}
function Start-PaperCutIDCheck {

    do { 
        Write-Host "1: Press '1' for set someones Papercut ID." 
        Write-Host "2: Press '2' to check if Papercut ID in use"
        Write-Host "3: Press '3' to find a users ID with their email"
        Write-Host "3: Press 'R' to Return to previous menus"
        $input = Read-Host "Please make a selection" 
        switch ($input) { '1' {
                Clear-Host
                $finding = $null
                $EmployeeID = $null
                $email = $null
                $email = Read-Host "Enter the email address of the User"
                  $EmployeeID = Read-Host "Enter the Papercut Code to set"
                  $finding = Get-ADUser -Filter { EmployeeId -eq $EmployeeID } -Properties EmployeeId
                  Clear-Host
                  if ($finding) {
                      $idLookup = $finding | ForEach-Object { $idLookup = @{ } } { if ($_.EmployeeId) { $idLookup[$_.EmployeeId] += 1 } } { $idLookup }
                      $filteredUsers = $finding | Where-Object { if ($_.EmployeeId) { $idLookup[$_.EmployeeId] -gt 1 } }
                      $report = $finding | Select-Object -Property SamAccountName, EmployeeId
                      Write-Host "!DUPLICATE PAPERCUT ID FOUND! - PLEASE CHANGE" -ForegroundColor:Red
                      $report | Format-Table | Out-String|% {Write-Host $_ -BackgroundColor:Yellow -ForegroundColor:Black}
                  } else {
                      $sam = Get-ADUser -Filter { emailaddress -eq $email } -Properties SamAccountName
                      $Server = "EZ-AZ-DC01.Ellisonslegal.com"
                      Get-ADUser $sam -Server $Server | Set-ADUser -EmployeeID $EmployeeID
                      Write-Host "Set PapercutID for $sam" -ForegroundColor:Green
                  }
             } '2' {
                Clear-Host
                $finding = $null
                $EmployeeID = $null
                $EmployeeID = Read-Host "Enter a Papercut ID to check"
                $finding = Get-ADUser -Filter { EmployeeId -eq $EmployeeID } -Properties EmployeeId
                if ($finding) {
                    $idLookup = $finding | ForEach-Object { $idLookup = @{ } } { if ($_.EmployeeId) { $idLookup[$_.EmployeeId] += 1 } } { $idLookup }
                    $filteredUsers = $finding | Where-Object { if ($_.EmployeeId) { $idLookup[$_.EmployeeId] -gt 1 } }
                    $report = $finding | Select-Object -Property SamAccountName, EmployeeId
                    Write-Host "!PAPERCUT ID FOUND!" -ForegroundColor:Red -BackgroundColor:Yellow
                    $report | Format-Table | Out-String|% {Write-Host $_ -BackgroundColor:Yellow -ForegroundColor:Black}
                } else {
                    Write-Host "Papercut ID is not in use" -ForegroundColor:Green
                }
            } '3' {
                Clear-Host
                $email = $null
                $EmployeeID = $null
                $email = Read-Host "Enter the email address of the User"
                $finding = Get-ADUser -Filter { emailaddress -eq $email } -Properties EmployeeId
                $report = $finding | Select-Object -Property SamAccountName, EmployeeId
                $report | Format-Table | Out-String|% {Write-Host $_ -BackgroundColor:Green -ForegroundColor:Black}
            } 'R' {
                return
            }
        } 
        pause
        Clear-Host
    } 
    until ($input -eq 'R')
}

function Start-UserLeft {
    & './UserLeft.ps1'
}
function Start-NewUser {
    & './NewUser.ps1'
}

function Start-NewStarterFix {
    Param
    (
         [Parameter(Mandatory=$false, Position=1)]
         [string] $email
    )
    Enter-Office365

    Clear-Host
    if (!$email) {
        $email = Read-Host "Please type in the users email address to turn off OWA and Set Calendar Defaults"
    }
    Write-Host 'I RAN'
    
    Set-CASMailbox -Identity $email -ActiveSyncEnabled $false -OWAforDevicesEnabled $false -OWAEnabled $false
    Set-Mailbox -Identity $email -MessageCopyForSentAsEnabled $True
    Set-Mailbox -Identity $email -MessageCopyForSendOnBehalfEnabled $True
    #Set Perms for Calendars
    $cal = $email + ":\Calendar"
    Set-MailboxFolderPermission $cal -User Default -AccessRights Reviewer
}
function Start-Egg {
    & './Snake.ps1'
}

function addToDF {
    $USERNAME = Read-Host "Enter the USERNAME of who you want to have the DF signature"
    Get-ADUser $USERNAME
    Set-ADUser -Identity $USERNAME -Replace @{extensionAttribute10="DF"}
}

function removeFromDF {
    $USERNAME = Read-Host "Enter the USERNAME of who you want to remove the DF signature from"
    Get-ADUser $USERNAME
    Set-ADUser -Identity $USERNAME -Clear "extensionAttribute10"
}
function Start-UpdatePhoneList{
    & './UpdatePhoneLists.ps1'
    
}

function Start-FullAccess {
    Enter-Office365
    $Requestee = $null
    $Target = $null
    $Requestee = Read-Host "Who do you want to have access?"
    Write-Host
    $Target = Read-Host "Whos Inbox do they need access to?"
    $AutoMap = Read-host "Do you want this to AutoMap to Outlook? (If n they will have to file > open > other user) (y/n)"
    if ($AutoMap -ieq 'y') {
    Add-MailboxPermission -Identity $Target -User $Requestee -AccessRights FullAccess -InheritanceType All -Automapping:$true} else {
    Add-MailboxPermission -Identity $Target -User $Requestee -AccessRights FullAccess -InheritanceType All -Automapping:$false
    }
    Write-Host "Done, This might take 20 minutes."
}
function Start-RemoveAccess {
    Enter-Office365
    Clear-Host
    $Requestee = $null
    $Target = $null
    $Requestee = Read-Host "Who do you want remove from having access?"
    Write-Host
    $Target = Read-Host "Whos Inbox do they need removing from?"
    Remove-MailboxPermission -Identity $Target -User $Requestee -AccessRights FullAccess -InheritanceType All
    Write-Host "Done, This might take 20 minutes."
}
function Start-SendOnBehalf {
    Enter-Office365
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
function Start-AccessBehalf {
    Enter-Office365
    Clear-Host
    $Requestee = $null
    $Requestee = Read-Host "Whos mailbox do you want to check permissions on?"
    Get-Mailbox $Requestee | Format-Table Name, grantsendonbehalfto -wrap
}

# Sets the MFA requirement state
function Set-MfaState {

    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipelineByPropertyName=$True)]
        $ObjectId,
        [Parameter(ValueFromPipelineByPropertyName=$True)]
        $UserPrincipalName,
        [ValidateSet("Disabled","Enabled","Enforced")]
        $State
    )

    Process {
        Write-Verbose ("Setting MFA state for user '{0}' to '{1}'." -f $ObjectId, $State)
        $Requirements = @()
        if ($State -ne "Disabled") {
            $Requirement =
                [Microsoft.Online.Administration.StrongAuthenticationRequirement]::new()
            $Requirement.RelyingParty = "*"
            $Requirement.State = $State
            $Requirements += $Requirement
        }

        Set-MsolUser -ObjectId $ObjectId -UserPrincipalName $UserPrincipalName `
                     -StrongAuthenticationRequirements $Requirements
    }
}

function Start-EnableRDS {
    Enter-Office365
    Clear-Host
    $email = $null
    $email = Read-Host "Whos email address do you want to enable Remote Desktop and MFA for??"
    $sam = Get-ADUser -Filter { emailaddress -eq $email } -Properties SamAccountName
    Add-ADGroupMember -Identity "WVDUsers" -Members $sam
    Add-ADGroupMember -Identity "MFA Users" -Members $sam
    Set-MfaState -UserPrincipalName $email -State Enforced
    Set-CASMailbox -Identity $email -OWAEnabled $true
   
    $mfaenabled = Get-MsolUser -UserPrincipalName $email | select UserPrincipalName, `
        @{Name = 'MFAEnabled'; Expression={if ($_.StrongAuthenticationRequirements) {Write-Output $true} else {Write-Output $false}}}
        Write-Output $mfaenabled | Sort-Object MFAEnabled
    
    $owaenabled = Get-CASMailbox -Identity $email | Select-Object Identity, `
        @{Name = 'OWAisEnabled'; Expression={if ($_.OWAEnabled) {Write-Output $true} else {Write-Output $false}}}
        Write-Output $owaenabled | Sort-Object OWAisEnabled
    }

function Start-EnableOWA {
    Enter-Office365
    Clear-Host
    $email = $null
    $email = Read-Host "Whos email address do you want to enable OWA and MFA on?"
    $sam = Get-ADUser -Filter { emailaddress -eq $email } -Properties SamAccountName
    Add-ADGroupMember -Identity "MFA Users" -Members $sam
    Set-MfaState -UserPrincipalName $email -State Enforced
    Set-CASMailbox -Identity $email -OWAEnabled $true
   
    $mfaenabled = Get-MsolUser -UserPrincipalName $email | select UserPrincipalName, `
        @{Name = 'MFAEnabled'; Expression={if ($_.StrongAuthenticationRequirements) {Write-Output $true} else {Write-Output $false}}}
        Write-Output $mfaenabled | Sort-Object MFAEnabled
    
    $owaenabled = Get-CASMailbox -Identity $email | Select-Object Identity, `
        @{Name = 'OWAisEnabled'; Expression={if ($_.OWAEnabled) {Write-Output $true} else {Write-Output $false}}}
        Write-Output $owaenabled | Sort-Object OWAisEnabled
    }

function Start-DisableOWA {
        Enter-Office365
        Clear-Host
        $email = $null
        $email = Read-Host "Whos email address do you want to disable OWA and MFA on?"
        $sam = Get-ADUser -Filter { emailaddress -eq $email } -Properties SamAccountName
        Remove-ADGroupMember -Identity "MFA Users" -Members $sam -Confirm:$false
        Set-MfaState -UserPrincipalName $email -State Disabled
        Set-CASMailbox -Identity $email -ActiveSyncEnabled $false -OWAforDevicesEnabled $false -OWAEnabled $false
        
        $mfaenabled = Get-MsolUser -UserPrincipalName $email | select UserPrincipalName, `
        @{Name = 'MFAEnabled'; Expression={if ($_.StrongAuthenticationRequirements) {Write-Output $true} else {Write-Output $false}}}
        Write-Output $mfaenabled | Sort-Object MFAEnabled

        $owaenabled = Get-CASMailbox -Identity $email | Select-Object Identity, `
            @{Name = 'OWAisEnabled'; Expression={if ($_.OWAEnabled) {Write-Output $true} else {Write-Output $false}}}
            Write-Output $owaenabled | Sort-Object OWAisEnabled

        
}
    
function Start-SyncAD {
    if (!$Global:LocalAdminUsername -OR !$Global:LocalAdminPassword) { Start-AdministratorUpdate }
    $AdminName = "Administrator"
    $Global:LocalAdminPassword = Get-Content ".\creds\$Global:currentUser-AdministratorPassword.txt" | ConvertTo-SecureString
    $Global:AdminCred = new-object -typename System.Management.Automation.PSCredential -argumentlist $AdminName, $Global:LocalAdminPassword
    $s = New-PSSession -computerName ez-az-dc01 -Credential $Global:AdminCred
    $Sess = Enter-PSSession $s
    Invoke-Command -ComputerName EZ-AZ-DC01 -Credential $Global:AdminCred -Scriptblock {Start-ScheduledTask -TaskName "SyncAll"}
    Exit-PSSession
    Write-Host "SyncAll started on DC01, wait 1 minute for sync"
}

function Start-DisableOutOfOffice {
    Enter-Office365
    Clear-Host
    $Requestee = $null
    $Requestee = Read-Host "Whos mailbox do you want to remove the Out Of Office for?"
    Set-MailboxAutoReplyConfiguration -Identity $Requestee -AutoReplyState Disabled
}
function Start-UpdateCreds {
    Write-Host "Enter your 365 username and password" -ForegroundColor Green
    Read-Host "Enter your Office 365 Admin email (First.Last@ellisonssolcitiors.com) etc..." | Out-File ".\creds\$Global:currentUser-O365AdminName.txt"
    Write-Host "Type in your Office 365 login email and password" -ForegroundColor Blue -BackgroundColor Black
    $UpdateCredscredential = Get-Credential
    $UpdateCredscredential.Password | ConvertFrom-SecureString | Out-File ".\creds\$Global:currentUser-O365Password.txt"
    $Global:365AdminPassword = Get-Content ".\creds\$Global:currentUser-O365Password.txt"
    $Global:365AdminUsername = Get-Content ".\creds\$Global:currentUser-O365AdminName.txt"
}
function Start-UpdateSRVCreds {
    Write-Host "Enter the SRV username and password" -ForegroundColor Green
    $UpdateDomainCredscredential = Get-Credential
    Read-Host "Enter your SRV Account Username (ELLNET\***-SRV)" | Out-File ".\creds\$Global:currentUser-SRVAdminName.txt"
    $UpdateDomainCredscredential.Password | ConvertFrom-SecureString | Out-File ".\creds\$Global:currentUser-SRVPassword.txt"
    $Global:SRVAdminPassword = Get-Content ".\creds\$Global:currentUser-SRVPassword.txt"
    $Global:SRVAdminUsername = Get-Content ".\creds\$Global:currentUser-SRVAdminName.txt"

}

function Start-AdministratorUpdate {
    Write-Host "Enter the ELLNET\Administrator username and password" -ForegroundColor Green
    $AdministratorUpdatecredential = Get-Credential
    $AdminNameLocal = "Administrator" | Out-File ".\creds\$Global:currentUser-AdministratorName.txt"
    $AdministratorUpdatecredential.Password | ConvertFrom-SecureString | Out-File ".\creds\$Global:currentUser-AdministratorPassword.txt"

    $Global:LocalAdminPassword = Get-Content ".\creds\$Global:currentUser-AdministratorPassword.txt"
    $Global:LocalAdminUsername = Get-Content ".\creds\$Global:currentUser-AdministratorName.txt"
}

function Start-EXCHAdministratorUpdate {
    Write-Host "Enter the ELLNET\EXCHAdmin username and password (Same as Domain)" -ForegroundColor Green
    $EXCHAdminUpdatecredential = Get-Credential
    $EXCHAdminLocal = "EXCHAdmin" | Out-File ".\creds\$Global:currentUser-EXCHAdminName.txt"
    $EXCHAdminUpdatecredential.Password | ConvertFrom-SecureString | Out-File ".\creds\$Global:currentUser-EXCHAdminPassword.txt"

    $Global:EXCHAdminPassword = Get-Content ".\creds\$Global:currentUser-EXCHAdminPassword.txt"
    $Global:EXCHAdminUsername = Get-Content ".\creds\$Global:currentUser-EXCHAdminName.txt"
}