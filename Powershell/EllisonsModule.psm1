#Module
Write-Host "Loading Powershell Ellisons Module" -ForegroundColor Green
Write-Host "Version 1.0.0" 
#######################################################################
#             Check AzureAD Module - Install If Missing               #
#######################################################################
Set-Location -Path $PSScriptRoot
$AzureAD = "AzureAD"

$Installedmodules = Get-InstalledModule

if ($Installedmodules.name -contains $AzureAD) {

    "$AzureAD is installed "

}

else {

    Install-Module AzureAD

    "$AzureAD now installed"

}

#######################################################################
#              Check MSOnline Module - Install If Missing             #
#######################################################################

$MSOnline = "MSOnline"

$Installedmodules = Get-InstalledModule

if ($Installedmodules.name -contains $MSOnline) {

    "$MSOnline is installed "

}

else {

    Install-Module MSOnline

    "$MSOnline now installed"

}

Set-Location -Path $PSScriptRoot
$Global:currentUser = $env:UserName
function Show-Menu { 
    param ( 
        [string]$Title = 'Procedures' 
    ) 
    Clear-Host 
    Write-Host "================ $Title ================" 
    Write-Host "1:  Press '1' for Full Access." 
    Write-Host "2:  Press '2' for Remove Access." 
    Write-Host "3:  Press '3' for Send On Behalf."
    Write-Host "4:  Press '4' for View Send on Behalf Permissions." 
    Write-Host "5:  Press '5' for the new user procedure." 
    Write-Host "6:  Press '6' for the user left procedure." 
    Write-Host "7:  Press '7' to connect to 365." 
    Write-Host "8:  Press '8' to select Disable Out Of Office." 
    Write-Host "9:  Press '9' to sync all AD Controllers."
    Write-Host "10: Press '10' to Unlock AD Accounts and sync"
    Write-Host "U:  Press 'U' to update stored creds in O365 file."
    Write-Host "D:  Press 'D' to update Domain Admin creds in DomainAdmin file." 
    Write-Host "Q:  Press 'Q' to quit." 
}

function Start-UnlockedADAccounts {
    Import-Module ActiveDirectory
    $UnlockedUsers = (Search-ADAccount -LockedOut | Unlock-ADAccount)
    if($UnlockedUsers) {Write-Host '$UnlockedUsers'; Start-SyncAD} else {Write-Host "Found No Locked Out Accounts, Returning to menu" -ForegroundColor Green}

}

function Enter-OnPrem365 {
    Import-Module ActiveDirectory
    Write-Output "Importing OnPrem Exchange Module"
    $OnPrem = New-PSSession -Authentication Kerberos -ConfigurationName Microsoft.Exchange -ConnectionUri 'http://ez-az-exchb.ellisonslegal.com/Powershell' -Credential $Global:AdminCred
    Import-Module MSOnline
    Import-PSSession $OnPrem | Out-Null
}

Function Enter-Office365 {
    Import-Module ActiveDirectory
    Import-Module MSOnline
    Set-CredsUp
    Connect-MsolService -Credential $Global:365Cred
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid -Credential $Global:365Cred -Authentication Basic -AllowRedirection
    Import-PSSession $Session -AllowClobber 
}

function Set-CredsUp {

    $Global:365AdminPassword = Get-Content ".\creds\$Global:currentUser-O365Password.txt"
    $Global:365AdminUsername = Get-Content ".\creds\$Global:currentUser-O365AdminName.txt"
    $Global:SRVAdminPassword = Get-Content ".\creds\$Global:currentUser-SRVPassword.txt"
    $Global:SRVAdminUsername = Get-Content ".\creds\$Global:currentUser-SRVAdminName.txt"
    $Global:LocalAdminPassword = Get-Content ".\creds\$Global:currentUser-AdministratorPassword.txt"
    $Global:LocalAdminUsername = Get-Content ".\creds\$Global:currentUser-AdministratorName.txt"

}

function Start-UserLeft {
    & './UserLeft.ps1'
}
function Start-NewUser {
    & './NewUser.ps1'
}
function Start-FullAccess {
    Enter-Office365
    Clear-Host
    $Requestee = $null
    $Target = $null
    $Requestee = Read-Host "Who do you want to have access?"
    Write-Host
    $Target = Read-Host "Whos Inbox do they need access to?"
    Add-MailboxPermission -Identity $Target -User $Requestee -AccessRights FullAccess -InheritanceType All -Automapping:$true
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
function Start-SyncAD {
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

Invoke-Command -ComputerName ez-az-dc01 -ScriptBlock { Start-ADSyncSyncCycle -PolicyType Delta }
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