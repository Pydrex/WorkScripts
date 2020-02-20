#Module
Write-Host "Loading Powershell Ellisons Module" -ForegroundColor Green
Write-Host "Version 1.0.0" 
Set-Location -Path $PSScriptRoot

function Get-WelcomeMessage {
    Write-Host "Hello World"    
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
    Write-Host "9: Press '9' to sync all AD Controllers."
    Write-Host "U: Press 'U' to update stored creds in O365 file."
    Write-Host "D: Press 'D' to update Domain Admin creds in DomainAdmin file." 
    Write-Host "Q: Press 'Q' to quit." 
} 

Function Enter-Office365 {
    Import-Module MSOnline
    Get-PSSession | Remove-PSSession
    if (!$365Admin) { $365Admin = Read-Host "Enter your Office 365 Admin email (First.Last@ellisonssolcitiors.com) etc..." }
    Import-Module MSOnline
    $365Pass = Get-Content "O365Account.txt" | ConvertTo-SecureString
    $365Cred = new-object -typename System.Management.Automation.PSCredential -argumentlist $365Admin, $365Pass
    Connect-MsolService -Credential $365Cred
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid -Credential $365Cred -Authentication Basic -AllowRedirection
    Import-PSSession $Session -AllowClobber 

}
function Start-UserLeft {
    & './UserLeft.ps1'
}
function Start-NewUser {
    & './NewUser.ps1'
}
function Start-FullAccess {
    Ellisons-Connect365
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
    Ellisons-Connect365
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
    Ellisons-Connect365
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
    Ellisons-Connect365
    Clear-Host
    $Requestee = $null
    $Requestee = Read-Host "Whos mailbox do you want to check permissions on?"
    Get-Mailbox $Requestee | Format-Table Name, grantsendonbehalfto -wrap
}
function Start-SyncAD {
    Start-DomaincredCheck
    Start-Process powershell.exe '.\SyncAD.ps1' -Credential $DomainCred
}
function Start-DisableOutOfOffice {
    Ellisons-Connect365
    Clear-Host
    $Requestee = $null
    $Requestee = Read-Host "Whos mailbox do you want to remove the Out Of Office for?"
    Set-MailboxAutoReplyConfiguration -Identity $Requestee -AutoReplyState Disabled
}
function Start-UpdateCreds {
    $credential = Get-Credential
    $credential.Password | ConvertFrom-SecureString | Out-File O365Account.txt
}
function Start-UpdateDomainCreds {
    $credential = Get-Credential
    $credential.Password | ConvertFrom-SecureString | Out-File DomainAdminAccount.txt
}

function Start-DomaincredCheck {
    if (Test-path -Path '.\DomainAdminName.txt') {$DomainAdminName = Get-Content '.\DomainAdminName.txt'}
    if (!$DomainAdminName) { $DomainAdminName = Read-Host "Enter your SRV Account Username (ELLNET\***-SRV)" | Out-File '.\DomainAdminName.txt'}
    if (!$DomainPass) { $DomainPass = Get-Content ".\DomainAdminAccount.txt" | ConvertTo-SecureString }
    $DomainCred = new-object -typename System.Management.Automation.PSCredential -argumentlist $DomainAdminName, $DomainPassS
}