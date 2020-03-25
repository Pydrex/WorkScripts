#Module

Write-Host "Loading Powershell Ellisons Module" -BackgroundColor Black -ForegroundColor Green
Write-Host "Version 1.5.1" -BackgroundColor Black -ForegroundColor Green
Write-Host "Created and Maintaned by Andrew Powell" -BackgroundColor Black -ForegroundColor Green
Write-Host "Updated 25/03/2020 - 10:43" -BackgroundColor Black -ForegroundColor Green

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

#requires -module ActiveDirectory
#requires -module MSOnline

Set-Location -Path $PSScriptRoot
$Global:currentUser = $env:UserName

function Start-UnlockedADAccounts {
    Import-Module ActiveDirectory
    $UnlockedUsers = (Search-ADAccount -LockedOut | Unlock-ADAccount)
    if($UnlockedUsers) {Write-Host '$UnlockedUsers'; Start-SyncAD} else {Write-Host "Found No Locked Out Accounts, Returning to menu" -ForegroundColor Green}

}

Function Enter-Office365 {

    if (!(Get-PSSession | Where { $_.ConfigurationName -eq "Microsoft.Exchange" })) { 
        Get-PSSession | Remove-PSSession
        Import-Module ActiveDirectory
        Import-Module MSOnline
        Set-CredsUp
        Connect-MsolService -Credential $Global:365Cred
        $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid -Credential $Global:365Cred -Authentication Basic -AllowRedirection
        Import-PSSession $Session -AllowClobber 
    }
    
}
function Set-CredsUp {

    $Global:365AdminPassword = Get-Content ".\creds\$Global:currentUser-O365Password.txt"
    $Global:365AdminUsername = Get-Content ".\creds\$Global:currentUser-O365AdminName.txt"
    $Global:SRVAdminPassword = Get-Content ".\creds\$Global:currentUser-SRVPassword.txt"
    $Global:SRVAdminUsername = Get-Content ".\creds\$Global:currentUser-SRVAdminName.txt"
    $Global:LocalAdminPassword = Get-Content ".\creds\$Global:currentUser-AdministratorPassword.txt"
    $Global:LocalAdminUsername = Get-Content ".\creds\$Global:currentUser-AdministratorName.txt"

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
                      $server = Get-ADDomain | Select-Object -ExpandProperty PDCEmulator
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

function Start-Egg {
    & './Snake.ps1'
}

function Start-UpdatePhoneList{
    & './UpdatePhoneLists.ps1'
    
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

function Start-EnableOWA {
    Enter-Office365
    #Clear-Host
    $email = $null
    $email = Read-Host "Whos email address do you want to enable OWA and MFA on?"
    $st = New-Object -TypeName Microsoft.Online.Administration.StrongAuthenticationRequirement
    $st.RelyingParty = "*"
    $st.State = "Enforced"
    $sta = @($st)
    Set-MsolUser -UserPrincipalName $email -StrongAuthenticationRequirements $sta
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
        Set-CASMailbox -Identity $email -ActiveSyncEnabled $false -OWAforDevicesEnabled $false -OWAEnabled $false
        
        $owaenabled = Get-CASMailbox -Identity $email | Select-Object Identity, `
            @{Name = 'OWAisEnabled'; Expression={if ($_.OWAEnabled) {Write-Output $true} else {Write-Output $false}}}
            Write-Output $owaenabled | Sort-Object OWAisEnabled
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