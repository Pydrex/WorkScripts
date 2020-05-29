#VERSION #0003 - UPDATED 28.01.2020
##This script is to disable users, change their passwords, move them to a different OU, force sync your domain controllers, remove Office 365 licenses,
##Also change yourdomain at various points.  The OU "Disabled Accounts" portion moves the account to that OU, and keeps things tidy.


##This section requires the profile.ps1 file found here:  https://github.com/Scine/Powershell/blob/master/profile.ps1
##Put that file under your Documents\Windows Powershell\ folder.

#If you don't have 2FA authentication enabled uncomment this section

#$UserCredential = Get-Credential
#Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
#Import-PSSession $Session

#With 2FA authentication enabled already.  If you don't have this enabled, use the above section on line 6 and comment out the next 3 lines below by putting a # at the beginning of each line.
Set-Location -Path $PSScriptRoot
function Start-CheckAllCreds {
    #OnPrem Local Administrator account Check
    if (Test-path -Path ".\creds\$Global:currentUser-AdministratorName.txt") { $Global:LocalAdminUsername = Get-Content ".\creds\$Global:currentUser-AdministratorName.txt" }
    if (!$Global:LocalAdminUsername) { Write-Host "MISSING SAVED Administrator NAME, QUEUING JOB" -ForegroundColor Red }
    if (Test-path -Path ".\creds\$Global:currentUser-AdministratorPassword.txt" ) { $Global:LocalAdminPassword = Get-Content ".\creds\$Global:currentUser-AdministratorPassword.txt" | ConvertTo-SecureString }
    if (!$Global:LocalAdminPassword) { Write-Host "MISSING SAVED Administrator PASSWORD, QUEUING JOB" -ForegroundColor Red }
    #Lets run the script to update the passwords
    if (!$Global:LocalAdminUsername -OR !$Global:LocalAdminPassword) { Start-AdministratorUpdate }
    $AdminName = "Administrator"
    $Global:LocalAdminPassword = Get-Content ".\creds\$Global:currentUser-AdministratorPassword.txt" | ConvertTo-SecureString
    $Global:AdminCred = new-object -typename System.Management.Automation.PSCredential -argumentlist $AdminName, $Global:LocalAdminPassword

    Write-Host "Passed Local Administrator Cred check" -ForegroundColor Green

    #365 Account Check
    if (Test-path -Path ".\creds\$Global:currentUser-O365AdminName.txt") { $Global:365AdminUsername = Get-Content ".\creds\$Global:currentUser-O365AdminName.txt" }
    if (!$Global:365AdminUsername) { Write-Host "MISSING SAVED O365 ADMIN NAME, QUEUING JOB" -ForegroundColor Red }
    if (Test-path -Path ".\creds\$Global:currentUser-O365Password.txt" ) { $Global:365AdminPassword = Get-Content ".\creds\$Global:currentUser-O365Password.txt" | ConvertTo-SecureString }
    if (!$Global:365AdminPassword) { Write-Host "MISSING SAVED O365 ADMIN PASSWORD, QUEUING JOB" -ForegroundColor Red } 
    #Lets run the script to update the passwords
    if (!$Global:365AdminUsername -OR !$Global:365AdminPassword) { Start-UpdateCreds }
    $Global:365Cred = new-object -typename System.Management.Automation.PSCredential -argumentlist $Global:365AdminUsername, $Global:365AdminPassword
    Write-Host "Passed 365 Account Cred check" -ForegroundColor Green

    #SRV Account Check
    if (Test-path -Path ".\creds\$Global:currentUser-SRVAdminName.txt") { $Global:SRVAdminUsername = Get-Content ".\creds\$Global:currentUser-SRVAdminName.txt" }
    if (!$Global:SRVAdminUsername) { Write-Host "MISSING SAVED SRV ADMIN NAME, QUEUING JOB" -ForegroundColor Red }
    if (Test-path -Path ".\creds\$Global:currentUser-SRVPassword.txt" ) { $Global:SRVAdminPassword = Get-Content ".\creds\$Global:currentUser-SRVPassword.txt" | ConvertTo-SecureString }
    if (!$Global:SRVAdminPassword) { Write-Host "MISSING SAVED SRV ADMIN PASSWORD, QUEUING JOB" -ForegroundColor Red } 
    #Lets run the script to update the passwords
    if (!$Global:SRVAdminUsername -OR !$Global:SRVAdminPassword) { Start-UpdateSRVCreds }
    $Global:SRVCred = new-object -typename System.Management.Automation.PSCredential -argumentlist $Global:SRVAdminUsername, $Global:SRVAdminPassword
    Write-Host "Passed SRV Account Cred check" -ForegroundColor Green
    Write-Host "Initialising" -BackgroundColor Gray
}
Start-CheckAllCreds
Enter-Office365
$Password = ([char[]]([char]33..[char]95) + ([char[]]([char]97..[char]126)) + 0..9 | Sort-Object { Get-Random })[0..8] -join ''
Do {
    if ($Password) {
        Clear-Host
        $EmailAddress = read-host 'Enter users email address you want to disable'
        Write-Host
        $sam = Get-ADUser -Filter { emailaddress -Like $EmailAddress } -Properties SamAccountName
        $logonname = (Get-ADUser $sam | Select-Object -ExpandProperty SamAccountName)
          
        Write-Host "Checking if $sam is a valid user..."
        If ($(Get-ADUser $sam)) {
            Write-Host "USER FOUND: " (Get-ADUser $sam | Select-Object -ExpandProperty DistinguishedName) -ForegroundColor:Green
            Write-Host "Username is " $logonname -ForegroundColor:Green
            Write-Host
  
            $Proceed = Read-Host "Is this correct? (y/n)"
            Write-Host
 
            if ($Proceed -ieq 'y') {
                $Exit = $true
            }
  
        }
        else {
            Write-Host "$sam was not a valid user - CHECK AGAIN" -ForegroundColor:Red
            Start-Sleep 4
            $Exit = $false
            Clear-Host
        }
  
    }
    else {
        $Exit = $true
    }
  
} until ($Exit -eq $true)


$Proceed = Read-Host "Do you want to add permission to acess this inbox (y/n)"
Write-Host
Import-Module ActiveDirectory
  
if ($Proceed -ieq 'y') {
    $supervisor = read-Host "User who is going to be having access to shared mailbox"
    Set-Mailbox $EmailAddress -Type shared
    Add-MailboxPermission -Identity $EmailAddress -User $supervisor -AccessRights FullAccess
}
else {
    Set-Mailbox $EmailAddress -Type shared
}

Set-MsolUser -UserPrincipalName $EmailAddress -StrongPasswordRequired $False
Set-MsolUserPassword -UserPrincipalName $EmailAddress -NewPassword $Password -ForceChangePassword $false

Write-host "Completed.  Password changed to $Password for account $EmailAddress"

##This section removes all licenses (use get-msolaccountsku to find out yours), and adds Exchange Enterprise license
##which is required for litigation hold.  You may not need that for your environment, so adjust accordingly.

(get-MsolUser -UserPrincipalName $EmailAddress).licenses.AccountSkuId |
ForEach-Object {
    Set-MsolUserLicense -UserPrincipalName $EmailAddress -RemoveLicenses $_
}
Get-ADUser $logonname | Move-ADObject -TargetPath 'OU=Disabled user accounts,DC=Ellisonslegal,DC=com'
Disable-ADAccount -identity $logonname

Set-ADUser -Identity $logonname -Replace @{msExchHideFromAddressLists = $True }

Get-ADUser $logonname -Properties MemberOf | Select-Object -Expand MemberOf | ForEach-Object { Remove-ADGroupMember $_ -member $logonname -Confirm:$False }
$datestamp = Get-Date -Format g
$initials = Read-host "Enter your Initials for the lock out stamp"
Get-aduser $logonname -Properties Description | ForEach-Object { Set-ADUser $_ -Description "$($_.Description) Disabled by $initials - $datestamp" }
$contactemail = read-host 'Enter the full email of the person who should be in the Out Of Office reply example (Contact HoD for email)'
$internalMsg = "Please note I am no longer working with Ellisons Solicitors. If you have questions please contact $contactemail and they will get back to you as soon as possible."
$externalMsg = "Please note I am no longer working with Ellisons Solicitors. If you have questions please contact $contactemail and they will get back to you as soon as possible."
Set-MailboxAutoReplyConfiguration -Identity $EmailAddress -AutoReplyState Enabled -InternalMessage $internalMsg -ExternalMessage $externalMsg

#Start-SyncAD

Get-PSSession | Remove-PSSession