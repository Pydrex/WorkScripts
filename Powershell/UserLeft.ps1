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
Get-PSSession | Remove-PSSession
Set-Location -Path $PSScriptRoot
Import-Module MSOnline
$AdminName = Read-Host "Enter your Office 365 Admin email (First.Last@ellisonssolcitiors.com) etc..."
$Pass = Get-Content ".\O365Account.txt" | ConvertTo-SecureString
$Cred = new-object -typename System.Management.Automation.PSCredential -argumentlist $AdminName, $Pass
Connect-MsolService -Credential $Cred
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid -Credential $cred -Authentication Basic -AllowRedirection
Import-PSSession $Session -AllowClobber
cls
$Password = ([char[]]([char]33..[char]95) + ([char[]]([char]97..[char]126)) + 0..9 | sort {Get-Random})[0..8] -join ''
  
    Do {
        if ($Password) {
         $EmailAddress = read-host 'Enter users email address you want to disable'
         Write-Host
          $sam = Get-ADUser -Filter { emailaddress -Like $EmailAddress} -Properties SamAccountName
          
            Write-Host "Checking if $sam is a valid user..."
            If ($(Get-ADUser $sam)) {
                    Write-Host "USER FOUND:" (Get-ADUser $sam | select -ExpandProperty DistinguishedName) -ForegroundColor:Green
                    Write-Host
  
                       $Proceed = Read-Host "Is this correct? (y/n)"
                       Write-Host
 
                    if ($Proceed -ieq 'y') {
                    $Exit = $true
                }
  
            } else {
            Write-Host "$sam was not a valid user - CHECK AGAIN" -ForegroundColor:Red
            Sleep 4
            $Exit = $false
            cls
            }
  
        } else {
        $Exit = $true
        }
  
    } until ($Exit -eq $true)




$Proceed = Read-Host "Do you want to add permission to acess this inbox (y/n)"
            Write-Host
  
  
                if ($Proceed -ieq 'y') {
                    $supervisor = read-Host "User who is going to be having access to shared mailbox"
                    Set-Mailbox $EmailAddress -Type shared
                    Add-MailboxPermission -Identity $EmailAddress -User $supervisor -AccessRights FullAccess
                } else {
               Set-Mailbox $EmailAddress -Type shared
            }

Set-MsolUser -UserPrincipalName $EmailAddress -StrongPasswordRequired $False
Set-MsolUserPassword -UserPrincipalName $EmailAddress -NewPassword $Password -ForceChangePassword $false

Write-host "Completed.  Password changed to $Password for account $EmailAddress"

##This section removes all licenses (use get-msolaccountsku to find out yours), and adds Exchange Enterprise license
##which is required for litigation hold.  You may not need that for your environment, so adjust accordingly.

(get-MsolUser -UserPrincipalName $EmailAddress).licenses.AccountSkuId |
foreach{
    Set-MsolUserLicense -UserPrincipalName $EmailAddress -RemoveLicenses $_
}

Get-ADUser $sam | Move-ADObject -TargetPath 'OU=Disabled user accounts,DC=Ellisonslegal,DC=com'
Disable-ADAccount -identity $sam

Set-ADUser -Identity $sam -Replace @{msExchHideFromAddressLists=$True}

Get-ADUser $sam -Properties MemberOf | Select -Expand MemberOf | %{Remove-ADGroupMember $_ -member $sam}
$datestamp = Get-Date -Format g
$initials = Read-host "Enter your Initials for the lock out stamp"
Get-aduser $sam -Properties Description | ForEach-Object { Set-ADUser $_ -Description "$($_.Description) Disabled by $initials - $datestamp" }
$contactemail = read-host 'Enter the full email of the person who should be in the Out Of Office reply example (Contact HoD for email)'
$internalMsg = "Please note I am no longer working with Ellisons Solicitors. If you have questions please contact $contactemail and they will get back to you as soon as possible."
$externalMsg = "Please note I am no longer working with Ellisons Solicitors. If you have questions please contact $contactemail and they will get back to you as soon as possible."
Set-MailboxAutoReplyConfiguration -Identity $EmailAddress -AutoReplyState Enabled -InternalMessage $internalMsg -ExternalMessage $externalMsg


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

Get-PSSession | Remove-PSSession