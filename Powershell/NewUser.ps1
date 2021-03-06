#requires -Module ActiveDirectory
#Import-Module ActiveDirectory -EA Stop
#VERSION 8.0.0 - AP - 19/02/2020
#Added Papercut duplicate checker - V7
#Multiple fixes for AD Sync
<#
.Synopsis
    This will create a user with a mailbox in Office365 in Hybrid Exchange.
    Consult Andrew in IT for any support or issues with this query.
 
#>
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

    #EXCHANGE Local Administrator account Check
    if (Test-path -Path ".\creds\$Global:currentUser-EXCHAdminName.txt") { $Global:EXCHAdminUsername = Get-Content ".\creds\$Global:currentUser-EXCHAdminName.txt" }
    if (!$Global:EXCHAdminUsername) { Write-Host "MISSING SAVED Administrator NAME, QUEUING JOB" -ForegroundColor Red }
    if (Test-path -Path ".\creds\$Global:currentUser-EXCHAdminPassword.txt" ) { $Global:EXCHAdminPassword = Get-Content ".\creds\$Global:currentUser-EXCHAdminPassword.txt" | ConvertTo-SecureString }
    if (!$Global:EXCHAdminPassword) { Write-Host "MISSING SAVED Administrator PASSWORD, QUEUING JOB" -ForegroundColor Red }
    #Lets run the script to update the passwords
    if (!$Global:EXCHAdminUsername -OR !$Global:EXCHAdminPassword) { Start-EXCHAdministratorUpdate }
    $AdminName = "EXCHAdmin"
    $Global:EXCHAdminPassword = Get-Content ".\creds\$Global:currentUser-EXCHAdminPassword.txt" | ConvertTo-SecureString
    $Global:EXCHAdminUpdatecredential = new-object -typename System.Management.Automation.PSCredential -argumentlist $AdminName, $Global:EXCHAdminPassword

}
Start-CheckAllCreds
#Clear-Host

$Server = "EZ-AZ-DC01.Ellisonslegal.com"
#$OnPrem = New-PSSession -Authentication Kerberos -ConfigurationName Microsoft.Exchange -ConnectionUri 'http://ez-az-ex01.ellisonslegal.com/Powershell' -Credential $Global:AdminCred

Write-Host "Done..."
Clear-Host

if (Get-Module -ListAvailable -Name ActiveDirectory) {
    Write-Host "Active Directory Imported"
}
else { 
    Import-Module ActiveDirectory
}

if (Get-Module -ListAvailable -Name MSOnline) {
    Write-Host "MSOnline Imported"
}
else { 
    Import-Module MSOnline
}
#
#if (!(Get-PSSession | Where { $_.ConfigurationName -eq "Microsoft.Exchange" })) { 
#    $OnPrem = New-PSSession -Authentication Kerberos -ConfigurationName Microsoft.Exchange -ConnectionUri 'http://ez-az-ex01.ellisonslegal.com/Powershell' -Credential $Global:AdminCred
#}
#Import-PSSession $OnPrem | Out-Null

# Connection URI to connect to OnPrem Exchange PowerShell Web Session
$ConnectionURI = "http://ez-az-ex01.ellisonslegal.com/Powershell"

# Get credentials for Exchange Online and Exchange OnPrem Admin account
$OnPremcredential = $Global:AdminCred
$O365credential = $Global:365Cred

# Connect to Exchange Online V2
Connect-ExchangeOnline -UserPrincipalName $Global:365AdminUsername `
                       -Prefix 'EO' `
                       -Credential $O365credential 

# Connect to Exchange OnPremises
$Session = New-PSSession -ConfigurationName Microsoft.Exchange `
                         -ConnectionUri $ConnectionURI `
                         -Authentication Kerberos `
                         -Credential $OnPremcredential
Import-PSSession $Session -Prefix OP | Out-Null


do {
function Show-Menu { 
    param ( 
        [string]$Title = 'What type of account is this?' 
    ) 
    Clear-Host     
    Write-Host "Press 'U' for A normal user" 
    Write-Host "Press 'R' for Room Mailbox." 
    Write-Host "Press 'S' for Shared Mailbox."
} 
do { 
    Show-Menu 
    $input = Read-Host "Please make a selection" 
    switch ($input) { 
        'U' { 
            'You chose a normal user'
            $Usertype = 'normal'

        } 'R' { 
            'You chose a Room Mailbox'
            $Usertype = 'room'

        }
        'S' { 
            'You chose a Shared Mailbox'
            $Usertype = 'shared'
        }
    } 
    pause 
} 
until ($input)

Clear-Host

Write-Host "Before we create the account"
$CopyUser = Read-Host "Would you like to copy from another user? (y/n)"
Write-Host
  
Do {
    if ($CopyUser -ieq 'y') {
        $CUser = Read-Host "Enter in the USERNAME that you would like to copy FROM"
        Write-Host
  
          
        Write-Host "Checking if $CUser is a valid user..." -ForegroundColor:Green
        If ($(Get-ADUser -Filter { SamAccountName -eq $CUser })) {
            Write-Host "Copying from user account" (Get-ADUser $CUser | Select-Object -ExpandProperty DistinguishedName)
            Write-Host
  
            $Proceed = Read-Host "Continue? (y/n)"
            Write-Host
  
  
            if ($Proceed -ieq 'y') {
                $CUser = Get-ADUser $CUser -Properties *
                $Exit = $true
            }
  
        }
        else {
            Write-Host "$CUser was not a valid user" -ForegroundColor:Red
            Start-Sleep 4
            $Exit = $false
            Clear-Host
        }
  
    }
    else {
        $Exit = $true
    }
  
} until ($Exit -eq $true)
  

  
Clear-Host
Write-Host "Gathering information for new account creation."
Write-Host
$firstname = Read-Host "Enter in the First Name"
Write-Host
$lastname = Read-Host "Enter in the Last Name"
Write-Host
if ($Usertype -ieq 'normal') { 
    $fullname = "$firstname $lastname"
    $logonname = "$firstname.$lastname"
}

if ($Usertype -ieq 'room') { 
    Write-Host "Enter the mailbox name such as ROOM-Chelmsford-1"
    $fullname = Read-Host "Enter the full Name (No Spaces)"
    $logonname = "$firstname.$lastname"
    $email = Read-Host "Enter the new full email including @ellisonssolicitors.com"
}

if ($Usertype -ieq 'shared') { 
    Write-Host "Enter the Shared Mailbox Name"
    $fullname = Read-Host "Enter the Mailbox Name (No Spaces)"
    $logonname = "$fullname"
    $email = Read-Host "Enter the new full email including @ellisonssolicitors.com"
}

$password = Read-Host "Enter in the password" -AsSecureString
  
$domain = "ellisonssolicitors.com"
  
if ($CUser) {
    #Getting OU from the copied User.
    $Object = $CUser | Select-Object -ExpandProperty DistinguishedName
    $pos = $Object.IndexOf(",OU")
    $OU = $Object.Substring($pos + 1)
    #$OU = "CN=Users,DC=Ellisonslegal,DC=com"
  
    #Getting Description from the copied User.
    $Description = $CUser.Description

  
    #Getting Office from the copied User.
    $Office = $CUser.Office
  
    #Getting Street Address from the copied User.
    $StreetAddress = $CUser.StreetAddress
  
    #Getting City from copied user.
    $City = $CUser.City
  
    #Getting State from copied user.
    $State = $CUser.State
  
    #Getting PostalCode from copied user.
    $PostalCode = $CUser.PostalCode
  
    #Getting Country from copied user.
    $Country = $CUser.Country

    #Getting POBox from copied user.
    $POBox = $CUser.POBox

    #Getting HomePhone from copied user.
    $HomePhone = $CUser.HomePhone

    #Getting OfficePhone from copied user.
    $OfficePhone = $CUser.OfficePhone

    #Getting Fax from copied user.
    $Fax = $CUser.Fax

    #Getting Fax from copied user.
    $Homepage = $CUser.Homepage

    #Getting Title from copied user.
    $Title = $CUser.Title
  
    #Getting Department from copied user.
    $Department = $CUser.Department
  
    #Getting Company from copied user.
    $Company = $CUser.Company
  
    #Getting Manager from copied user.
    $Manager = $CUser.Manager
  
    #Getting Membership groups from copied user.
    #$MemberOf = (Get-ADUser -Identity $CUser -Properties MemberOf).MemberOf -replace '^CN=([^,]+),OU=.+$','$1'
    #$MemberOf = Get-ADPrincipalGroupMembership $CUser -Server $server | Where-Object { $_.Name -ine "Domain Users" }
    $MemberOf = Get-ADUser -Identity $CUser -Properties memberof | Select-Object -ExpandProperty memberof 

    if (!$MemberOf) {
        Write-Host "FAILED TO GET GROUPS FROM USER - RESTART POWERSHELL AS ADMIN AND MAKE SURE ActiveDirectory is IMPORTED" -ForegroundColor:RED
        Write-Host "You can also carry on and manually add groups later" -ForegroundColor:RED
        pause
    }
  
}
else {
    #Getting the default Users OU for the domain.
    $OU = (Get-ADObject -Filter 'ObjectClass -eq "Domain"' -Properties wellKnownObjects).wellKnownObjects | Select-String -Pattern 'CN=Users'
    $OU = $OU.ToString().Split(':')[3]
  
}



Clear-Host
Write-Host "======================================="
Write-Host
Write-Host "Firstname:      $firstname"
Write-Host "Lastname:       $lastname"
Write-Host "Display name:   $fullname"
Write-Host "Logon name:     $logonname"
Write-Host "Email Address:  $firstname.$lastname@$domain"
Write-Host "OU:             $OU"
  
  
DO {
    If ($(Get-ADUser -Filter { SamAccountName -eq $logonname })) {
        Write-Host "WARNING: Logon name" $logonname.toUpper() "already exists!!" -ForegroundColor:Green
        $i++
        $logonname = $firstname + $lastname.substring(0, $i)
        Write-Host
        Write-Host
        Write-Host "Changing Logon name to" $logonname.toUpper() -ForegroundColor:Green
        Write-Host
        $taken = $true
        Start-Sleep 4
    }
    else {
        $taken = $false
    }
} Until ($taken -eq $false)
$logonname = $logonname.toLower()

if ($Usertype -ieq 'normal') {
    $email = "$firstname.$lastname@$domain"
    $EmployeeID = Read-Host "Enter the Papercut Code (Usually Todays date)"
    $finding = Get-ADUser -Filter { EmployeeId -eq $EmployeeID } -Properties EmployeeId

    if ($finding) {
        $idLookup = $finding | ForEach-Object { $idLookup = @{ } } { if ($_.EmployeeId) { $idLookup[$_.EmployeeId] += 1 } } { $idLookup }
        $filteredUsers = $finding | Where-Object { if ($_.EmployeeId) { $idLookup[$_.EmployeeId] -gt 1 } }
        $filteredUsers | Select-Object -Property SamAccountName, EmployeeId
        Write-Host "!DUPLICATE PAPERCUT ID FOUND! - PLEASE CHANGE" -ForegroundColor:Red
    }
}

Clear-Host
Write-Host "======================================="
Write-Host
Write-Host "Firstname:      $firstname"
Write-Host "Lastname:       $lastname"
Write-Host "Display name:   $fullname"
Write-Host "Logon name:     $logonname"
Write-Host "Email Address:  $email"
Write-Host "OU:             $OU"
if ($Usertype -ieq 'normal') { Write-Host "Papercut ID:      $EmployeeID" }
Write-Host
Write-Host

Write-Host "Continuing will create the AD account and O365 Email." -ForegroundColor:Green
Write-Host
$Proceed = $null
$Proceed = Read-Host "Continue? (y/n)"

if ($Proceed -ieq 'y') {
    
    if ($Usertype -ieq 'normal') { 
        Write-Host "Creating the O365 mailbox and AD Account."
        New-OPRemoteMailbox -Name $fullname -FirstName $firstname -LastName $lastname -DisplayName $fullname -SamAccountName $logonname -UserPrincipalName $logonname@$domain -PrimarySmtpAddress $email -Password $password -OnPremisesOrganizationalUnit $OU -DomainController $Server
        Write-Host "Done..."
        Write-Host
        Write-Host
        Start-Sleep 5
   
  
        Write-Host "Adding Properties to the new user account."
        Get-ADUser $logonname -Server $Server | Set-ADUser -Server $Server -Description $Description -Office $Office -StreetAddress $StreetAddress -City $City -State $State -PostalCode $PostalCode -Country $Country -Title $Title -Department $Department -Company $Company -Manager $Manager -EmployeeID $EmployeeID -Fax $Fax -Homepage $Homepage -HomePhone $HomePhone -POBox $POBox -OfficePhone $OfficePhone
        Write-Host "Done..."
        Write-Host "Adding Default Groups (for All Users + FollowMe)"
        Add-ADGroupMember -Identity "All Users" -Members $logonname
        Add-ADGroupMember -Identity "FollowMeUsers" -Members $logonname

        #Add SIP for skype #ADDED 25/09/2019
        foreach ($user in (Get-ADUser -Identity $logonname -Properties mail, ProxyAddresses, UserPrincipalName)) {
        $user.ProxyAddresses += ("SIP:" + $email)
        Set-ADUser -instance $user
    }
    }

    if ($Usertype -ieq 'room') { 
        Write-Host "Creating the O365 room mailbox and AD Account."
        New-OPRemoteMailbox -Name $fullname -FirstName $firstname -LastName $lastname -DisplayName $fullname -SamAccountName $logonname -UserPrincipalName $logonname@$domain -PrimarySmtpAddress $email -Password $password -OnPremisesOrganizationalUnit $OU -DomainController $Server -Room
        Write-Host "Done..."
        Write-Host
        Write-Host
        Start-Sleep 5
       
      
        Write-Host "Adding Properties to the new user account."
        Get-ADUser $logonname -Server $Server | Set-ADUser -Server $Server -Description $Description -Office $Office -StreetAddress $StreetAddress -City $City -State $State -PostalCode $PostalCode -Country $Country -Title $Title -Department $Department -Company $Company -Manager $Manager -Fax $Fax -Homepage $Homepage -HomePhone $HomePhone -POBox $POBox -OfficePhone $OfficePhone
        Write-Host "Done..."
        Write-Host
        Write-Host
    }

    if ($Usertype -ieq 'shared') { 
        Write-Host "Creating the O365 Shared mailbox and AD Account."
        New-OPRemoteMailbox -Name $fullname -FirstName $firstname -LastName $lastname -DisplayName $fullname -SamAccountName $logonname -UserPrincipalName $logonname@$domain -PrimarySmtpAddress $email -Password $password -OnPremisesOrganizationalUnit $OU -DomainController $Server -Shared
        Write-Host "Done..."
        Write-Host
        Write-Host
        Start-Sleep 5
           
          
        Write-Host "Adding Properties to the new user account."
        Get-ADUser $logonname -Server $Server | Set-ADUser -Server $Server -Description $Description -Office $Office -StreetAddress $StreetAddress -City $City -State $State -PostalCode $PostalCode -Country $Country -Title $Title -Department $Department -Company $Company -Manager $Manager -Fax $Fax -Homepage $Homepage -HomePhone $HomePhone -POBox $POBox -OfficePhone $OfficePhone
        Write-Host "Done..."
        Write-Host
        Write-Host
    }
    
    
    if ($MemberOf) {
        $CopyName = $CUser.Name 
        $ProceedMemberOf = Read-host "Do you want to copy AD User Groups from $CopyName to $fullname (y/n)?"
        if ($ProceedMemberOf -ieq 'y') {
        
        Write-Host "Adding Membership Groups to the new user account."
        $MemberOf | Add-ADGroupMember -Members $logonname
        #Get-ADUser $logonname -Server $Server | Add-ADPrincipalGroupMembership -Server $Server -MemberOf $MemberOf
        Write-Host "Done..."
        Write-Host
        Write-Host
    }
    }

}

Start-SyncAD
Connect-MsolService -Credential $Global:365Cred
Write-Host "Creating $logonname in the cloud...."
Invoke-Command -ComputerName ez-az-dc01 -ScriptBlock { Start-ADSyncSyncCycle -PolicyType Delta }
Write-Host "Sleeping until sync completed 2 minutes remaining"
Start-Sleep -s 120


if ($Usertype -ieq 'normal') { 

    #LICENSE USER ACCOUNT

    Set-MsolUser -UserPrincipalName $email -UsageLocation GB
    Set-MsolUserLicense -UserPrincipalName $email -AddLicenses "reseller-account:SPE_E3"
}

Write-Host "Applying licenses 2 minutes remaining"
Start-Sleep -s 120

Connect-SPOService -Url https://ellisonssolicitors-admin.sharepoint.com -Credential $Global:365Cred
Request-SPOPersonalSite -UserEmails $email -NoWait
#Disable OWA+SendAs and Behalf defaults

    Set-EOCASMailbox -Identity $email -ActiveSyncEnabled $false -OWAforDevicesEnabled $false -OWAEnabled $false
    Set-EOMailbox -Identity $email -MessageCopyForSentAsEnabled $True
    Set-EOMailbox -Identity $email -MessageCopyForSendOnBehalfEnabled $True
    #Set Perms for Calendars
    $cal = $email + ":\Calendar"
    Set-EOMailboxFolderPermission $cal -User Default -AccessRights Reviewer

Write-Host "Login into the cloud to see if the user exists and is licensed properly"
Write-Host "Also check they have the correct job title and member groups"

$Completed = Read-host "Type y to exit or n if you wish to run this script again (y/n)"
if ($Completed -ieq 'y') {
    Get-PSSession | Remove-PSSession
    $done = $true
}

} while ($done -ne $true)

$done = $Null;
Get-PSSession | Remove-PSSession