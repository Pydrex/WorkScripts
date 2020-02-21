##VERSION 5.0  - AP - 21/02/2020
Set-Location -Path $PSScriptRoot
Import-Module ".\EllisonsModule.psm1" -Force


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
 Clear-Host

do { 
     Show-Menu 
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
          }
          '3' { 
               Clear-Host 
               'You chose the Send On Behalf procedure'
               Start-SendOnBehalf
          }
          '4' { 
               Clear-Host 
               'You chose the view access procedure'
               Start-AccessBehalf
          }
          '5' { 
               Clear-Host 
               'You chose the New user procedure'
               Start-NewUser
          }
          '6' { 
               Clear-Host 
               'You chose Employee Left procedure'
               Start-UserLeft
          }
          '7' { 
               Clear-Host 
               'You chose Office 365 connection'
               Enter-Office365
          } '8' { 
               Clear-Host 
               'You Selected the Out Of Office Procedure Office 365 connection'
               Start-DisableOutOfOffice
          } '9' { 
               Clear-Host 
               'You Selected the sync AD procedure'
               Start-SyncAD
           
          } 'U' { 
               Clear-Host
               'Please update your creds'
               Start-UpdateCred
           
          } 'D' { 
               Clear-Host
               'Please update your domain admin creds'
               Start-UpdateDomainCreds 
           
          } 'q' { 
               'Thank you, come again'
               return 
          } 
     } 
     pause 
} 
until ($input -eq 'q')