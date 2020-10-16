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
     Clear-Host
     [string]$Title = 'Main Menu' 
      Write-Host "================ $Title ================" 
      Write-Host "1:  Press '1' to Access the 365 Menu" 
      Write-Host "2:  Press '2' for the new user procedure." 
      Write-Host "3:  Press '3' for the user left procedure." 
      Write-Host "4:  Press '4' to sync all AD Controllers."
      Write-Host "5:  Press '5' to Unlock AD Accounts and sync"
      Write-Host "6:  Press '6' to add a user to RDS"
      Write-Host "7:  Press '7' to access PaperCUT Menu"
      Write-Host "8:  Press '8' to update the phone list"
      Write-Host "9:  Press '9' to update ADConnect sync"
      Write-Host "10: Press '10' to reset someones password expiry timer"
      Write-Host "Q:  Press 'Q' to quit." 
     $input = Read-Host "Please make a selection" 
     switch ($input) { 
          '1' { 
               Clear-Host
               'Starting the 365 Options menu'
               Start-365Menu
          }'2' { 
               Clear-Host 
               'You chose the New user procedure'
               Start-NewUser
          }'3' { 
               Clear-Host 
               'You chose Employee Left procedure'
               Start-UserLeft
          } '4' { 
               Clear-Host 
               'You Selected the sync AD procedure'
               Start-SyncAD
          } '5' { 
               Clear-Host 
               'Unlocking AD Accounts...'
               Start-UnlockedADAccounts 
          } '6' { 
               Clear-Host
               'Launching RDS Enable'
               Start-EnableRDS
          } '7' { 
               Clear-Host
               Start-PaperCutIDCheck
          } '8' { 
               Clear-Host
               'Updating phone list'
               Start-UpdatePhoneList
          } '9' { 
               Clear-Host
               Invoke-Command -ComputerName ez-az-dc01 -ScriptBlock { Start-ADSyncSyncCycle -PolicyType Delta }
          } '10' { 
               Clear-Host
               Start-PWDReset
          } 'U' { 
               Clear-Host
               'Please update your creds'
               Start-UpdateCred
           
          } 'D' { 
               Clear-Host
               'Please update your domain admin creds'
               Start-UpdateDomainCreds 
           
          } 'egg' { 
               Clear-Host
               Start-Egg
           
          } 'q' { 
               'Thank you, come again'
               Get-PSSession | Remove-PSSession
               return 
          } 
     } 
     pause 
} 
until ($input -eq 'q')
