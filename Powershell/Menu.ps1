##VERSION 3.0  - AP - 20/02/2020
Set-Location -Path $PSScriptRoot
#requires -Module ActiveDirectory
Import-Module ".\EllisonsModule.psm1" -Force

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