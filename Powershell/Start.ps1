#Start.Ps1
Set-Location -Path $PSScriptRoot
$Global:currentUser = $env:UserName
Import-Module ".\EllisonsModule.psm1" -Force


Write-Host "Loading checks...."
Write-Host "Informational: Script has been run as user: $Global:currentUser" -ForegroundColor Green

$dir = ".\creds"
if(!(Test-Path -Path $dir )){
    New-Item -ItemType directory -Path $dir
}

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

    #365 Account Check 4
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
Clear-Host
do { 
    Write-Host "1: Press '1' for General Menu." 
    Write-Host "2: Press '2' for Updating Stored SRV Admin Creds."
    Write-Host "3: Press '3' for Updated Stored Office 365 Creds."
    Write-Host "4: Press '4' to add the local 'Administrator' Creds"
    Write-Host "5: Press '5' to add the local Exchange admin Creds" 
    $input = Read-Host "Please make a selection" 
    switch ($input) { 
         '1' { 
              Clear-Host
              Start-Process powershell.exe '.\Menu.ps1' -Credential $Script:SRVCred
         } '2' { 
                Start-UpdateSRVCreds
         } '3' { 
                Start-UpdateCreds
        }  '4' { 
                Start-AdministratorUpdate
        }  '5' { 
                Start-EXCHAdministratorUpdate
    }
    } 
    pause
} 
until (!$input)