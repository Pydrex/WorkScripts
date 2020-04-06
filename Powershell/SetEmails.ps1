
###WILL NEED TO SET PEOPLES OWA TO PRIMARY SMTP IF RAN!
Set-Location -Path $PSScriptRoot
$CSVFileName = Read-host "Enter the full Path to the CSV (EmailAddress,ContactAddress)"
$logfile = "resultssetup.log"

Function Connect365 {
     Set-ExecutionPolicy Unrestricted -Force 
     Import-Module MSOnline
     Get-PSSession | Remove-PSSession
     Import-Module ActiveDirectory

     #$credential = Get-Credential
     #$credential.Password | ConvertFrom-SecureString | Out-File C:\PowerShell\O365Account.txt
     $AdminName = "Andrew.Powell@ellisonssolicitors.com"
     Import-Module MSOnline
     $Pass = Get-Content "C:\PowerShell\O365Account.txt" | ConvertTo-SecureString
     $Cred = new-object -typename System.Management.Automation.PSCredential -argumentlist $AdminName, $Pass
     Connect-MsolService -Credential $Cred
     $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid -Credential $cred -Authentication Basic -AllowRedirection
     Import-PSSession $Session -AllowClobber 

}

Get-Date | Out-File $logfile

#Check if OU exists

If (Test-Path $CSVFileName) {

    #Import the CSV file
    $csvfile = Import-CSV $CSVFileName
        
    #Loop through CSV file
    foreach ($usr in $csvfile) {

        try {
            $ellisons = $usr.Ellisons
            $SMTP1 = $usr.SMTP1
            $SMTP2 = $usr.SMTP2
            $SMTP3 = $usr.SMTP3

            $filtersam = Get-ADUser -Filter { emailaddress -eq $Ellisons } -Properties SamAccountName
            $server = Get-ADDomain | Select-Object -ExpandProperty PDCEmulator
            #$sam = $filtersam | Select-Object -Property SamAccountName
            $identitysam = (Get-ADUser -Identity $filtersam -Properties mail, ProxyAddresses, UserPrincipalName)
            
            if ($SMTP1) {$identitysam.ProxyAddresses += ($SMTP1); Set-ADUser -instance $identitysam; Write-Host "Set $SMTP1"}
            if ($SMTP2) {$identitysam.ProxyAddresses += ($SMTP2); Set-ADUser -instance $identitysam; Write-Host "Set $SMTP2"}
            if ($SMTP3) {$identitysam.ProxyAddresses += ($SMTP3); Set-ADUser -instance $identitysam; Write-Host "Set $SMTP3"}
            
            "$($usr.EmailAddress) was set successfully." | Out-File $logfile -Append
        }
        catch {
            
            $message = "A problem occured updating $($usr.Ellisons)"
            $message | Out-File $logfile -Append
            Write-Warning $message
            Write-Warning $_.Exception.Message
            $_.Exception.Message | Out-File $logfile -Append
        }

    }
}
else {

    $message = "The CSV file $CSVFileName was not found."
    $message | Out-File $logfile -Append
    throw $message

}