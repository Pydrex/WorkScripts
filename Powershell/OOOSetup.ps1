Set-Location -Path $PSScriptRoot
$CSVFileName = Read-host "Enter the full Path to the CSV (EmailAddress,ContactAddress)"
$logfile = "results.log"

Function Connect365 {
     Set-ExecutionPolicy Unrestricted -Force 
     Import-Module MSOnline
     Get-PSSession | Remove-PSSession

     #$credential = Get-Credential
     #$credential.Password | ConvertFrom-SecureString | Out-File C:\PowerShell\O365Account.txt
     $AdminName = "Andrew.Powell@ellisonssolicitors.com"
     Import-Module MSOnline
     $Pass = Get-Content "C:\PowerShell\O365Account.txt" | ConvertTo-SecureString
     $Cred = new-object -typename System.Management.Automation.PSCredential -argumentlist $AdminName, $Pass
     Connect-MsolService -Credential $Cred
     $Session = Connect-ExchangeOnline -UserPrincipalName $Global:365AdminUsername -ShowProgress $false
     #$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid -Credential $cred -Authentication Basic -AllowRedirection
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
            #Create OOO message
            $contact = $usr.ContactAddress
            $email = $usr.EmailAddress
            $message = "Thank you for your email. In light the COVID19 pandemic we are following government guidelines on social distancing and therefore unable to attend our offices. If your query is urgent, please contact: $contact"
            Set-MailboxAutoReplyConfiguration $usr.EmailAddress -AutoReplyState enabled -ExternalAudience all -InternalMessage $message -ExternalMessage $message -ErrorAction STOP
            Set-Mailbox -Identity $email -Type Regular
            Set-MsolUser -UserPrincipalName $email -UsageLocation GB
            Set-MsolUserLicense -UserPrincipalName $email -AddLicenses reseller-account:SPE_E3
            $ServicePlans = "KAIZALA_O365_P3", "TEAMS1", "MICROSOFT_SEARCH", "MYANALYTICS_P2", "POWERAPPS_O365_P2", "FLOW_O365_P2", "YAMMER_ENTERPRISE", "SWAY", "Deskless", "WHITEBOARD_PLAN2", "BPOS_S_TODO_2", "FORMS_PLAN_E3", "STREAM_O365_E3"
            $AccountSkuId = "reseller-account:SPE_E3"
            $LO = New-MsolLicenseOptions -AccountSkuId $AccountSkuId -DisabledPlans $ServicePlans
            Set-MsolUserLicense -UserPrincipalName $email -LicenseOptions $LO -Verbose
            "$($usr.EmailAddress) was set successfully." | Out-File $logfile -Append
        }
        catch {
            
            $message = "A problem occured trying to create the $($usr.EmailAddress) out of office"
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