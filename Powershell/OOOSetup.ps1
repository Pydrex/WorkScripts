Set-Location -Path $PSScriptRoot
$CSVFileName = Read-host "Enter the full Path to the CSV (EmailAddress,ContactAddress"
$logfile = "results.log"

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
            $message = "Thank you for your email. In light the COVID19 pandemic we are following government guidelines on social distancing and therefore unable to attend our offices. If your query is urgent, please contact: $contact"
            Set-MailboxAutoReplyConfiguration $usr.EmailAddress -AutoReplyState enabled -ExternalAudience all -InternalMessage $message -ExternalMessage $message -ErrorAction STOP
            "$($usr.EmailAddress) was created successfully." | Out-File $logfile -Append
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