$FormatEnumerationLimit=-1
Set-ExecutionPolicy Unrestricted -Force 
Import-Module MSOnline

#$credential = Get-Credential
#$credential.Password | ConvertFrom-SecureString | Out-File C:\PowerShell\O365Account.txt

$AdminName = "Andrew.Powell@ellisonssolicitors.com"
$Pass = Get-Content "O365Account.txt" | ConvertTo-SecureString
$Cred = new-object -typename System.Management.Automation.PSCredential -argumentlist $AdminName, $Pass
Import-Module MSOnline
Connect-SPOService -Url https://ellisonssolicitors-admin.sharepoint.com -Credential $cred

#$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid -Credential $cred -Authentication Basic -AllowRedirection#
#Import-PSSession $Session

## Request-SPOPersonalSite