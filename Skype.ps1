Import-Module "C:\Program Files\Common Files\Skype for Business Online\Modules\SkypeOnlineConnector\SkypeOnlineConnector.psd1"
$AdminName = "Andrew.Powell@ellisonssolicitors.com"
$Pass = Get-Content "O365Account.txt" | ConvertTo-SecureString
$Cred = new-object -typename System.Management.Automation.PSCredential -argumentlist $AdminName, $Pass
Import-Module MSOnline
$session = New-CsOnlineSession -Credential $cred
Import-PSSession $session

#Set-CsClientPolicy -DisableEmoticons $False