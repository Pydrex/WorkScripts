#SyncAD Controllers
#Require Administrator rights#

Get-PSSession | Remove-PSSession
$DomainControllers = Get-ADDomainController -Filter *
ForEach ($DC in $DomainControllers.Name) {
    Write-Host "Processing for "$DC -ForegroundColor Green
    If ($Mode -eq "ExtraSuper") {
        REPADMIN /kcc $DC
        REPADMIN /syncall /A /e /q $DC
    }
    Else {
        REPADMIN /syncall $DC "DC=Ellisonslegal,DC=com" /d /e /q
    }
}

Invoke-Command -ComputerName ez-az-dc01 -ScriptBlock { Start-ADSyncSyncCycle -PolicyType Delta }