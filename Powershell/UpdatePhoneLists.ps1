#Author: Andrew Powell
#Version: 1.0.1
#Updated: 10/05/2019
#Purpose: Script to replace phone lists

Invoke-Command -ComputerName EZ-AZ-FS01 -ScriptBlock {Get-SmbOpenFile | Where-Object -Property sharerelativepath -match "TELEPHONE LISTS AND FLOOR PLAN" | Close-SmbOpenFile -Force ;

$AlphaListPath = 'F:\Global Share\TELEPHONE LISTS AND FLOOR PLAN\Alphabetical DD and Internal Nos - New.pdf'
$DDIListPath = 'F:\Global Share\TELEPHONE LISTS AND FLOOR PLAN\Direct Dials and Extensions - Departments - New.pdf'

$WantFile = $AlphaListPath ;
$FileExists = Test-Path $WantFile;
If ($FileExists -eq $True) {$copyAlphabetical = "File Updated" ; $copyAlphabeticalTask = Copy-Item -path $AlphaListPath -Destination 'F:\Global Share\TELEPHONE LISTS AND FLOOR PLAN\Alphabetical DD and Internal Nos.pdf' -Force -PassThru -ErrorAction SilentlyContinue;
Remove-Item -Path $AlphaListPath -Force}
Else {$copyAlphabetical = "No update found"};


$WantFile = $DDIListPath;
$FileExists = Test-Path $WantFile;
If ($FileExists -eq $True) {$copyDDI = "File Updated" ; $copyDDITask =  Copy-Item -path $DDIListPath -Destination 'F:\Global Share\TELEPHONE LISTS AND FLOOR PLAN\Direct Dials and Extensions - Departments.pdf' -Force -PassThru -ErrorAction SilentlyContinue;
Remove-Item -Path $DDIListPath -Force ;}
Else {$copyDDI = "No update found"};



function Get-TimeStamp {
    
    return "[{0:MM/dd/yy} {0:HH:mm:ss}]" -f (Get-Date)
    
};

Write-Output "$(Get-TimeStamp) Last updated | Status: Alphabetical: $copyAlphabetical | Departmental: $copyDDI" | Out-file 'F:\Global Share\TELEPHONE LISTS AND FLOOR PLAN\LastUpdateTime.txt' -append

}