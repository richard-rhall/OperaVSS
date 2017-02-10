# OperaVSS - Richard Hall
$strFileName="F:\richardscripts\watchdirectory\opera.txt"
IF (Test-Path $strFileName){
Rename-Item -Path "F:\richardscripts\watchdirectory\opera.txt" -NewName "working.txt"
Start-Sleep -s 5
# gwmi - Starts VSS Shadow Copy for Set Path.
(gwmi -list win32_shadowcopy).Create('F:\','ClientAccessible') | Out-File "F:\richardscripts\logfiles\Backup Success vssdebug.txt" -Width 120;
# $lastvss Gets date of last shadow copy and outputs blank text file with last shadow copy dd/tt as file name.
$lastvss = Get-Date(Get-CimInstance -ClassName Win32_shadowcopy | Sort-Object InstallDate | Select-Object -Last 1 -ExpandProperty 'InstallDate') -Format "dd-MM-yyyy hh-mm-ss"
Start-Sleep -s 3
Out-File "F:\richardscripts\logfiles\Backup Success $lastvss.txt"
Rename-Item -Path "F:\richardscripts\watchdirectory\working.txt" -NewName "Backup Success $lastvss.txt"
# Writes successful VSS copy event to Applciation EventLog with ID of 1981
# To find me in Event Viewer, goto Application Logs, Search for
# EventId 1981 or Source OperaVSS or Level Information
Write-EventLog -LogName Application -Source OperaVSS -EntryType Information -EventId 1981 -Message 'OperaVSS Completed, Shadow Copy Successful'
}Else{

$lastvss = Get-Date(Get-CimInstance -ClassName Win32_shadowcopy | Sort-Object InstallDate | Select-Object -Last 1 -ExpandProperty 'InstallDate') -Format "dd-MM-yyyy hh-mm-ss"
Out-File "F:\richardscripts\logfiles\No Backup Performed When ran, Last Successful VSS ran at $lastvss.txt"
} 
