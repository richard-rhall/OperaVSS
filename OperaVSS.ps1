# OperaVSS - Richard Hall
$strFileName="FULL PATH HERE"
IF (Test-Path $strFileName){
# msg can be removed if wanted
msg * /SERVER:LT-0676 FILE EXISTS
# gwmi - Starts VSS Shadow Copy for Set Path.
(gwmi -list win32_shadowcopy).Create('F:\','ClientAccessible') | Out-File "Full path here with $lastvss.txt" -Width 120;
# $lastvss Gets date of last shadow copy and outputs blank text file with last shadow copy dd/tt as file name.
$lastvss = Get-Date(Get-CimInstance -ClassName Win32_shadowcopy | Sort-Object InstallDate | Select-Object -Last 1 -ExpandProperty 'InstallDate') -Format "dd-MM-yyyy hh-mm-ss"
Out-File "Full Path Here\Backup Success $lastvss.txt"
$renamepath = "Full Path Here"
Rename-Item -path $renamepath -newname "Backup Success $lastvss.txt"
# Writes successful VSS copy event to Applciation EventLog with ID of 1981
# To find me in Event Viewer, goto Application Logs, Search for
# EventId 1981 or Source OperaVSS or Level Information
Write-EventLog -LogName Application -Source OperaVSS -EntryType Information -EventId 1981 -Message 'OperaVSS Completed, Shadow Copy Successful'
}Else{
# Writes Failed VSS copy event to Applciation EventLog with ID of 1981
# To find me in Event Viewer, goto Application Logs, Search for;
# EventId 1981 or Source OperaVSS or Level Information
Write-EventLog -LogName Application -Source OperaVSS -EntryType Error -EventId 1981 -Message 'OperaVSS Completed, Shadow Copy Failed'
Out-File "Full Path Here"
Out-File "Full Path here\Backup Fail $lastvss.txt"
}
