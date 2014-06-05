# Execute a la suite les scripts powershell 
$Date = Get-Date

$datelog=Get-Date -UFormat %d-%m-%y
Start-Transcript -path "D:\_Stordata\log\$datelog-cv.log" -Append

# extract des cab generé par CV
$asupv2="C:\ASUPV2"
$destination="D:\_Stordata\traitement-cv"
$7z="$asupv2\7za.exe"
$scripts = "D:\_Stordata\_Scripts\Commvault"
# Extract des Cab generés par 

 $directory=dir -Path $destination -Directory
 foreach ($dir in $directory){


do {

$archives=Get-Item -Path $destination\$dir\* -Include *.cab

foreach ($archive in $archives.Name)
{
& "$7z" x "$destination\$dir\$archive" -o"$destination\$dir" -aoa
Remove-Item "$destination\$dir\$archive"
}}
while ($flagArchive=(Test-Path -Path $destination\$dir\* -Include *.cab))

}


echo "Execution du Script En_cours  (01/11)"
& $scripts\01_en_cours.ps1
echo "Execution du Script Events    (02/11)"
& $scripts\02_events.ps1
echo "Execution du Script Status    (03/11)"
& $scripts\03_status.ps1
echo "Execution du Script Size      (04/11)"
& $scripts\04_size.ps1
echo "Execution du Script Durée     (05/11)"
& $scripts\05_duree.ps1
echo "Execution du Script Echec     (06/11)"
& $scripts\06_echec.ps1
echo "Execution du Script Completed (07/11)"
& $scripts\07_completed.ps1
echo "Execution du Script dedup     (08/11)"
& $scripts\08_dedup.ps1
echo "Execution du Script Bandes    (09/11)"
& $scripts\09_bandes.ps1
if ($date.DayOfWeek -eq "Monday") {
echo "Execution du Script free      (10/11)"
& $scripts\10_free.ps1 }
echo "Execution du Script Resto     (11/11)"
& $scripts\11-resto.ps1

$elapsed=[math]::round(((Get-Date) - $Date).TotalMinutes,2)
echo "This report took $elapsed minutes to run all scripts."
Stop-Transcript
read-host "pressez une touche pour continuer"