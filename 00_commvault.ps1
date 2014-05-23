# Execute a la suite les scripts powershell 
$Date = Get-Date

echo "Execution du Script En_cours"
.\01_en_cours.ps1
echo "Execution du Script Events"
.\02_events.ps1
echo "Execution du Script Status"
.\03_status.ps1
echo "Execution du Script Size"
.\04_size.ps1
echo "Execution du Script Durée"
.\05_duree.ps1
echo "Execution du Script Echec"
.\06_echec.ps1
echo "Execution du Script Completed"
.\07_completed.ps1
echo "Execution du Script dedup"
.\08_dedup.ps1
echo "Execution du Script Bandes"
.\09_bandes.ps1
if ($date.DayOfWeek -eq "Monday") {
echo "Execution du Script free"
.\10_free.ps1}

# ajout d'un job pour les restauration si un fichier restauration et disponible 
If ( (Test-Path RestoreJobSummaryReport*.xls)) { 
echo "Execution du Script Resto"
.\11-resto.ps1 }

$elapsed=[math]::round(((Get-Date) - $Date).TotalMinutes,2)
echo "This report took $elapsed minutes to run all scripts."

read-host "pressez une touche pour continuer"