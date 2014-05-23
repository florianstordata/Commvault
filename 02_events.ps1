#SERVICE MANAGES - Copyright SOTRDATA 2014
#Script de Collecte d'infos Commvault
# $version : V1.0 
# $initiate : 01/01/2014
# $author : Firminhac Florian  (florian.firminhac@stordata.fr)


# on declare le repertoire ou les fichiers generés seront mis 
$sources="sources"

# on verifie que le repertoire existe si non on le créé si oui on continue
If (-not (Test-Path $sources)) { New-Item -ItemType Directory -Name $sources }

# on recupere le repertoire courant pour y enregistrer les fichiers 
$rep = (Get-Location).path

# on verifie si le fichier destination exite si oui on le supprime sinon on continue
if(test-path $sources\02_events.xlsx) {remove-item $sources\02_events.xlsx}

# on parse le fichier eventreport* généré par Commvault et on  ne recupere que les lignes qui ont une description
$processes = Import-Csv -header "Severity", "Event ID", "Job ID", "Time", "Program", "Computer", "Event Code", "Description" -delimiter "`t" EventReport* | where {$_.Description -ne $null}
$Excel = New-Object -ComObject excel.application 
$workbook = $Excel.workbooks.add() 


# on specifie le fichier de sortie
# on boucle sur chaque ligne avec le status active
# on ecrit un tableau avec les colonnes qui nous interesse 
# on cache la fenetre Excel 
# on enregistre le fichier

$xlout = "$($rep)\$sources\02_events.xlsx"
$i = 1 
foreach($process in $processes) 
{ 
 $excel.cells.item($i,1) = $process.Severity
 $excel.cells.item($i,2) = $process.Time
 $excel.cells.item($i,3) = $process.Program
 $excel.cells.item($i,4) = $process.Computer
 $excel.cells.item($i,5) = $process.Description
 $i++ 
} 
$Excel.visible = $false

$Workbook.SaveAs($xlout, 51)
$excel.Quit()


