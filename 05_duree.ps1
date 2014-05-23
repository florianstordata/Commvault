#SERVICE MANAGES - Copyright SOTRDATA 2014
#Script de Collecte d'infos Commvault
# $version : V1.0 
# $initiate : 01/01/2014
# $author : Firminhac Florian (florian.firminhac@stordata.fr)


# on declare le repertoire ou les fichiers generés seront mis 
$sources="sources"

# on verifie que le repertoire existe si non on le créé si oui on continue
If (-not (Test-Path $sources)) { New-Item -ItemType Directory -Name $sources }

# on recupere le repertoire courant pour y enregistrer les fichiers 
$rep = (Get-Location).path

# on verifie si le fichier destination exite si oui on le supprime sinon on continue
if(test-path $sources\05_duree.xlsx) {remove-item $sources\05_duree.xlsx}

# on parse le fichier *details.xls généré par Commvault et on ne recupere que les lignes qui ont la colonne status ne contenant pas N/A
$processes = Import-Csv BackupJobSummaryReport*details.xls -delimiter "`t" | where {$_."Status" -ne "N/A"}


# on specifie le fichier de sortie
# on boucle sur chaque ligne avec le status active
# on ecrit un tableau avec les colonnes qui nous interesse 
# on cache la fenetre Excel 
# on enregistre le fichier

$Excel = New-Object -ComObject excel.application 
$workbook = $Excel.workbooks.add() 

$xlout = "$($rep)\$sources\05_duree.xlsx"
$i = 1 
foreach($process in $processes) 
{ 
 $excel.cells.item($i,1) = $process."Client"
 $excel.cells.item($i,2) = $process."Start Time"
 $excel.cells.item($i,3) = $process."End Time or Current Phase"
 $excel.cells.item($i,4) = $process."Transfer Time"
 
 
 $i++ 
} 
$Excel.visible = $false

$Workbook.SaveAs($xlout, 51)
$excel.Quit()