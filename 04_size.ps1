#SERVICE MANAGES - Copyright SOTRDATA 2014
#Script de Collecte d'infos Commvault
# $version : V1.1
# $initiate : 01/01/2014
# $revision : 17/04/2014
# $author : Firminhac Florian  (florian.firminhac@stordata.fr)

$asupv2="C:\ASUPV2"
$destination="D:\_Stordata\traitement-cv"
$7z="$asupv2\7za.exe"


 $directory=dir -Path $destination -Directory
 foreach ($dir in $directory){
 
  # on recupere le repertoire courant pour y enregistrer les fichiers 
$rep = "$destination\$dir"

If ( (Test-Path $rep\*BackupJobSummaryReport*details.xls)) { 



# on declare le repertoire ou les fichiers generés seront mis 
$sources="$destination\$dir\sources"

# on verifie que le repertoire existe si non on le créé si oui on continue
If (-not (Test-Path $sources)) { New-Item -ItemType Directory -path $sources }

# on recupere le repertoire courant pour y enregistrer les fichiers 
$rep = "$destination\$dir"

# on verifie si le fichier destination exite si oui on le supprime sinon on continue
if(test-path $sources\04_size.xlsx) {remove-item $sources\04_size.xlsx}

# on parse le fichier *details.xls généré par Commvault 
$processes = Import-Csv "$rep\*BackupJobSummaryReport*details.xls" -delimiter "`t"


# on specifie le fichier de sortie
# on boucle sur chaque ligne avec le status active
# on ecrit un tableau avec les colonnes qui nous interesse 
# on cache la fenetre Excel 
# on enregistre le fichier

$Excel = New-Object -ComObject excel.application 
$workbook = $Excel.workbooks.add() 


$xlout = "$sources\04_size.xlsx"
$i = 1 
foreach($process in $processes) 
{

 $excel.cells.item($i,1) = $process.Client
 $excel.cells.item($i,2) = $process."Size of Application"
 $excel.cells.item($i,3) = $process."Transfer Time"
 $excel.cells.item($i,4) = $process."Data Written"
 $excel.cells.item($i,5) = $process."Start Time"
 $excel.cells.item($i,6) = $process."End Time or Current Phase"
 
 $i++ 
} 
$Excel.visible = $false

$Workbook.SaveAs($xlout, 51)
$excel.Quit()

}
else 
{ echo "pas de fichier *BackupJobSummaryReport*details.xls present"}

}