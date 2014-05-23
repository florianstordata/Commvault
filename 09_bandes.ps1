#SERVICE MANAGES - Copyright SOTRDATA 2014
#Script de Collecte d'infos Commvault
# $version : V1.1
# $initiate : 01/01/2014
# $revision : 17/04/2014
# $author : Firminhac Florian  (florian.firminhac@stordata.fr)
$Date = Get-Date
$asupv2="C:\ASUPV2"
$destination="D:\_Stordata\traitement-cv"
$7z="$asupv2\7za.exe"


 $directory=dir -Path $destination -Directory
 foreach ($dir in $directory){
 
  # on recupere le repertoire courant pour y enregistrer les fichiers 
$rep = "$destination\$dir"

If ( (Test-Path $rep\*StorageInformationReport*)) { 



# on declare le repertoire ou les fichiers generés seront mis 
$sources="$destination\$dir\sources"

# on verifie que le repertoire existe si non on le créé si oui on continue
If (-not (Test-Path $sources)) { New-Item -ItemType Directory -path $sources }

# on recupere le repertoire courant pour y enregistrer les fichiers 
$rep = "$destination\$dir"

# on verifie si le fichier destination exite si oui on le supprime sinon on continue
if(test-path $sources\09_bandes.xlsx) {remove-item $sources\09_bandes.xlsx}


# on parse le fichier StorageInformationReport* généré par Commvault et on ne recupere que les lignes qui ont la colonne library contenant Overland et Robots Bandes Washington
$processes = Import-Csv -header "Bar Code", "Media Type", "Description", "Library", "Storage Policy [Copy]", "Media Group", "Retain Data Until", "Location", "Container", "Prevent Export", "Exportable Time", "Last Export Time", "Side", "Status", "Last Read", "Last Write", "Total Data (MB)" "$rep\*StorageInformationReport*" -delimiter "`t" | Where {$_."Bar Code" -notmatch "magnetic"}| where {$_.Status -ne $null} | where {$_."Bar Code" -ne "total"} | where {$_."Bar Code" -ne "Bar Code"} | where {$_."Total Data (MB)" -ne $null }
#$processes = Import-Csv -header "Bar Code", "Library", "Storage Policy [Copy]", "Media Group", "Retain Data Until", "Location", "Container", "Prevent Export", "Exportable Time", "Last Export Time", "Side", "Status", "Last Read", "Last Write", "Total Data (MB)", "Usage and error" "$rep\*StorageInformationReport*" -delimiter "`t" | Where {$_."Bar Code" -notmatch "magnetic"}| where {$_.Status -ne $null}



# on specifie le fichier de sortie
# on boucle sur chaque ligne avec le status active
# on ecrit un tableau avec les colonnes qui nous interesse 
# on cache la fenetre Excel 
# on enregistre le fichier

$Excel = New-Object -ComObject excel.application 
$workbook = $Excel.workbooks.add() 
$xlout = "$sources\09_bandes.xlsx"
$i = 1 
foreach($process in $processes) 
{ 
 $excel.cells.item($i,1) = $process."Bar Code"
 $excel.cells.item($i,2) = $process."Media Type"
 $excel.cells.item($i,3) = $process."Library"
 $excel.cells.item($i,4) = $process."Location"
 $excel.cells.item($i,5) = $process."Media Group"
 $excel.cells.item($i,6) = $process."Status"
 $excel.cells.item($i,7) = $process."Last Export Time"
 
 $i++ 
} 
$Excel.visible = $false

$Workbook.SaveAs($xlout, 51)
$excel.Quit()
}
else 
{ echo "pas de fichier *StorageInformationReport* present"}

}