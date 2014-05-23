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

If ( (Test-Path $rep\*Nondiskmedia.xls)) { 



# on declare le repertoire ou les fichiers generés seront mis 
$sources="$destination\$dir\sources"

# on verifie que le repertoire existe si non on le créé si oui on continue
If (-not (Test-Path $sources)) { New-Item -ItemType Directory -path $sources }

# on recupere le repertoire courant pour y enregistrer les fichiers 
$rep = "$destination\$dir"

# on verifie si le fichier destination exite si oui on le supprime sinon on continue
if(test-path $sources\10_free.xlsx) {remove-item $sources\10_free.xlsx}

# on parse le fichier *Nondiskmedia.xls généré par Commvault et on ne recupere que les lignes qui ont la colonne library ne contenant pas le moto Total
$bla=Import-Csv "$rep\*Nondiskmedia.xls" -delimiter "`t" | where {$_.Library -notmatch "Total"}


# on specifie le fichier de sortie
# on boucle sur chaque ligne avec le status active
# on ecrit un tableau avec les colonnes qui nous interesse 
# on cache la fenetre Excel 
# on enregistre le fichier

$Excel = New-Object -ComObject excel.application 
$workbook = $Excel.workbooks.add() 


$xlout = "$sources\10_free.xlsx"
$i = 1 

foreach ($blah in $bla) {
# comme Commvault nous sort la date sous le format 18/02/2012 22:00:26  (Romance Standard Time                                                                                                                                                                                                                                          )
#on fait une boucle pour spliter la valeur en utilisant comme separateur "("
$bash=$blah."Estimated Aging Date".Split("(")

#powershell nous renvoit la valeur divisé 
#on redefini la valeur dans la boucle en ne prenant que la 1ere partie
$blah."Estimated Aging Date"=$bash[0]

#on continue le traitement comme avant
$excel.cells.item($i,1) = $blah.Library
$excel.cells.item($i,2) = $blah."Media Location"
$excel.cells.item($i,3) = $blah."Media"
$excel.cells.item($i,4) = $blah."Storage Policy"
$excel.cells.item($i,5) = $blah."Copy Name"
$excel.cells.item($i,6) = $blah."Estimated Aging Date"

 
 $i++
}

$Excel.visible = $false

$Workbook.SaveAs($xlout, 51)
$excel.Quit()

}
else 
{ echo "pas de fichier *Nondiskmedia.xls present"}

}
