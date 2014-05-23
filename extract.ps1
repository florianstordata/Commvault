$asupv2="C:\ASUPV2"
$destination="D:\_Stordata\traitement-cv"
$7z="$asupv2\7za.exe"

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