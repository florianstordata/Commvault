$pdfs=Get-ChildItem D:\_Stordata\Template\_CR -Recurse -filter *0525.pdf

foreach ($pdf in $pdfs) {
$pdf
$pdf.BaseName
}