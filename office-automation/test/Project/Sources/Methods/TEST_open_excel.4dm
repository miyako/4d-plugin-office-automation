//%attributes = {}
/*

tell excel to open document

*/

$file:=Folder:C1567(fk resources folder:K87:11).file("テスト.xlsx")

$params:=New object:C1471
$params.app:="excel"
$params.path:=$file.platformPath  //if omitted, save in temporary folder (mac)
$params.command:="open"

$status:=office automation ($params)