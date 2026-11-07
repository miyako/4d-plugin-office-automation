//%attributes = {}
/*

tell word to open document

*/

$file:=Folder:C1567(fk resources folder:K87:11).file("テスト.docx")

$params:=New object:C1471
$params.app:="word"
$params.path:=$file.platformPath  //if omitted, save in temporary folder (mac)
$params.command:="open"

$status:=office automation ($params)