//%attributes = {}
/*

tell word to save and close document

*/

$file:=Folder:C1567(fk resources folder:K87:11).file("テスト.docx")

$params:=New object:C1471
$params.app:="word"
$params.command:="close"
$params.name:=$file.name+$file.extension  //specify the name, not a path

$status:=office automation ($params)