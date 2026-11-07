//%attributes = {}
/*

tell word and excel to open, save and close document

open documents are identified by theier name
confirmation dialogs are surpressed for save and close
when number of documents reach 0, the app is terminated

*/

TEST_open_word 
TRACE:C157  //edit the doc here, dont't save it
TEST_close_word 
TEST_open_word 
TRACE:C157  //it should be saved anyway


TEST_open_excel 
TRACE:C157  //edit the doc here, dont't save it
TEST_close_excel 
TEST_open_excel 
TRACE:C157  //it should be saved anyway


