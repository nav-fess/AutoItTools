#include <WordConstants.au3>
#include <File.au3>
#include <Word.au3>
#include <Excel.au3>
#include <PowerPoint.au3>

$Path_to_file_input = ""&@ScriptDir&"\files"
$Path_to_file_output = ""&@ScriptDir&"\result"

Func GetFileList($FilePath)
   $list = _FileListToArray($FilePath)
   _ArrayDelete($list, 0)    ;first name is clear
   Return($list)
EndFunc
;----------------------------------------------------------------------

Func GetFormatByPath($Path)
   ; Path - is a String(filename or file path) WITH FORMAT
   ; Return only format like string
Return(_PathSplit($Path, "","","","")[4])
EndFunc
;----------------------------------------------------------------------

Func PowerPoint($element)

   $oPowerPoint = _PPT_PowerPointApp(0)
      $Presentation =_PPT_PresentationOpen($oPowerPoint, $Path_to_file_input & "\" & $element)

   $text = StringReplace($Path_to_file_output&"\"&$element, ".ppt", ".pptx" )
   _PPT_PresentationSaveAs($Presentation,$text)
   _PPT_PresentationClose($Presentation)
   _PPT_PowerPointQuit($oPowerPoint)
	 ProcessClose("POWERPNT.EXE")

EndFunc
;----------------------------------------------------------------------

Func Excel($element)

   $Excel = _Excel_Open(False)

   $oExcelBook = _Excel_BookOpen($Excel,$Path_to_file_input & "\" & $element)

   $text = StringReplace($Path_to_file_output&"\"&$element, ".xls", ".xlsx" )
   _Excel_BookSaveAs($oExcelBook, $text, 51)  ;https://msdn.microsoft.com/en-us/vba/excel-vba/articles/xlfileformat-enumeration-excel

   _Excel_BookClose($oExcelBook)
   _Excel_Close($Excel)
   ProcessClose("EXCEL.EXE")

   ConsoleWrite("Converted " & $element & @CRLF)
EndFunc
;----------------------------------------------------------------------

Func Document($element)

   $oWord = _Word_Create(False)
   $oDoc =  _Word_DocOpen($oWord,$Path_to_file_input & "\" & $element)

    $text = StringReplace($Path_to_file_output&"\"&$element, ".doc", ".docx" )
	_Word_DocSaveAs($oDoc, $text,16)

   _Word_DocClose($oDoc)
   _Word_Quit($oWord,0)
   ProcessClose("WINWORD.EXE")

   ConsoleWrite("Converted " & $element & @CRLF)
EndFunc
;----------------------------------------------------------------------

Func Main()

 $FileList = GetFileList($Path_to_file_input)
 $format =  GetFormatByPath($FileList[0])

Switch $format
	  Case ".doc"
		 For $element In $FileList
			Document($element)
		 Next
	  Case ".xls"
		 For $element In $FileList
			Excel($element)
		 Next
	   Case ".ppt"
	   	   For $element In $FileList
			PowerPoint($element)
		 Next
	EndSwitch



EndFunc
;----------------------------------------------------------------------

Main()






