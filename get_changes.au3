#include <Array.au3>
#include <File.au3>
Opt ("TrayAutoPause",0)
$aScrPath = ""&@ScriptDir&"\screenshots" ; Path to autotests folder
$aFilePath = ""&@ScriptDir&"\files" ; Path to Office folder in Programm Files\Microsoft Office
$aFileOutput = ""&@ScriptDir&"\result"
$aFormat_from = "xlsx"
$aFormat_to = "xls"
$aScrPath = _FileListToArray($aScrPath) ;get array of scr name


For $element In $aScrPath
	If StringInStr($element, "~$") == 0 Then   ;skip all files, if name started whith ~
		 $String = ""&$aFilePath&"\"&$element&""  ;path to file
		 $String = StringReplace($String,$aFormat_from, $aFormat_to )
		 FileCopy ($String, $aFileOutput)
		 ConsoleWrite($String & @LF)
   EndIf
Next




