#include <Array.au3>
#include <File.au3>
#include <GUIConstants.au3>

$Path_to_x2t = ""&@ScriptDir&"\x2t.exe" ;$Pach_to_x2t like a c:\autotests\x2t.exe
$Path_to_file_input = ""&@ScriptDir&"\files" ;Folder with files without last backslash
$Path_to_file_output = ""&@ScriptDir&"\result" ;Empty folder for file result without last backslash. Wil be clear before start convertion
$Path_to_fonts = ""&@WindowsDir&"\Fonts" ; Folder with fonts
$Path_to_dll = ""&@ScriptDir&"\libs" ;Folder with files without last backslash
;----------------------------------------------------------------------------------------------------------

$Format_to = "docx"
$NumberOfFile = 0
Func ConvertFile($filename, $format_to)
   ConsoleWrite("Convert file "&$filename&" from format "&$format_to&"" & @LF)
   $Filepath_in = ""&$Path_to_file_input&"\"&$filename&""
   $Filepath_out = ""&$Path_to_file_output&"\"&$filename&"."&$format_to&""
   ConsoleWrite(""&$Path_to_x2t&" """&$Filepath_in&""" """&$Filepath_out&"""" & @LF)
   if WinExists("x2t") Then
	  WinClose("x2t")

   EndIf

   $I=RunWait(""&$Path_to_x2t&" """&$Filepath_in&""" """&$Filepath_out&""" "&$Path_to_fonts&"", "", @SW_HIDE)

   ConsoleWrite('Code=' & $I & @CRLF)

EndFunc

Func GetFileList($FilePath)
   $list = _FileListToArray($FilePath)
   _ArrayDelete($list, 0)    ;first name is clear
   Return($list)
EndFunc

Func DeleteTempFile($FilePath)
   $Filelist = GetFileList($FilePath)
   ConsoleWrite($Filelist[0])
   if $Filelist[0] == 0 Then Return(0)
   For $element In $FileList
	  ConsoleWrite("delete file "&$element&"")
		  ConsoleWrite(FileGetSize($element) & @LF)
	  	  ConsoleWrite(($element) & @LF)
	  if (FileGetSize(""&$Path_to_file_input&"\"&$element&"")) == 0 Then
		 ConsoleWrite("Delete file "&$element&" because its empty" & @LF)
		 FileDelete(""&$Path_to_file_input&"\"&$element&"")
	  EndIf
   Next
   Return _FileListToArray($Path_to_file_input)
EndFunc

Func ConvertAll($Progressbar)

if CheckDLL() Then

;#comments-start
   ClearResultDir()

   DeleteTempFile($Path_to_file_input)

   $FileList = GetFileList($Path_to_file_input)

   Local $i = 1
   Local $Count = UBound($FileList, 1)
   ConsoleWrite($Count & @LF)
	  For $element In $FileList
		 ConvertFile(""&$element&"", $Format_to)
		 GUICtrlSetData($Progressbar, $i/$Count * 100)
		 ;Sleep(4000)
		 $i = $i+1
	  Next
	  GUICtrlSetData($Progressbar, 0)
;#comments-end

   GetFileDiff()

EndIf

EndFunc

Func CheckDLL()

  Local $aArrayDLL[10] = ['ascdocumentscore.dll','DjVuFile.dll','doctrenderer.dll','HtmlFile.dll','HtmlRenderer.dll','icudt55.dll','icuuc55.dll','PdfReader.dll','PdfWriter.dll','XpsFile.dll']
  Local $aResult = []
  _ArrayDelete($aResult,0)

  $aArrayExistDLL= GetFileList($Path_to_dll)

   For $element In $aArrayDLL
	  If _ArraySearch($aArrayExistDLL, $element) == -1 Then
		 _ArrayAdd($aResult, $element)
	  EndIf
   Next

   If( UBound($aResult) <> 0) Then
	  MsgBox(0,"Проверка DLL","File Not Found:"& @CRLF & _ArrayToString($aResult,@CRLF))
	  Return 0
   Else
	   ;MsgBox(0,"Проверка DLL","All dll Exist. Start")
	   Return 1
   EndIf

EndFunc

Func ClearResultDir()
   DirRemove ($Path_to_file_output, 1)
   DirCreate ($Path_to_file_output)
   Local $NotConv = ""&$Path_to_file_output&"\notconverted"
   DirCreate ($NotConv)
EndFunc

Func GetFileDiff()

   ConsoleWrite("TIME GetFileList($Path_to_file_input" & @CRLF )
   Sleep(15000)

   $AllFilesInput = GetFileList($Path_to_file_input)
   $AllFilesOutput = GetFileList($Path_to_file_output)

   $aNumberDeleteFile=_ArraySearch($AllFilesOutput,'notconverted')
   _ArrayDelete($AllFilesOutput,$aNumberDeleteFile)

   Local $Result = []
   _ArrayDelete($Result,0)

     	  ;#comments-start

	  ConsoleWrite('$AllFilesInput' & @CRLF)
	  For $el in $AllFilesInput
		 ConsoleWrite('FI=' & $el & @CRLF)
	   Next

	  ConsoleWrite('$AllFilesOutput' & @CRLF)
	   For $el in $AllFilesOutput
		 ConsoleWrite('$FO=' & $el & @CRLF)
	  Next

   For $element In $AllFilesInput
	  $str = $element & '.' & $Format_to

	  If _ArraySearch($AllFilesOutput, $str) == -1 Then
			ConsoleWrite("@error=" & @error & @CRLF)
			ConsoleWrite("@error=" & @error & @CRLF)
			ConsoleWrite("$str = " & $str & @CRLF)
		 _ArrayAdd($Result, $element)
	  EndIf

   Next
    _ArrayDisplay($Result)
    CopyFiles($Result)
	  ;#comments-end
EndFunc

Func CopyFiles($List)

   ConsoleWrite("CopyFiles" & @CRLF)

   For $element In $List
	  if $element <> '' Then
		  FileCopy(""&$Path_to_file_input&"\"&$element&"", ""&$Path_to_file_output&"\notconverted")
		  ConsoleWrite("FilesCopy11" & @CRLF)
	  EndIf
   Next

EndFunc

; -----------------------------------------GUI--------------------------------------------
 Example()

Func Example()
    ; Create a GUI with various controls.
    Local $hGUI = GUICreate("Convertion settings", 400, 600)
    ; Create a button control.
	$InputLabul = GUICtrlCreateLabel ("Input folder path  "&$Path_to_file_input&"", 20, 20)
    $OutputLabel = GUICtrlCreateLabel ("Output folder path  "&$Path_to_file_output&"", 20, 40)
	$X2tLabel = GUICtrlCreateLabel ("Path to x2t  "&$Path_to_x2t&"", 20, 70)
	GUICtrlCreateLabel ("Path to fonts  "&$Path_to_fonts&"", 20, 90)
	GUICtrlCreateLabel ("Convert to", 20, 124)
    $StartButton = GUICtrlCreateButton("Start", 50, 500, 300, 75)
	$hCombo = GUICtrlCreateCombo("", 100, 120, 200, 30)
    GUICtrlSetData($hCombo, "docx|rtf|txt|xls|xlsx|pptx|odt|ods|odp|doct", "docx")
	Local $idProgressbar = GUICtrlCreateProgress(20, 200, 360, 20, $PBS_SMOOTH)
    $Label = GUICtrlCreateLabel("Complite!", 170, 220, 80) ; first cell 70 width
	GUICtrlSetState($Label, $GUI_HIDE)
    ; Display the GUI.
    GUISetState(@SW_SHOW, $hGUI)
    Local $iPID = 0
    ; Loop until the user exits.
    While 1
        Switch GUIGetMsg()
            Case $GUI_EVENT_CLOSE
                ExitLoop

            Case $StartButton
                ; Start converte all files from
				$iPID = ConvertAll($idProgressbar)
				GUICtrlSetState($Label, $GUI_SHOW)
				ConsoleWrite("Convert all" & @LF)
			 Case $hCombo
				$Format_to =  GUICtrlRead($hCombo)
				ConsoleWrite("Format change " & $Format_to & "" & @LF)
        EndSwitch
    WEnd
    GUIDelete($hGUI)
    If $iPID Then ProcessClose($iPID)
EndFunc
