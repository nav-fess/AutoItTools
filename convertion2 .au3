#include <Array.au3>
#include <File.au3>
#include <GUIConstants.au3>
#include <GuiListView.au3>

$Path_to_x2t = ""&@ScriptDir&"\x2t.exe" ;$Pach_to_x2t like a c:\autotests\x2t.exe
$Path_to_file_input = ""&@ScriptDir&"\files" ;Folder with files without last backslash
$Path_to_file_output = ""&@ScriptDir&"\result" ;Empty folder for file result without last backslash. Wil be clear before start convertion
$Path_to_fonts = ""&@WindowsDir&"\Fonts" ; Folder with fonts
$Path_to_dll = ""&@ScriptDir&"\libs" ;Folder with files without last backslash
;----------------------------------------------------------------------------------------------------------

Global $ListItems[1]=1

$Format_to = "docx"
$NumberOfFile = 0


Func ConvertFile($filename, $format_to)
   ConsoleWrite("Convert file "&$filename&" from format "&$format_to&"" & @LF)
   $Filepath_in = ""&$Path_to_file_input&"\"&$filename&""
   $Filepath_out = ""&$Path_to_file_output&"\"&$filename&"."&$format_to&""
   ConsoleWrite(""&$Path_to_x2t&" """&$Filepath_in&""" """&$Filepath_out&"""" & @LF)
   $I=RunWait(""&$Path_to_x2t&" """&$Filepath_in&""" """&$Filepath_out&""" "&$Path_to_fonts&"", "", @SW_HIDE)

   ;Sleep(2000)

   ConsoleWrite('Code=' & $I & @CRLF)

EndFunc


Func GetFileDiff()

   $AllFilesInput = GetFileList($Path_to_file_input)
   $AllFilesOutput = GetFileList($Path_to_file_output)

   $aNumberDeleteFile=_ArraySearch($AllFilesOutput,'notconverted')
   _ArrayDelete($AllFilesOutput,$aNumberDeleteFile)

   Local $Result = []
   _ArrayDelete($Result,0)

   For $element In $AllFilesInput
	  $str = $element & '.' & $Format_to

	  If _ArraySearch($AllFilesOutput, $str) == -1 Then
		 _ArrayAdd($Result, $element)
	  EndIf

   Next
    _ArrayDisplay($Result)
    CopyFiles($Result)

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


Func ClearResultDir($FolderOut)
   DirRemove ($Path_to_file_output & '\' & $FolderOut, 1)
   DirCreate ($Path_to_file_output & '\' & $FolderOut)
   Local $NotConv = ""&$Path_to_file_output & '\' & $FolderOut & "\notconverted"
   DirCreate ($NotConv)
EndFunc


Func ConvertAll($FolderIn, $FolderOut, $Progressbar)

   ClearResultDir($FolderIn & ' to ' & $FolderOut)

   DeleteTempFile($Path_to_file_input & '\' & $FolderIn )

   $FileList = GetFileList($Path_to_file_input & '\' & $FolderIn)

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

   ;GetFileDiff() ДОДЕЛАТЬ

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


Func CopyFiles($List)

   ConsoleWrite("CopyFiles" & @CRLF)

   For $element In $List
	  if $element <> '' Then
		  FileCopy(""&$Path_to_file_input&"\"&$element&"", ""&$Path_to_file_output&"\notconverted")
	  EndIf
   Next

EndFunc


Func CreateGroup($FormatsString, $xStart, $yStart)

    Local $hGuiGroup	;$hGuiGroup[2]
   Local $Formats =StringSplit($FormatsString ,'|')
   _ArrayDelete($Formats,0)

   GUICtrlCreateGroup("Format " & $Formats[0] , $xStart, $yStart, 170, 20 * (UBound($Formats)+1.5), $BS_GROUPBOX)	;Start group

   $hGuiGroup = GUICtrlCreateCheckbox($Formats[0], $xStart+15, $yStart+15, 50, 25)

   $ListView=GUICtrlCreateListView("convert to", $xStart + 80, $yStart + 15, 80, 20 * (0.5 + UBound($Formats)), -1, $LVS_EX_CHECKBOXES)

   For $ListItem = 1 To UBound($Formats)-1
	    GUICtrlCreateListViewItem($Formats[$ListItem], $ListView)
   Next

   GUICtrlCreateGroup("", -99, -99, 1, 1)			;End group
   Return $hGuiGroup

EndFunc



Func RezultConvertion($ProgressBar, $Discriptors )
;#comments-start
	  	  $FoldersIn = GetFileList($Path_to_file_input)

		 ;if ( CheckDLL() ) Then

			For $NumberGroup = 0 To Ubound($Discriptors)-1

			   if(GUICtrlRead($Discriptors[$NumberGroup][0]) == 1) Then		; get state Checkbox

				  $CheckBoxData = GUICtrlRead($Discriptors[$NumberGroup][0],1)
				  $ListViewDescriptor = $Discriptors[$NumberGroup][1]
				  $CountItemListView = _GUICtrlListView_GetItemCount($ListViewDescriptor)
				  $aCountCheckIn  = 0

				   For $item = 0 to $CountItemListView-1
					  if(_GUICtrlListView_GetItemChecked($ListViewDescriptor, $item)) Then

						$aCountCheckIn +=1
						$ListViewData = _GUICtrlListView_GetItemText($ListViewDescriptor ,$item)

						if (_ArraySearch($FoldersIn,$CheckBoxData) <> -1) Then
								 ConvertAll($CheckBoxData, $ListViewData ,$ProgressBar)
							  Else
								 ContinueLoop
						EndIf

					 EndIf
				  Next

				  if($aCountCheckIn == 0) Then
					 MsgBox(0,"Ошибка выбора","не выбраны пункты для конвертации " & $CheckBoxData & @CRLF)
				  EndIf

			   EndIf

			Next
	 ;EndIf
;#comments-end
EndFunc

; -----------------------------------------GUI--------------------------------------------

Example()


Func Example()
    ; Create a GUI with various controls.
    Local $hGUI = GUICreate("Convertion settings", 400, 700)
    ; Create a button control.
	$InputLabul = GUICtrlCreateLabel ("Input folder path  "&$Path_to_file_input&"", 20, 20)
    $OutputLabel = GUICtrlCreateLabel ("Output folder path  "&$Path_to_file_output&"", 20, 40)
	$X2tLabel = GUICtrlCreateLabel ("Path to x2t  "&$Path_to_x2t&"", 20, 70)
	GUICtrlCreateLabel ("Path to fonts  "&$Path_to_fonts&"", 20, 90)

   Local $MassivConvertation[]= ['docx|doct|odt|rtf', 'doc|docx', 'rtf|docx', 'xlsx|ods', 'xls|xlsx','pptx|odp', 'ppt|pptx', 'odt|docx','ods|xlsx','odp|pptx']

   Local $IdGroup [10][2]
   $yStartGroup = 0

	  For $item=0 To UBound($MassivConvertation)-1
		 if Mod($item ,2) <> 0 Then
			$xStartGroup = 220
		 Else
			$xStartGroup = 15
			if($item == 1) Then
			   $yStartGroup+=125
			Else
			   $yStartGroup+=110
			EndIf
		 EndIf

		 $hGuiGroups = CreateGroup($MassivConvertation[$item],$xStartGroup,$yStartGroup)
		 $hGuiCheckBox = 0
		 $IdGroup[$item][$hGuiCheckBox] = $hGuiGroups
		 $hGuiCheckBox = 1
		 $IdGroup[$item][$hGuiCheckBox] = $hGuiGroups+1

	  Next

   $StartButton = GUICtrlCreateButton("Start", 100, $yStartGroup + 80 , 200, 50)

	Local $idProgressbar = GUICtrlCreateProgress(210, 75, 180, 20, $PBS_SMOOTH)
    $Label = GUICtrlCreateLabel("Complite!", 170, 220, 80) ; first cell 70 width

	GUICtrlSetState($Label, $GUI_HIDE)

    GUISetState(@SW_SHOW, $hGUI) ; Display the GUI.

    Local $iPID = 0
    ; Loop until the user exits.
    While 1
        Switch GUIGetMsg()
            Case $GUI_EVENT_CLOSE
                ExitLoop

            Case $StartButton
                ;Start converte all files from
				RezultConvertion($idProgressbar, $IdGroup)
				;$iPID = ConvertAll($idProgressbar)
				GUICtrlSetState($Label, $GUI_SHOW)
				ConsoleWrite("Convert all" & @CRLF)

			 ;Case $idCheckbox
;~ 				$nItem = _GUICtrlListView_GetItemCount($ListView)
;~ 				For $item=0 to $nItem
;~ 				   if(_GUICtrlListView_GetItemChecked($ListView, $item)) Then
;~ 					 _ArrayAdd($ListItems, _GUICtrlListView_GetItemText($ListView,$item))
;~ 				  EndIf
;~ 				 _ArrayDelete($ListItems,1)
;~ 			   Next

			   _ArrayDisplay($ListItems)
				;_ArrayDisplay($ListItems)
        EndSwitch
    WEnd
    GUIDelete($hGUI)
    If $iPID Then ProcessClose($iPID)
EndFunc