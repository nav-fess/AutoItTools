#include <Array.au3>
#include <File.au3>
#include <GUIConstantsEx.au3>
Opt ("TrayAutoPause",0)
$aMainPath = "C:\" ; Path to autotests folder
$ax2tPath = "C:\Program Files\AVS4YOU\AVSDocumentEditor\converter\x2t"

$a = Run(""&$ax2tPath&"", '',  @SW_MINIMIZE, $STDERR_CHILD + $STDOUT_CHILD)


Local $output
Local $log = ''
While 1
    $output = StdoutRead($a)
    If @error Then ExitLoop
    $log = $log & $output
 Wend
Example()
Func Example()
    GUICreate("Version Log")
    GUICtrlCreateLabel($log, 0, 0, 500)
    GUISetState(@SW_SHOW)
    While 1
        Switch GUIGetMsg()
            Case $GUI_EVENT_CLOSE
                ExitLoop

        EndSwitch
    WEnd
EndFunc   ;==>Example