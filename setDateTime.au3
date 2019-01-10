#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <GUIListBox.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <GuiDateTimePicker.au3>
#include <Date.au3>
#include <File.au3>
#include <FileConstants.au3>
#include <MsgBoxConstants.au3>
#include <WinAPIFiles.au3>
#Region ### START Koda GUI section ### Form=C:\Dropbox\_Programs\AutoIt\DateTime_SetSystemTime\Form1.kxf

$Form1 = GUICreate("Set System Time", 300, 392, 542, 459)
$Txt = GUICtrlCreateInput("2019-12-30 23:59:59", 8, 48, 214, 21)
$Reset = GUICtrlCreateButton("Reset - Use Live date-time", 8, 8, 214, 25)
$List = GUICtrlCreateList("", 8, 160, 294, 201)
$Add = GUICtrlCreateButton("Add", 8, 95, 214, 25)
$Delete= GUICtrlCreateButton("Delete", 8, 128, 214, 25)
$Format = GUICtrlCreateLabel("Format", 8, 368, 36, 17)

GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###

If IsAdmin() = 0 Then
    MsgBox($MB_SYSTEMMODAL, "Run program as Administrator", "To change the system time, this program needs to run as an administrator. Righclick on the program .exe file, and 'Run as administrator'")
EndIf

readFile()
While 1
	Switch GUIGetMsg()
	  Case $GUI_EVENT_CLOSE
		 reset()
		 writeFile()
		 Exit
	  Case $Reset
		 reset()
	  Case $Add
		add()
	  Case $Delete
		delete()
	  Case $List
		ListBoxClickedSetDateTime()
	EndSwitch
WEnd

Func add()

   If _DateIsValid(GUICtrlRead($Txt)) Then
		; this is a valid date

	  ; add string to listbox
	  $str = GUICtrlRead($Txt) & " |  " & getPrettyDateString(  GUICtrlRead($Txt)  )
	  _GUICtrlListBox_AddString($List,  $str)
	   writeFile()
   Else
	   MsgBox($MB_SYSTEMMODAL, "Valid Date", "The specified date is invalid.")
   EndIf

EndFunc


Func delete()
   If IsAdmin() = 0 Then
    MsgBox($MB_SYSTEMMODAL, "Run program as Administrator", "To change the system time, this program needs to run as an administrator. Righclick on the program .exe file, and 'Run as administrator'")
   EndIf

   $z = _GUICtrlListBox_GetCurSel ($List)
   _GUICtrlListBox_DeleteString ($List, $z)
   writeFile()
EndFunc


Func reset()
   ; enable windows live time sync (sets your windows system clock to normal time)
   			$CMD = "w32tm /resync"
			RunWait(@ComSpec & " /c " & $CMD)
EndFunc

Func ListBoxClickedSetDateTime()

   $s = GUICtrlRead($List)
	  ;$s = _GUICtrlListBox_GetCurSel($List)
ConsoleWrite($s)
	  ; set date
	  $CMD = "date " & day($s) & "-" & month($s) & "-" & year($s)
	  ConsoleWrite($CMD)
	  RunWait(@ComSpec & " /c " & $CMD)

	  ; set time
	  $CMD = "time " & hour($s) & ":" & minute($s)
	  RunWait(@ComSpec & " /c " & $CMD)

EndFunc


;-----------------------------------------------------------

Func writeFile()
   $file = FileOpen(@ScriptDir & "\config.txt", 2)
	If $file = -1 Then
	 MsgBox($MB_SYSTEMMODAL, "", "An error occurred whilst writing the temporary file. Check permissions.")
	 Return False
   EndIf
   $numb_of_el_2 = _GUICtrlListBox_GetCount($List)
        For $ir = 0 To $numb_of_el_2 Step 1
            If _GUICtrlListBox_GetText($List, $ir) <> "" Then ; Checking for 0 doesn't work with strings, check for an empty string instead
                FileWriteLine($file, _GUICtrlListBox_GetText($List, $ir))
            EndIf
		 Next
   FileClose($file)
EndFunc

Func readFile()
   $file = FileRead(@ScriptDir & "\config.txt")
   If $file = -1 Then
	 MsgBox($MB_SYSTEMMODAL, "", "An error occurred whilst opening config.txt. Check permissions.")
	 Return False
  Else
	 For $ir = 1 To _FileCountLines ( @ScriptDir & "\config.txt" ) Step 1
			 _GUICtrlListBox_AddString($List,  FileReadLine( @ScriptDir & "\config.txt", $ir )  )
	  Next
  EndIf

   FileClose($file)
EndFunc

;-----------------------------------------------------------

Func getPrettyMonth($s)
	  Return _DateToMonth ( month($s), $DMW_LOCALE_SHORTNAME )
EndFunc

Func getPrettyDay($s)
  Return _DateDayOfWeek (  _DateToDayOfWeek (  year($s), month($s), day($s) ), $DMW_LOCALE_LONGNAME)
EndFunc

Func getPrettyDateString($s)
   Return getPrettyDay($s) & " " & month($s) & ". " & getPrettyMonth($s) & " " & year($s) & ", kl " & hour($s) & ":" & minute($s)
EndFunc

;----------------------------------
Func year($s)
   Return StringLeft ( $s, 4 )
EndFunc

Func month($s)
   $s = StringTrimLeft ( $s, 5 )
   $s = StringLeft ( $s, 2 )
   Return $s
EndFunc

Func day($s)
   $s = StringTrimLeft ( $s, 8 )
   $s = StringLeft ( $s, 2 )
      Return $s
EndFunc

Func hour($s)
   $s = StringTrimLeft ( $s, 11 )
   $s = StringLeft ( $s, 2 )
   Return $s
EndFunc

Func minute($s)
   $s = StringTrimLeft ( $s, 14 )
   $s = StringLeft ( $s, 2 )
   Return $s
EndFunc

Func second($s)
   $s = StringTrimLeft ( $s, 17 )
   $s = StringLeft ( $s, 2 )
   Return $s
EndFunc

;https://www.autoitscript.com/autoit3/docs/libfunctions/_GUICtrlListBox_ClickItem.htm