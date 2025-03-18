#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=edd_icon.ico
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#include <Date.au3>
#include <GUIConstantsEx.au3>

Global $emailAddresses[10] = ["Example e-mail 1", "Example e-mail 2", "Example e-mail 3", "Example e-mail 4", "Example e-mail 5)", _
"Example e-mail 6", "Example e-mail 7", "Example e-mail 8", "Example e-mail 9", "Example e-mail 10"]
Global $email = ""

#Region TEMPLATES
	Global $template_1 = "Example text for template_1"

	Global $template_2 = "Example text for template_2"

	Global $template_3 = "Example text for template_3"

	Global $template_4 = "Example text for template_4"

	Global $template_5 = "Example text for template_5"

	Global $template_6 = "Example text for template_6"

	Global $template_7 = "Example text for template_7"

	Global $template_8 = "Example text for template_8"

	Global $template_9 = "Example text for template_9"

	Global $template_10 = "Example text for template_10"
#EndRegion

; Create GUI
$firstWindowGUI = GUICreate("Name", 200, 260)
$firstWindowState = GUISetState(@SW_SHOW)

; Create radio buttons
Global $names[10] = ["Example name 1", "Example name 2", "Example name 3", "Example name 4", "Example name 5", _
"Example name 6", "Example name 7", "Example name 8", "Example name 9", "Example name 10"]
Global $radioButtons[10]
For $i = 0 To UBound($names) - 1
    $radioButtons[$i] = GUICtrlCreateRadio($names[$i], 10, 10 + ($i * 20), 200, 20)
Next

; Create OK button
$okButton = GUICtrlCreateButton("OK", 60, 220, 80, 30)

While 1
    Switch GUIGetMsg()
		Case $okButton
			For $i = 0 To UBound($radioButtons) - 1
				If GUICtrlRead($radioButtons[$i]) = $GUI_CHECKED Then
					$email = $emailAddresses[$i]
					GUIDelete($firstWindowGUI)
				EndIf
			Next
			ExitLoop
		Case $GUI_EVENT_CLOSE
			Exit
	EndSwitch
WEnd

; Create and activate GUI
$secondWindowGUI = GUICreate("Templates", 250, 245)
$secondWindowState = GUISetState(@SW_SHOW)

; Creating "paste" buttons
$paste_0 = GUICtrlCreateButton("Paste", 10, 11, 40, 20)
$paste_1 = GUICtrlCreateButton("Paste", 10, 31, 40, 20)
$paste_2 = GUICtrlCreateButton("Paste", 10, 51, 40, 20)
$paste_3 = GUICtrlCreateButton("Paste", 10, 71, 40, 20)
$paste_4 = GUICtrlCreateButton("Paste", 10, 91, 40, 20)
$paste_5 = GUICtrlCreateButton("Paste", 10, 111, 40, 20)
$paste_6 = GUICtrlCreateButton("Paste", 10, 131, 40, 20)
$paste_7 = GUICtrlCreateButton("Paste", 10, 151, 40, 20)
$paste_8 = GUICtrlCreateButton("Paste", 10, 171, 40, 20)
$paste_9 = GUICtrlCreateButton("Paste", 10, 191, 40, 20)
$paste_slash = GUICtrlCreateButton("Paste", 10, 211, 40, 20)

; Creating GUI labels
GUICtrlCreateLabel("NumPad 0 - Name and Date", 60, 15, 200, 20)
GUICtrlCreateLabel("NumPad 1 - template_1", 60, 35, 200, 20)
GUICtrlCreateLabel("NumPad 2 - template_2", 60, 55, 200, 20)
GUICtrlCreateLabel("NumPad 3 - template_3", 60, 75, 200, 20)
GUICtrlCreateLabel("NumPad 4 - template_4", 60, 95, 200, 20)
GUICtrlCreateLabel("NumPad 5 - template_5", 60, 115, 200, 20)
GUICtrlCreateLabel("NumPad 6 - template_6", 60, 135, 200, 20)
GUICtrlCreateLabel("NumPad 7 - template_7", 60, 155, 200, 20)
GUICtrlCreateLabel("NumPad 8 - template_8", 60, 175, 200, 20)
GUICtrlCreateLabel("NumPad 9 - template_9", 60, 195, 200, 20)
GUICtrlCreateLabel("NumPad / - template_10", 60, 215, 200, 20)

; Activates a function if a certain keystroke is pressed
HotKeySet("{Esc}", "CloseScript")          ; "Esc" - Closes program
HotKeySet("{NUMPAD0}", "NameAndDate")      ; NumPad 0 - Name and Date
HotKeySet("{NUMPAD1}", "Template_1") 	   ; NumPad 1 - template_1
HotKeySet("{NUMPAD2}", "Template_2")       ; NumPad 2 - template_2
HotKeySet("{NUMPAD3}", "Template_3")       ; NumPad 3 - template_3
HotKeySet("{NUMPAD4}", "Template_4")       ; NumPad 4 - template_4
HotKeySet("{NUMPAD5}", "Template_5")       ; NumPad 5 - template_5
HotKeySet("{NUMPAD6}", "Template_6")       ; NumPad 6 - template_6
HotKeySet("{NUMPAD7}", "Template_7")       ; NumPad 7 - template_7
HotKeySet("{NUMPAD8}", "Template_8")       ; NumPad 8 - template_8
HotKeySet("{NUMPAD9}", "Template_9")   	   ; NumPad 9 - template_9
HotKeySet("{NUMPADDIV}", "Template_10")    ; NumPad / - template_10

While 1
	Switch GUIGetMsg()
		Case $paste_0
			NameAndDate()
		Case $paste_1
			Template_1()
		Case $paste_2
			Template_2()
		Case $paste_3
			Template_3()
		Case $paste_4
			Template_4()
		Case $paste_5
			Template_5()
		Case $paste_6
			Template_6()
		Case $paste_7
			Template_7()
		Case $paste_8
			Template_8()
		Case $paste_9
			Template_9()
		Case $paste_slash
			Template_10()
		Case $GUI_EVENT_CLOSE
			CloseScript()
	EndSwitch
Wend

; 0
Func NameAndDate()

	AdobeCheck()
		If @error Then Return

    ClipPut($email)                                     ; Inserts e-mail
    Send("^v")
	Sleep(50)                                       ; Has to sleep due to a bug

    Send("{TAB}") 					; Presses "tab"
	Sleep(50)                                       ; Has to sleep due to a bug

    $year = @YEAR 					; Get the current date components
    $month = StringFormat("%02d", @MON)
    $day = StringFormat("%02d", @MDAY)
    $todaysDate = $year & "-" & $month & "-" & $day 	; Combine them into the desired format YYYY-MM-DD

	ClipPut($todaysDate)                            ; Inserts date
    Send("^v")

EndFunc

; 1
Func Template_1()

	AdobeCheck()
		If @error Then Return
    ClipPut($template_1)     ; Makes the pasting instant
    Send("^v") 		     ; Makes the pasting instant

EndFunc

; 2
Func Template_2()

	AdobeCheck()
		If @error Then Return
    ClipPut($template_2)
    Send("^v")

EndFunc

; 3
Func Template_3()

	AdobeCheck()
		If @error Then Return
    ClipPut($template_3)
    Send("^v")
EndFunc

; 4
Func Template_4()

	AdobeCheck()
		If @error Then Return
    ClipPut($template_4)
    Send("^v")

EndFunc

; 5
Func Template_5()

	AdobeCheck()
		If @error Then Return
    ClipPut($template_5)
    Send("^v")

EndFunc

; 6
Func Template_6()

	AdobeCheck()
		If @error Then Return
    ClipPut($template_6)
    Send("^v")

EndFunc

; 7
Func Template_7()

	AdobeCheck()
		If @error Then Return
    ClipPut($template_7)
    Send("^v")

EndFunc

; 8
Func Template_8()

	AdobeCheck()
		If @error Then Return
    ClipPut($template_8)
    Send("^v")

EndFunc

; 9
Func Template_9()

	AdobeCheck()
		If @error Then Return
    ClipPut($template_9)
    Send("^v")

EndFunc

; /
Func Template_10()

	AdobeCheck()
		If @error Then Return
    ClipPut($template_10)
    Send("^v")

EndFunc

Func AdobeCheck()

	WinActivate("[CLASS:AcrobatSDIWindow]")
	WinWaitActive("[CLASS:AcrobatSDIWindow]", "", 2)
	If Not WinActive("[CLASS:AcrobatSDIWindow]") Then
		MsgBox($MB_SYSTEMMODAL, "Error", "Adobe is not opened")
		Return SetError(1)
	EndIf

EndFunc

Func CloseScript()
    Exit
EndFunc
