#AutoIt3Wrapper_Run_After=del "%scriptfile%_x32.exe"
;#AutoIt3Wrapper_Run_After=ren "%out%" "%scriptfile%_x32.exe"
#AutoIt3Wrapper_Run_After=del "%scriptfile%_stripped.au3"
#AutoIt3Wrapper_Run_Au3Stripper=y
#Au3Stripper_Parameters=/PreExpand /StripOnly /RM ;/RenameMinimum
#AutoIt3Wrapper_Compile_both=n
#AutoIt3Wrapper_UseX64=y
#AutoIt3Wrapper_Res_Description=LW Lang Restorer
#cs
[FileVersion]
#ce
#AutoIt3Wrapper_Res_Fileversion=1.0.9.1
#AutoIt3Wrapper_Res_LegalCopyright=Copyright (C) https://lior.weissbrod.com

#include <Date.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <EditConstants.au3>
#include <TrayConstants.au3>
#include <Array.au3>
#include <WinAPISysWin.au3>
#include <WinAPISys.au3>
#include <SendMessage.au3>
#include <WinAPILocale.au3>
#include <APILocaleConstants.au3>

Opt("TrayAutoPause", 0)

#cs
Copyright (C) https://lior.weissbrod.com

    This program is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.

    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with this program.  If not, see <https://www.gnu.org/licenses/>.

Additional restrictions under GNU GPL version 3 section 7:

In accordance with item 7b), it is required to preserve the reasonable legal notices/author attributions in the material and in the Appropriate Legal Notices displayed by works containing it (including in the footer).
In accordance with item 7c), misrepresentation of the origin of the material must be marked in reasonable ways as different from the original version.
#ce

$programname = "LW Lang Restorer"
;$version = "0.1"
$version = StringRegExpReplace(@Compiled ? StringRegExpReplace(FileGetVersion(@ScriptFullPath), "\.0+$", "") : IniRead(@ScriptFullPath, "FileVersion", "#AutoIt3Wrapper_Res_Fileversion", "0.0.0"), "(\d+\.\d+\.\d+)\.(\d+)", "$1 beta $2")
$thedate = @YEAR
$iniFile = @ScriptDir & "\" & $programname & ".ini"

; Global variables for tray/menu/background logic
Global $g_selectedLangIndex = -1, $g_waitSeconds = 30, $g_hklList[0], $g_running = True, $g_monitorPaused = False
Global $g_hConfigGUI = 0, $g_firstRun = True
Global $hTrayShowMenu = 0
Global $hTrayNextCheck = 0  ; Tray item for next check time

Func Main()
    ShowConfigDialog()
    ; No main loop needed; AdlibRegister is called after OK in the dialog
EndFunc

Func ShowConfigDialog()
    ; Load saved seconds from INI file
    $g_waitSeconds = IniRead($iniFile, "Settings", "WaitSeconds", $g_waitSeconds)
    If $g_waitSeconds < 1 Then $g_waitSeconds = 30

    ; Pause background checks and remove existing tray menu before showing the dialog
    AdlibUnRegister("CheckTrayMenu")
    AdlibUnRegister("CheckInactivity")
    If $hTrayShowMenu Then
        TrayItemDelete($hTrayShowMenu)
        $hTrayShowMenu = 0
    EndIf
    If $hTrayNextCheck Then
        TrayItemDelete($hTrayNextCheck)
        $hTrayNextCheck = 0
    EndIf
    $g_monitorPaused = True
    ; Get active input languages using Windows API
    Local $layouts = _WinAPI_GetKeyboardLayoutList()

    ; Populate global HKL list for restoration
    $g_hklList = $layouts

    Local $preselectIndex = 0
    ; Build display names for each layout, using registry if possible, else just a number
    Local $languages[$layouts[0]]
    Local $nameCount = ObjCreate("Scripting.Dictionary")

    For $i = 1 To $layouts[0]
        Local $langName = _WinAPI_GetLocaleInfo(BitAND($layouts[$i], 0xFFFF), $LOCALE_SLANGUAGE)

        ; Normalize HKL to standard format for registry lookup
        Local $lowWord = BitAND($layouts[$i], 0xFFFF)
        Local $highWord = BitAND(BitShift($layouts[$i], 16), 0xFFFF)

        ; If high word equals low word, it's likely a duplicated value - normalize to 0000
        If $highWord = $lowWord Then
            $highWord = 0x0000
        EndIf

        ; For Hebrew (0x040D), check if this is the standard variant
        If $lowWord = 0x040D Then
            ; If high word is 0xF03D, it should be 0x0002 for Hebrew Standard
            If $highWord = 0xF03D Then
                $highWord = 0x0002
            EndIf
        EndIf

        ; For registry key, format as 8-digit hex: highWord + lowWord
        Local $layoutID = StringFormat("%04X%04X", $highWord, $lowWord)

        Local $layoutText = RegRead("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Keyboard Layouts\" & $layoutID, "Layout Text")

        If $layoutText <> "" Then
            $languages[$i-1] = $langName & " - " & $layoutText
        Else
            If $nameCount.Exists($langName) Then
                $nameCount.Item($langName) += 1
                $languages[$i-1] = $langName & " " & $nameCount.Item($langName)
            Else
                $nameCount.Add($langName, 1)
                $languages[$i-1] = $langName
            EndIf
        EndIf
    Next
    ; Fix pre-selection: match current HKL to $layouts
    Local $currentHKL = _WinAPI_GetKeyboardLayout(WinGetHandle("[ACTIVE]"))
    For $i = 1 To $layouts[0]
        If $layouts[$i] = $currentHKL Then
            $preselectIndex = $i-1
            ExitLoop
        EndIf
    Next
    ; Build the list string
    Local $listStr = ""
    For $i = 0 To UBound($languages) - 1
        $listStr &= $languages[$i] & "|"
    Next
    $listStr = StringTrimRight($listStr, 1)
    ; Show language selection dialog
    $g_hConfigGUI = _ShowLanguageDialogWithHide($languages, True)
    $g_monitorPaused = False
EndFunc

Func _ShowLanguageDialogWithHide($languages, $chooseLive = False)
    Local $layouts = _WinAPI_GetKeyboardLayoutList()
    Local $preselectIndex = 0
    Local $width = 400
    Local $height = 400
    Local $left = (@DesktopWidth - $width) / 2
    Local $top = (@DesktopHeight - $height) / 2
    Local $hGUI = GUICreate($programname & " - Select Default Language", $width, $height, $left, $top, BitOR($GUI_SS_DEFAULT_GUI, $WS_MAXIMIZEBOX))
    $g_hConfigGUI = $hGUI
    Local $hMenu = GUICtrlCreateMenu("Help")
    Local $hAbout = GUICtrlCreateMenuItem("About", $hMenu)
    GUICtrlCreateLabel("Which language to restore after X seconds of inactivity:", 10, 10, 380, 20)
    Local $hList = GUICtrlCreateList("", 10, 35, 380, 200)
    ; Build the list string
    Local $listStr = ""
    For $i = 0 To UBound($languages) - 1
        $listStr &= $languages[$i] & "|"
    Next
    $listStr = StringTrimRight($listStr, 1)
    GUICtrlSetData($hList, $listStr)
    GUICtrlSetData($hList, $languages[$preselectIndex])
    GUICtrlCreateLabel("Seconds of inactivity before restoring:", 10, 245, 200, 20)
    Local $hSpin = GUICtrlCreateInput($g_waitSeconds, 210, 242, 60, 22, BitOr($GUI_SS_DEFAULT_INPUT, $ES_NUMBER))
    Local $hUpDown = GUICtrlCreateUpdown($hSpin)
    GUICtrlSetLimit($hUpDown, 3600, 1)
    Local $hStatusLabel = GUICtrlCreateLabel("Ready to choose", 10, 275, 380, 20)
    GUICtrlSetColor($hStatusLabel, 0x000000)
Local $hLoad = GUICtrlCreateButton("Load", 130, 305, 45, 25)
Local $hSave = GUICtrlCreateButton("Save", 180, 305, 45, 25)
Local $hOK = GUICtrlCreateButton("OK", 110, 340, 70, 30)
Local $hCancel = GUICtrlCreateButton("Cancel", 190, 340, 70, 30)
    GUISetState(@SW_SHOW)
    Local $selectedLanguage = ""
    While 1
        Local $msg = GUIGetMsg()
        Switch $msg
            Case $GUI_EVENT_CLOSE, $hCancel
                If $g_firstRun Then
                    Exit
                Else
                    GUISetState(@SW_HIDE, $hGUI)
                    Return $hGUI
                EndIf
            Case $hLoad
                ; Load settings from INI file
                Local $loadedSeconds = IniRead($iniFile, "Settings", "WaitSeconds", $g_waitSeconds)
                If $loadedSeconds >= 1 Then
                    $g_waitSeconds = $loadedSeconds
                    GUICtrlSetData($hSpin, $g_waitSeconds)
                    GUICtrlSetData($hStatusLabel, "Loaded")
                    GUICtrlSetColor($hStatusLabel, 0x008000) ; Green
                Else
                    GUICtrlSetData($hStatusLabel, "Load Failed")
                    GUICtrlSetColor($hStatusLabel, 0xFF0000) ; Red
                EndIf
            Case $hSave
                ; Save current settings to INI file
                $g_waitSeconds = Number(GUICtrlRead($hSpin))
                If $g_waitSeconds < 1 Then $g_waitSeconds = 1
                If IniWrite($iniFile, "Settings", "WaitSeconds", $g_waitSeconds) Then
                    GUICtrlSetData($hStatusLabel, "Saved")
                    GUICtrlSetColor($hStatusLabel, 0x008000) ; Green
                Else
                    GUICtrlSetData($hStatusLabel, "Save Failed")
                    GUICtrlSetColor($hStatusLabel, 0xFF0000) ; Red
                EndIf
            Case $hOK
                Local $selected = GUICtrlRead($hList)
                If $selected <> "" Then
                    $selectedLanguage = $selected
                EndIf
                $g_waitSeconds = Number(GUICtrlRead($hSpin))
                If $g_waitSeconds < 1 Then $g_waitSeconds = 1
                $g_firstRun = False
                GUISetState(@SW_HIDE, $hGUI)
                $g_hConfigGUI = 0
                $g_monitorPaused = False
                AdlibRegister("CheckTrayMenu", 100)
                AdlibRegister("CheckInactivity", $g_waitSeconds * 1000)
                TraySetIcon(@SystemDir & "\shell32.dll", 137) ; Keyboard icon
                TraySetToolTip($programname & " - Click to configure")
                $hTrayShowMenu = TrayCreateItem("Show Menu")
                Local $nextCheckTime = _DateAdd('s', $g_waitSeconds, _NowCalc())
                Local $aDatePart, $aTimePart
                _DateTimeSplit($nextCheckTime, $aDatePart, $aTimePart)
                Local $timeStr = StringFormat("%02d:%02d:%02d", $aTimePart[1], $aTimePart[2], $aTimePart[3])
                $hTrayNextCheck = TrayCreateItem("Next check at " & $timeStr & " for " & $selectedLanguage)
                TrayItemSetState($hTrayNextCheck, $TRAY_DISABLE)
                ; Find the selected index
                For $i = 0 To UBound($languages) - 1
                    If $languages[$i] = $selectedLanguage Then
                        $g_selectedLangIndex = $i
                        ExitLoop
                    EndIf
                Next
            Case $hAbout
                about($hGUI)
        EndSwitch
    WEnd
    Return $hGUI
EndFunc

Func GetActiveInputLanguages()
    Local $layouts[0]
    Local $hUser32 = DllOpen("user32.dll")
    If $hUser32 = -1 Then Return $layouts
    ; Get the number of active layouts
    Local $count = DllCall($hUser32, "int", "GetKeyboardLayoutList", "int", 0, "ptr", 0)
    If @error Or $count[0] = 0 Then
        DllClose($hUser32)
        Return $layouts
    EndIf
    ; Allocate memory for the layout list
    Local $layoutList = DllStructCreate("ptr[" & $count[0] & "]")
    Local $pLayoutList = DllStructGetPtr($layoutList)
    ; Get the actual layout list
    Local $result = DllCall($hUser32, "int", "GetKeyboardLayoutList", "int", $count[0], "ptr", $pLayoutList)
    If @error Or $result[0] = 0 Then
        DllClose($hUser32)
        Return $layouts
    EndIf
    ; For each layout, get the friendly name
    For $i = 1 To $result[0]
        Local $layout = DllStructGetData($layoutList, 1, $i)
        Local $langID = BitAND($layout, 0xFFFF)
        Local $desc = GetLocaleName($langID)
        If $desc <> "" Then
            ReDim $layouts[UBound($layouts) + 1]
            $layouts[UBound($layouts) - 1] = $desc
        EndIf
    Next
    DllClose($hUser32)
    Return $layouts
EndFunc

Func GetLocaleName($langID)
    ; Get the language name (e.g., "English (United States)")
    Local $len = DllCall("kernel32.dll", "int", "GetLocaleInfoW", "dword", $langID, "dword", 0x00000002, "wstr", "", "int", 0)
    If @error Or $len[0] = 0 Then Return ""
    Local $buf = DllStructCreate("wchar[" & $len[0] & "]")
    Local $ret = DllCall("kernel32.dll", "int", "GetLocaleInfoW", "dword", $langID, "dword", 0x00000002, "ptr", DllStructGetPtr($buf), "int", $len[0])
    If @error Or $ret[0] = 0 Then Return ""
    Return DllStructGetData($buf, 1)
EndFunc

Func about($MainWindow = "")
  if $MainWindow == "" then
	GUICreate("About " & $programname, 435, 410, -1, -1, -1)
  Else
	GUICreate("About " & $programname, 435, 410, -1, -1, -1, $WS_EX_MDICHILD, $MainWindow)
  EndIf
  $localleft=10
  $localtop=10
  $message=$programname & " - Version " & $version & @crlf & _
  @crlf & _
  "A portable way to restore one's default input keyboard language."
  GUICtrlCreateLabel($message, $localleft, $localtop)
  $message = chr(169) & $thedate & " LWC"
  GUICtrlCreateLabel($message, $localleft, ControlGetPos(GUICtrlGetHandle(-1), "", 0)[3]+18)
  local $aLabel = GUICtrlCreateLabel("https://lior.weissbrod.com", ControlGetPos(GUICtrlGetHandle(-1), "", 0)[2]+10, _
  ControlGetPos(GUICtrlGetHandle(-1), "", 0)[1]+ControlGetPos(GUICtrlGetHandle(-1), "", 0)[3]-$localtop-12)
  GUICtrlSetFont(-1,-1,-1,4)
  GUICtrlSetColor(-1,0x0000cc)
  GUICtrlSetCursor(-1,0)
  $message="    This program is free software: you can redistribute it and/or modify" & _
@crlf & "    it under the terms of the GNU General Public License as published by" & _
@crlf & "    the Free Software Foundation, either version 3 of the License, or" & _
@crlf & "    (at your option) any later version." & _
@crlf & _
@crlf & "    This program is distributed in the hope that it will be useful," & _
@crlf & "    but WITHOUT ANY WARRANTY; without even the implied warranty of" & _
@crlf & "    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the" & _
@crlf & "    GNU General Public License for more details." & _
@crlf & _
@crlf & "    You should have received a copy of the GNU General Public License" & _
@crlf & "    along with this program.  If not, see <https://www.gnu.org/licenses/>." & _
@crlf & @crlf & _
"Additional restrictions under GNU GPL version 3 section 7:" & _
@crlf & @crlf & _
"* In accordance with item 7b), it is required to preserve the reasonable legal notices/author attributions in the material and in the Appropriate Legal Notices displayed by works containing it (including in the footer)." & _
@crlf & @crlf & _
"* In accordance with item 7c), misrepresentation of the origin of the material must be marked in reasonable ways as different from the original version."
  GUICtrlCreateLabel($message, $localleft, $localtop+85, 420, 280)
  $okay=GUICtrlCreateButton("OK", $localleft+160, $localtop+365, 100)

  GUISetState(@SW_SHOW)
  While 1
	$msg=guigetmsg()
	switch $msg
		case $GUI_EVENT_CLOSE, $okay
			guidelete()
			ExitLoop
		case $aLabel
			ShellExecute(GUICtrlRead($msg))
	endswitch
  WEnd
EndFunc

Func _IsUserInactive($seconds)
    ; Use correct struct for LASTINPUTINFO: dword;dword (cbSize, dwTime)
    Local $struct = DllStructCreate("dword;dword")
    DllStructSetData($struct, 1, DllStructGetSize($struct))
    Local $ret = DllCall("user32.dll", "bool", "GetLastInputInfo", "ptr", DllStructGetPtr($struct))
    If @error Or Not $ret[0] Then Return False
    Local $lastInputTick = DllStructGetData($struct, 2)
    Local $currentTick = DllCall("kernel32.dll", "dword", "GetTickCount")[0]
    Local $idleTime = ($currentTick - $lastInputTick) / 1000
    Return $idleTime >= $seconds
EndFunc

Func _SwitchToLanguage($targetIndex)
    ; Get the current language index
    Local $currentIndex = _GetCurrentLanguageIndex()
    If $currentIndex = -1 Then Return
    ; Use $g_hklList[0] as the count, and start from index 1
    Local $count = $g_hklList[0]
    Local $steps = Mod($targetIndex - $currentIndex + $count, $count)
    For $i = 1 To $steps
        Send("!+") ; Alt+Shift
        Sleep(100)
    Next
EndFunc

Func _GetCurrentLanguageIndex()
    Local $hWnd = WinGetHandle("[ACTIVE]")
    Local $hkl = _WinAPI_GetKeyboardLayout($hWnd)
    For $i = 1 To $g_hklList[0]
        If $g_hklList[$i] = $hkl Then Return $i
    Next
    Return -1
EndFunc

Func _RestoreInputLanguage($index)
    If $index < 0 Or $index >= $g_hklList[0] Then Return
    Local $hWnd = WinGetHandle("[ACTIVE]")
    Local $hkl = $g_hklList[$index + 1]  ; Array format: [0]=count, [1]=first HKL, [2]=second HKL, etc.
    _WinAPI_SetKeyboardLayout($hWnd, $hkl)
EndFunc

Func CheckTrayMenu()
    Local $trayMsg = TrayGetMsg()
    Switch $trayMsg
        Case $hTrayShowMenu
            ShowConfigDialog()
    EndSwitch
EndFunc

Func CheckInactivity()
    If $g_selectedLangIndex > -1 And Not $g_monitorPaused Then
        _RestoreInputLanguage($g_selectedLangIndex)
    EndIf
EndFunc

Main()