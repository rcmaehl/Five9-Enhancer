#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=Five9.ico
#AutoIt3Wrapper_Res_Comment=Five9 Enhancer
#AutoIt3Wrapper_Res_Description=Five9 Enhancer add-on
#AutoIt3Wrapper_Res_Fileversion=1.0.0.0
#AutoIt3Wrapper_Res_LegalCopyright=Created by Robert Maehl
#AutoIt3Wrapper_Res_Language=1033
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

#include <File.au3>
#include <Array.au3>
#include <WinAPIEx.au3>
#include <TrayConstants.au3>
#include <GUIConstantsEx.au3>
#include <ButtonConstants.au3>
#include <MsgBoxConstants.au3>
#include <StringConstants.au3>
#include <WindowsConstants.au3>

#include ".\Includes\Five9Agent.au3"
#include ".\Includes\_ExtendedFunctions.au3"

Opt("TrayMenuMode", 3) ; Disable Default Tray menu

Global $sVer = "1.0.0.0"
Global $bDebug = True

Global Const $CTRL_ALL = 0
Global Const $CTRL_CREATED = 1
Global Const $ABM_GETTASKBARPOS = 0x5

Global Enum $hRemindMe, $hCustomLess, $honStartup

$oMyError = ObjEvent("AutoIt.Error","_ThrowError") ; Initialize a COM error handler

Main()

Func Main()

	Local $TrayOpts = TrayCreateItem("Settings")
	TrayCreateItem("")
	Local $TrayExit = TrayCreateItem("Exit"    )

	Local $hGUI = GUICreate("Settings", 280, 100, @DesktopWidth - 300, @DesktopHeight - 180, BitXOR($GUI_SS_DEFAULT_GUI, $WS_MINIMIZEBOX), $WS_EX_TOOLWINDOW + $WS_EX_TOPMOST)

	Local $aControls[3]

	$aControls[$hRemindMe  ] = GUICtrlCreateCheckbox("Show Reminders for Non-Ready Status", 10, 00, 260, 20, $BS_RIGHTBUTTON)
	$aControls[$hCustomLess] = GUICtrlCreateCheckbox("Enable Custom Reminders"            , 10, 20, 260, 20, $BS_RIGHTBUTTON)
	$aControls[$honStartup]  = GUICtrlCreateCheckbox("Start with Windows"                 , 10, 40, 260, 20, $BS_RIGHTBUTTON)

	GUICtrlSetTip($aControls[$hRemindMe], "Display 15 minute, 30 minute, & 1 hour reminders for Not Ready")
	GUICtrlSetTip($aControls[$hCustomLess], "Use 1 minute reminders for SD Tasks")

	$aSettings = _LoadSettings()
;	_ArrayDisplay($aSettings)
	GUICtrlSetState($aControls[$hRemindMe  ], $aSettings[$hRemindMe  ])
	GUICtrlSetState($aControls[$hCustomLess], $aSettings[$hCustomLess])
	GUICtrlSetState($aControls[$honStartup] , $aSettings[$honStartup] )

	Global $aAPI = StringSplit($aSettings[3], ",")
	If @error And Not $bDebug Then
		$aAPI = MsgBox($MB_OK + $MB_ICONWARNING + $MB_TOPMOST, "ALERT", "Five9 API  was not specified" & @CRLF & "Five9 Enhancer will now exit", 30)
		Exit 1
	Else
		$aAPI[0] = 1
	EndIf

;	_ArrayDisplay($aAPI)

	Local $iPoll = Number($aSettings[4])
	Local $iUser = $aSettings[5]
	Local $iPass = $aSettings[6]

	Local $sStatus = Null

	Local $iNRC = 0
	Local $aCtrls = Null
	Local $bCLock = False
	Local $bLLock = False
	Local $bTimer = False
	Local $hTimer = TimerInit()
	Local $hMsgBox = Null
	Local $hTaskBar = Null
	Local $hNRTimer = TimerInit()
	Local $hLastPoll = TimerInit()
	Local $bReserved = False

	Select

		Case $bDebug
			;;;

		Case Not $aAPI And Not IsArray($aAPI)
			ContinueCase

		Case Not $iUser
			ContinueCase

		Case Not $iPass
			MsgBox($MB_OK + $MB_ICONWARNING + $MB_TOPMOST, "ALERT", "Five9 API specified but API or Credentials Invalid" & @CRLF & "Five9 Enhancer will now exit", 30)
			Exit 1

	EndSelect

	While 1

		$hGMsg = GUIGetMsg()
		$hTMsg = TrayGetMsg()

		Select

			Case $hTMsg = $TrayOpts Or $hTMsg = $TRAY_EVENT_PRIMARYDOUBLE
				GUISetState(@SW_SHOW, $hGUI)
				$hTaskBar = _GetTaskBarPos()
				WinMove($hGUI, "", @DesktopWidth - 300, $hTaskBar[2] - 120)

			Case $hTMsg = $TrayExit
				Exit

			Case Else

				Switch $hGMsg

					Case $GUI_EVENT_CLOSE
						GUISetState(@SW_HIDE, $hGUI)


					Case $aControls[$hRemindMe] To $aControls[$honStartup]
						If _IsChecked($aControls[$hOnStartup]) And Not FileExists(@StartupDir & "\Five9 Enhancer.lnk") Then
							FileCreateShortcut(@AutoItExe, @StartupDir & "\Five9 Enhancer.lnk")
						ElseIf Not _IsChecked($aControls[$honStartup]) And FileExists(@StartupDir & "\Five9 Enhancer.lnk") Then
							FileDelete(@StartupDir & "\Five9 Enhancer.lnk")
						EndIf
						_SaveSettings($aControls)

				EndSwitch

		EndSelect


		If Not _IsChecked($aSettings[$hRemindMe]) Then
			TraySetToolTip("Not Running - Disabled in Settings")
			ContinueLoop
		EndIf

		TraySetToolTip("Running...")

			If TimerDiff($hLastPoll) >= $iPoll Then ; Rate Limiting is important
				$iPoll = 500
				$hLastPoll = TimerInit()
				$sStatus = _Five9AgentGetState($aAPI[$aAPI[0]], $iUser, $iPass)
;				If StringInStr($sStatus, "SERVER_OUT_OF_SERVICE") Then $iPoll = 30000
				If @error Then
					ConsoleWrite($sStatus & @CRLF & _
					 @error & @CRLF & _
					 @extended & @CRLF)
				EndIf
			EndIf

		$bCLock = _GetDesktopLock()

#cs
		Switch $sStatus

			Case "Not Ready", "NOT_READY"
				If Not $bTimer Then
					$bTimer = True
					$hNRTimer = TimerInit()
				ElseIf TimerDiff($hNRTimer) >= 900000 Then
					$iNRC += 1
					If $iNRC = 3 Or $iNRC = 6 Or $iNRC = 7 Then
						;;;
					Else
						WinMinimizeAll()
						$hMsgBox = MsgBox($MB_YESNO + $MB_ICONWARNING + $MB_TOPMOST, "Reminder", "You've been in Not Ready Status for over " & $iNRC * 15 & " minutes. Would you like to go back to Ready Status?", 15)
						If $hMsgBox = $IDYES Then
							_Five9AgentSetState($aAPI[$aAPI[0]], "CALL", "", $iUser, $iPass)
						EndIf
						WinMinimizeAllUndo()
						$hMsgBox = Null
					EndIf

					$hNRTimer = TimerInit()
				EndIf

				If $bCLock = True Then ; If Desktop is locked
					$bLLock = True
					$CLock = Null
				ElseIf $bCLock = False And $bLLock = True Then ; If Desktop is unlocked but WAS locked
					$bLLock = False
					$bCLock = True
					$hMsgBox = MsgBox($MB_YESNO + $MB_ICONWARNING + $MB_TOPMOST, "Reminder", "You've just logged back in while Not Ready. You've been Not Ready for " & ($iNRC * 15) + Floor(TimerDiff($hNRTimer) / 60000) & " minutes. Would you like to go back to Ready Status?", 30)
					If $hMsgBox = $IDYES Then
						_Five9AgentSetState($aAPI[$aAPI[0]], "CALL", "", $iUser, $iPass)
					EndIf
					$hMsgBox = Null
				Else
					$bLLock = $bCLock
					$CLock = Null
				EndIf

			Case "Ready", "READY"
				If $bTimer And $bReserved Then
;								FileWrite(@ScriptDir & "\IdleTime.log", "RONA/Hangup, " & Round(TimerDiff($hNRTimer) / 1000, 2) & "s" & @CRLF)
;								MsgBox($MB_OK + $MB_ICONWARNING + $MB_TOPMOST, "ALERT", "Ready/Reserved Issue Detected.", 30)
					$hBTimer = TimerInit()
					$bReserved = False
				EndIf
				$iNRC = 0
				$bTimer = False
				$hNRTimer = TimerInit()

			Case "Reserved"
				If Not $bReserved Then
;								FileWrite(@ScriptDir & "\IdleTime.log", Round(TimerDiff($hBTimer) / 1000, 2) & "s, ")
				EndIf
				$iNRC = 0
				$bTimer = True
				If Not $bReserved Then
					$hNRTimer = TimerInit()
					$bReserved = True
				EndIf

			Case "Talking"
				If $bReserved Then
;								FileWrite(@ScriptDir & "\IdleTime.log", "Answered, " & Round(TimerDiff($hNRTimer) / 1000, 2) & "s" & @CRLF)
				EndIf
				$iNRC = 0
				$bTimer = False
				$hBTimer = TimerInit()
				$hNRTimer = TimerInit()
				$bReserved = False

			Case "WrapUp", "WORK_READY"
				If Not $bTimer Then
					$bTimer = True
					$hTimer = TimerInit()
				ElseIf _IsChecked($aSettings[$hCustomLess]) And TimerDiff($hNRTimer) >= 60000 Then
					$iNRC += 1
					$hMsgBox = MsgBox($MB_YESNO + $MB_ICONWARNING + $MB_TOPMOST, "Reminder", "You've been in WrapUp Status for over " & 1 * $iNRC & " minute(s). Would you like to go back to Ready Status?", 15)
					If $hMsgBox = $IDYES Then
						_Five9AgentSetState($aAPI[$aAPI[0]], "CALL", "", $iUser, $iPass)
					EndIf
					$hMsgBox = Null
					$hNRTimer = TimerInit()
				ElseIf TimerDiff($hNRTimer) >= 120000 Then
					$iNRC += 1
					$hMsgBox = MsgBox($MB_YESNO + $MB_ICONWARNING + $MB_TOPMOST, "Reminder", "You've been in WrapUp Status for over " & 2 * $iNRC & "minutes. Would you like to go back to Ready Status?", 15)
					If $hMsgBox = $IDYES Then
						_Five9AgentSetState($aAPI[$aAPI[0]], "CALL", "", $iUser, $iPass)
					EndIf
					$hMsgBox = Null
					$hNRTimer = TimerInit()
				EndIf

			Case "LOGOUT"

			Case "HOLD"

			Case Else
				ConsoleWrite($sStatus & @CRLF)

		EndSwitch
#ce

	WEnd

EndFunc   ;==>Main

; Additional Functions Called

Func _IsChecked($idControlID)
	Return Number(BitAND(GUICtrlRead($idControlID), $GUI_CHECKED) = $GUI_CHECKED)
EndFunc   ;==>_IsChecked

Func _GetDesktopLock()
    Local $fIsLocked = False
    Local Const $hDesktop = _WinAPI_OpenDesktop('Default', $DESKTOP_SWITCHDESKTOP)
    If @error = 0 Then
        $fIsLocked = Not _WinAPI_SwitchDesktop($hDesktop)
        _WinAPI_CloseDesktop($hDesktop)
    EndIf
    Return $fIsLocked
EndFunc   ;==>_GetDesktopLock

Func _GetTaskBarPos()
    $h_taskbar = WinGetHandle("","Start")
    $AppBarData = DllStructCreate("dword;int;uint;uint;int;int;int;int;int")
;~ DWORD cbSize;
;~   HWND hWnd;
;~   UINT uCallbackMessage;
;~   UINT uEdge;
;~   RECT rc;
;~   LPARAM lParam;
    DllStructSetData($AppBarData,1,DllStructGetSize($AppBarData))
    DllStructSetData($AppBarData,2,$h_taskbar)
    $lResult = DllCall("shell32.dll","int","SHAppBarMessage","int",$ABM_GETTASKBARPOS,"ptr",DllStructGetPtr($AppBarData))
    If Not @error Then
        If $lResult[0] Then
            Return StringSplit(DllStructGetData($AppBarData,5) & "|" & _
                DllStructGetData($AppBarData,6) & "|"   & DllStructGetData($AppBarData,7) & "|" & _
                DllStructGetData($AppBarData,8),"|")
        EndIf
    EndIf
    SetError(1)
    Return 0
EndFunc   ;==>_GetTaskBarPos

Func _LoadSettings()
	_UpdateSettings(IniRead(".\Five9E.ini", "#Meta", "FileVer", "0.0.0.0"))
	Local $aSettings[7]
	$aSettings[$hRemindMe  ] = _IniRead(".\Five9E.ini", "Five9E", "Monitor Status"       , "1|0", $GUI_UNCHECKED)
	$aSettings[$hCustomLess] = _IniRead(".\Five9E.ini", "Five9E", "Use Custom Reminders" , "1|0", $GUI_UNCHECKED)
	$aSettings[$honStartup]  = _IniRead(".\Five9E.ini", "Five9E", "Start with Windows"   , "1|0", $GUI_UNCHECKED)
	$aSettings[3] = _IniRead(".\Five9E.ini", "Five9 API", "API URLs"     , ""   , False)
	$aSettings[4] = _IniRead(".\Five9E.ini", "Five9 API", "API Poll Rate", ""   , 500  )
	$aSettings[5] = _IniRead(".\Five9E.ini", "Five9 API", "User ID"      , ""   , False)
	$aSettings[6] = _IniRead(".\Five9E.ini", "Five9 API", "Password"     , ""   , False)
	Return $aSettings
EndFunc   ;==>_LoadSettings

Func _SaveSettings($aSettings)
 	IniWrite(".\Five9E.ini", "#Meta"  , "FileVer"            , $sVer)
	IniWrite(".\Five9E.ini", "Five9E" , "Monitor Status"      , _IsChecked($aSettings[$hRemindMe]  ))
	IniWrite(".\Five9E.ini", "Five9E" , "Use Custom Reminders", _IsChecked($aSettings[$hCustomLess]))
	IniWrite(".\Five9E.ini", "Five9E" , "Start with Windows"  , _IsChecked($aSettings[$hOnStartup] ))
EndFunc   ;==>_SaveSettings

Func _ThrowError()
	$aAPI[0] += 1
	If $aAPI > UBound($aAPI) Then
		Msgbox($MB_ICONERROR,"COM Error", "Five9 Enhancer is unable to communicate with the Five9 API and will now exit." & @CRLF  & @CRLF & _
				"Description: " & @TAB & $oMyError.description  & @CRLF & _
				"Full Description:"   & @TAB & $oMyError.windescription & @CRLF & _
				"Error Number: "       & @TAB & hex($oMyError.number,8)  & @CRLF & _
				"DLL Error ID: "   & @TAB & $oMyError.lastdllerror   & @CRLF & _
				"Error Occured: "   & @TAB & $oMyError.scriptline   & @CRLF & _
				"COM Called: "       & @TAB & $oMyError.source       & @CRLF & _
				"Help File: "       & @TAB & $oMyError.helpfile     & @CRLF & _
				"Help Context: " & @TAB & $oMyError.helpcontext _
				, 10)
	Else
		MsgBox($MB_ICONWARNING, "API Error", "Five9 Enhancer is unable to communite with the Five9 API and will attempt another." & @CRLF & @CRLF & _
				"Current API: "	& @TAB & $aAPI[$aAPI[0]] & @CRLF &_
				"Previous API: " & @TAB & $aAPI[$aAPI[0] - 1] _
				, 10)
	EndIf
	Exit 1
EndFunc

Func _UpdateSettings($sVersion)

	Local $aSettings[0]
	IniWrite(".\Five9E.ini", "#META", "FileVer", $sVer)

	Switch $sVersion

		Case "0.0.0.0"

			FileInstall(".\Five9E.ini", ".\Five9E.ini")

		Case Else
			;;;

	EndSwitch

EndFunc