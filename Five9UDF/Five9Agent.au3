#include <String.au3>
#include "./Five9Constants.au3"

; #FUNCTION# ====================================================================================================================
; Name ..........: _Five9AgentGetPresence
; Description ...: This UDF allows a user to get a copy of the Agent object's Presence
; Syntax ........: _Five9AgentGetPresence($sAPI, $sUser, $sPassword)
; Parameters ....: $sAPI                - Five9 API URL
;                  $sUser               - The ID of the user
;                  $sPassword           - The Password of the user
; Return values .: Success - Returns Response from Five9
;                  Failure - Sets @error:
;                  |1   - WinHTTP Open Error, see Return Value for @error and @extended for @extended
;                  |2   - WinHTTP Send Error, see Return Value for @error and @extended for @extended
;                  |400 - Bad Request (Five9)
;                  |401 - Unauthorized, See Returned Data (Five9)
;                  |404 - Not Found (Five9)
;                  |500 - Internal Server Error (Five9)
;                  |503 - Service Unavailable (Five9)
; Author ........: Robert Maehl
; Modified ......: 3/10/2021
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _Five9AgentGetPresence($sAPI, $sUser, $sPassword)

	Local $oHTTP = ObjCreate("WinHttp.WinHttpRequest.5.1")

	$oHTTP.Open("GET", $sAPI & "agents/" & $sUser & "/presence", False)
	If @error Then Return SetError(1, @extended, @error)

	$oHTTP.SetCredentials($sUser, $sPassword, 0)

	$oHTTP.Send()
	If @error Then Return SetError(2, 0, 0)

	If $oHTTP.Status <> $FIVE9_STATUS_SUCCESS Then Return SetError($oHTTP.Status, 0, $oHTTP.ResponseText)

	$sState = $oHTTP.ResponseText

	Return SetError(0, 0, $sState)

EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _Five9AgentGetState
; Description ...: This UDF allows a user to get a copy of the Agent object's State
; Syntax ........: _Five9AgentGetState($sAPI, $sUser, $sPassword)
; Parameters ....: $sAPI                - Five9 API URL
;                  $sUser               - The ID of the user
;                  $sPassword           - The Password of the user
; Return values .: Success - Returns Response from Five9. An interger if Not Ready, an Array if Ready
;                  Failure - Sets @error:
;                  |1   - WinHTTP Open Error, see Return Value for @error and @extended for @extended
;                  |2   - WinHTTP Send Error, see Return Value for @error and @extended for @extended
;                  |400 - Bad Request (Five9)
;                  |401 - Unauthorized, See Returned Data (Five9)
;                  |404 - Not Found (Five9)
;                  |500 - Internal Server Error (Five9)
;                  |503 - Service Unavailable (Five9)
; Author ........: Robert Maehl
; Modified ......: 3/10/2021
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _Five9AgentGetState($sAPI, $sUser, $sPassword)

	Local $oHTTP = ObjCreate("WinHttp.WinHttpRequest.5.1")

	$oHTTP.Open("GET", $sAPI & "agents/" & $sUser & "/presence", False)
	If @error Then Return SetError(1, @extended, @error)

	$oHTTP.SetCredentials($sUser, $sPassword, 0)

	$oHTTP.Send()
	If @error Then Return SetError(2, 0, 0)

	If $oHTTP.Status <> $FIVE9_STATUS_SUCCESS Then Return SetError($oHTTP.Status, 0, $oHTTP.ResponseText)

	$sState = $oHTTP.ResponseText

	$sState = StringReplace($sState, @CR, "") ; Trim Carriage Returns
	$sState = StringReplace($sState, @LF, "") ; Trim Line Feeds
	$sState = StringStripWS($sState, $STR_STRIPSPACES)
	If _StringBetween($sState, '"notReadyReasonCodeId": "', '"')[0] Then
		$sState = _StringBetween($sState, '"notReadyReasonCodeId": "', '"')[0]
	Else
		$sState = StringSplit(_StringBetween($sState, 'readyChannels": [', '],"')[0], ",")
	EndIf

	Return SetError(0, 0, $sState)

EndFunc


; #FUNCTION# ====================================================================================================================
; Name ..........: _Five9GetAgent
; Description ...: This UDF allows a user to get a copy of the Agent object.
; Syntax ........: _Five9GetAgent($sAPI, $sUser, $sPassword)
; Parameters ....: $sAPI                - Five9 API URL
;                  $sUser               - The ID of the user
;                  $sPassword           - The Password of the user
; Return values .: Success - Returns Response from Five9
;                  Failure - Sets @error:
;                  |1   - WinHTTP Open Error, see Return Value for @error and @extended for @extended
;                  |2   - WinHTTP Send Error, see Return Value for @error and @extended for @extended
;                  |400 - Bad Request (Five9)
;                  |401 - Unauthorized, See Returned Data (Five9)
;                  |404 - Not Found (Five9)
;                  |500 - Internal Server Error (Five9)
;                  |503 - Service Unavailable (Five9)
; Author ........: Robert Maehl
; Modified ......: 3/10/2021
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _Five9GetAgent($sAPI, $sUser, $sPassword)

	Local $oHTTP = ObjCreate("WinHttp.WinHttpRequest.5.1")

	$oHTTP.Open("GET", $sAPI & "agents/" & $sUser, False)
	If @error Then Return SetError(1, @extended, @error)

	$oHTTP.SetCredentials($sUser, $sPassword, 0)

	$oHTTP.Send()
	If @error Then Return SetError(2, 0, 0)

	If $oHTTP.Status <> $FIVE9_STATUS_SUCCESS Then Return SetError($oHTTP.Status, 0, $oHTTP.ResponseText)

	Return SetError(0, 0, $oHTTP.ResponseText)

EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _Five9GetAgents
; Description ...: This API allows an administrator to get a list of agents.
; Syntax ........: _Five9GetAgents($sAPI, $sUser, $sPassword)
; Parameters ....: $sAPI                - Five9 API URL
;                  $sUser               - The ID of the user
;                  $sPassword           - The Password of the user
; Return values .: Success - Returns Response from Five9
;                  Failure - Sets @error:
;                  |1   - WinHTTP Open Error, see Return Value for @error and @extended for @extended
;                  |2   - WinHTTP Send Error, see Return Value for @error and @extended for @extended
;                  |400 - Bad Request (Five9)
;                  |401 - Unauthorized, See Returned Data (Five9)
;                  |404 - Not Found (Five9)
;                  |500 - Internal Server Error (Five9)
;                  |503 - Service Unavailable (Five9)
; Author ........: Robert Maehl
; Modified ......: 3/10/2021
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _Five9GetAgents($sAPI, $sUser, $sPassword)

	Local $oHTTP = ObjCreate("WinHttp.WinHttpRequest.5.1")

	$oHTTP.Open("GET", $sAPI & "agents/" , False)
	If @error Then Return SetError(1, @extended, @error)

	$oHTTP.SetCredentials($sUser, $sPassword, 0)

	$oHTTP.Send()
	If @error Then Return SetError(2, 0, 0)

	If $oHTTP.Status <> $FIVE9_STATUS_SUCCESS Then Return SetError($oHTTP.Status, 0, $oHTTP.ResponseText)

	Return SetError(0, 0, $oHTTP.ResponseText)

EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _Five9AgentSetState
; Description ...: This UDF allows a user to change the state of an agent on the Five9 server. Agents can change their own states.
;                  Additionally, when changing state to no Ready Channels, this UDF allows a user to change the agent state
;                  in the Five9 server and pass along the code value of a corresponding reason code.
; Syntax ........: _Five9AgentSetState($sAPI, $sChannels, $iReason, $sUser, $sPassword)
; Parameters ....: $sAPI                - Five9 API URL
;                  $sChannels           - The new Ready Channels the user wants to be in (VOICE_MAIL, CALL)
;                  $iReason             - The database ID for the reason code
;                  $sUser               - The ID of the user
;                  $sPassword           - The Password of the user
; Return values .: Success - Returns Response from Five9
;                  Failure - Sets @error:
;                  |1   - WinHTTP Open Error, see Return Value for @error and @extended for @extended
;                  |2   - WinHTTP Send Error, see Return Value for @error and @extended for @extended
;                  |400 - Bad Request (Five9)
;                  |401 - Unauthorized, See Returned Data (Five9)
;                  |404 - Not Found (Five9)
;                  |500 - Internal Server Error (Five9)
;                  |503 - Service Unavailable (Five9)
; Author ........: Robert Maehl
; Modified ......: 3/10/2021
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _Five9AgentSetState($sAPI, $sChannels, $iReason, $sUser, $sPassword)

	Local $oHTTP = ObjCreate("WinHttp.WinHttpRequest.5.1")

	$oHTTP.Open("PUT", $sAPI & "agents/" & $sUser & "/presence", False)
	If @error Then Return SetError(1, @extended, @error)

	$oHTTP.SetCredentials($sUser, $sPassword, 0)

	$oHTTP.SetRequestHeader("Content-Type", "application/json")

	$sChannels = StringStripWS($sChannels, $STR_STRIPALL)
	$sChannels = StringReplace($sChannels, ',', '","')

	If $iReason = "" Then
		$oHTTP.Send('{"currentState": {"readyChannels": ["' & $sChannels & '],"notReadyReasonCodeId": "0"},"pendingState": null,"currentStateTime": 0,"pendingStateDelayTime": 0,"gracefulModeOn": false}')
	Else
		$oHTTP.Send('{"currentState": {"readyChannels": [],"notReadyReasonCodeId": "' & $iReason & '"},"pendingState": null,"currentStateTime": 0,"pendingStateDelayTime": 0,"gracefulModeOn": false}')
	EndIf
	If @error Then Return SetError(2, 0, 0)

	If $oHTTP.Status = $FIVE9_STATUS_SUCCESS Then
		;;;
	ElseIf $oHTTP.Status = $FIVE9_STATUS_ACCEPTED Then
		;;;
	Else
		Return SetError($oHTTP.Status, 0, $oHTTP.ResponseText)
	EndIf

	Return SetError(0, 0, True)

EndFunc