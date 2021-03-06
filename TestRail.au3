#include-once
#Include <Array.au3>
#include <GuiEdit.au3>
#include <String.au3>
#include "cURL.au3"
#include "Json.au3"
#Region Header
#cs
	Title:   		Janison Insights Automation UDF Library for AutoIt3
	Filename:  		JanisonInsights.au3
	Description: 	A collection of functions for creating, attaching to, reading from and manipulating Janison Insights
	Author:   		seangriffin
	Version:  		V0.1
	Last Update: 	25/02/18
	Requirements: 	AutoIt3 3.2 or higher,
					Janison Insights Release x.xx,
					cURL xxx
	Changelog:		---------24/12/08---------- v0.1
					Initial release.
#ce
#EndRegion Header
#Region Global Variables and Constants
Global Const $sap_vkey[100] = [ "Enter", "F1", "F2", "F3", "F4", "F5", "F6", "F7", "F8", "F9", "F10", _
								"F11", _ ; NOTE - "F11" is the same as "CTRL+S"
								"F12", _ ; NOTE - "F12" is the same as "Esc"
								"Shift+F1", "Shift+F2", "Shift+F3", "Shift+F4", "Shift+F5", "Shift+F6", "Shift+F7", "Shift+F8", "Shift+F9", _
								"Shift+Ctrl+0", "Shift+F11", "Shift+F12", _
								"Ctrl+F1", "Ctrl+F2", "Ctrl+F3", "Ctrl+F4", "Ctrl+F5", "Ctrl+F6", "Ctrl+F7", "Ctrl+F8", "Ctrl+F9", "Ctrl+F10", _
								"Ctrl+F11", "Ctrl+F12", _
								"Ctrl+Shift+F1", "Ctrl+Shift+F2", "Ctrl+Shift+F3", "Ctrl+Shift+F4", "Ctrl+Shift+F5", _
								"Ctrl+Shift+F6", "Ctrl+Shift+F7", "Ctrl+Shift+F8", "Ctrl+Shift+F9", "Ctrl+Shift+F10", "Ctrl+Shift+F11", _
								"Ctrl+Shift+F12", _
								"", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", _
								"Ctrl+E", "Ctrl+F", "Ctrl+A", "Ctrl+D", "Ctrl+N", "Ctrl+O", "Shift+D", "Ctrl+I", "Shift+I", "Alt+B", _
								"Ctrl+Page up", "Page up", "Page down", "Ctrl+Page down", "Ctrl+G", "Ctrl+R", "Ctrl+P", _
								"", "", "", "", "", "", "", "Shift+F10", "", "", "", "", "" ]
Global $testrail_domain = ""
Global $testrail_username = ""
Global $testrail_password = ""
Global $testrail_json = ""
Global $testrail_html = ""
#EndRegion Global Variables and Constants
#Region Core functions
; #FUNCTION# ;===============================================================================
;
; Name...........:	_InsightsSetup()
; Description ...:	Setup activities including cURL initialization.
; Syntax.........:	_InsightsSetup()
; Parameters ....:
; Return values .: 	On Success			- Returns True.
;                 	On Failure			- Returns False, and:
;											sets @ERROR = 1 if unable to find an active SAP session.
;												This means the SAP GUI Scripting interface is not enabled.
;												Refer to the "Requirements" section at the top of this file.
;											sets @ERROR = 2 if unable to find the SAP window to attach to.
;
; Author ........:	seangriffin
; Modified.......:
; Remarks .......:	A prerequisite is that the SAP GUI Scripting interface is enabled,
;					and the SAP user is already logged in (ie. The "SAP Easy Access" window is displayed).
;					Refer to the "Requirements" section at the top of this file for information
;					on enabling the SAP GUI Scripting interface.
;
; Related .......:
; Link ..........:
; Example .......:	Yes
;
; ;==========================================================================================
Func _TestRailSetup()

	; Initialise cURL
	cURL_initialise()


EndFunc

; #FUNCTION# ;===============================================================================
;
; Name...........:	_InsightsShutdown()
; Description ...:	Setup activities including cURL initialization.
; Syntax.........:	_InsightsShutdown()
; Parameters ....:
; Return values .: 	On Success			- Returns True.
;                 	On Failure			- Returns False, and:
;											sets @ERROR = 1 if unable to find an active SAP session.
;												This means the SAP GUI Scripting interface is not enabled.
;												Refer to the "Requirements" section at the top of this file.
;											sets @ERROR = 2 if unable to find the SAP window to attach to.
;
; Author ........:	seangriffin
; Modified.......:
; Remarks .......:	A prerequisite is that the SAP GUI Scripting interface is enabled,
;					and the SAP user is already logged in (ie. The "SAP Easy Access" window is displayed).
;					Refer to the "Requirements" section at the top of this file for information
;					on enabling the SAP GUI Scripting interface.
;
; Related .......:
; Link ..........:
; Example .......:	Yes
;
; ;==========================================================================================
Func _TestRailShutdown()

	; Clean up cURL
	cURL_cleanup()

EndFunc


; #FUNCTION# ;===============================================================================
;
; Name...........:	_InsightsDomainSet()
; Description ...:	Sets the domain to use in all other functions.
; Syntax.........:	_InsightsDomainSet($domain)
; Parameters ....:	$win_title			- Optional: The title of the SAP window (within the session) to attach to.
;											The window "SAP Easy Access" is used if one isn't provided.
;											This may be a substring of the full window title.
;					$sap_transaction	- Optional: a SAP transaction to run after attaching to the session.
;											A "/n" will be inserted at the beginning of the transaction
;											if one isn't provided.
; Return values .: 	On Success			- Returns True.
;                 	On Failure			- Returns False, and:
;											sets @ERROR = 1 if unable to find an active SAP session.
;												This means the SAP GUI Scripting interface is not enabled.
;												Refer to the "Requirements" section at the top of this file.
;											sets @ERROR = 2 if unable to find the SAP window to attach to.
;
; Author ........:	seangriffin
; Modified.......:
; Remarks .......:	A prerequisite is that the SAP GUI Scripting interface is enabled,
;					and the SAP user is already logged in (ie. The "SAP Easy Access" window is displayed).
;					Refer to the "Requirements" section at the top of this file for information
;					on enabling the SAP GUI Scripting interface.
;
; Related .......:
; Link ..........:
; Example .......:	Yes
;
; ;==========================================================================================
Func _TestRailDomainSet($domain)

	$testrail_domain = $domain
EndFunc

; #FUNCTION# ;===============================================================================
;
; Name...........:	_InsightsLogin()
; Description ...:	Login a user to Janison Insights.
; Syntax.........:	_InsightsLogin($username, $password)
; Parameters ....:	$win_title			- Optional: The title of the SAP window (within the session) to attach to.
;											The window "SAP Easy Access" is used if one isn't provided.
;											This may be a substring of the full window title.
;					$sap_transaction	- Optional: a SAP transaction to run after attaching to the session.
;											A "/n" will be inserted at the beginning of the transaction
;											if one isn't provided.
; Return values .: 	On Success			- Returns True.
;                 	On Failure			- Returns False, and:
;											sets @ERROR = 1 if unable to find an active SAP session.
;												This means the SAP GUI Scripting interface is not enabled.
;												Refer to the "Requirements" section at the top of this file.
;											sets @ERROR = 2 if unable to find the SAP window to attach to.
;
; Author ........:	seangriffin
; Modified.......:
; Remarks .......:	A prerequisite is that the SAP GUI Scripting interface is enabled,
;					and the SAP user is already logged in (ie. The "SAP Easy Access" window is displayed).
;					Refer to the "Requirements" section at the top of this file for information
;					on enabling the SAP GUI Scripting interface.
;
; Related .......:
; Link ..........:
; Example .......:	Yes
;
; ;==========================================================================================
Func _TestRailLogin($username, $password)

	$testrail_username = $username
	$testrail_password = $password
EndFunc

; Authentication


Func _TestRailAuth()

;	$response = cURL_easy($testrail_domain, "cookies.txt", 2, 0, "", "Content-Type: text/html", "name=sgriffin@janison.com&password=Gri01ffo&rememberme=1", 0, 1, 0, $testrail_username & ":" & $testrail_password)
;	Exit

	Local $iPID = Run('curl.exe -k -u ' & $testrail_username & ':' & $testrail_password & ' ' & $testrail_domain & '/index.php?/auth/login -c cookies.txt', @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
    ProcessWaitClose($iPID)
    $testrail_json = StdoutRead($iPID)

	Local $iPID = Run('curl.exe -k -u ' & $testrail_username & ':' & $testrail_password & ' ' & $testrail_domain & '/index.php?/auth/login -c cookies.txt -d "name=' & $testrail_username & '&password=' & $testrail_password & '&rememberme=1" -X POST', @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
    ProcessWaitClose($iPID)
    $testrail_json = StdoutRead($iPID)

EndFunc



Func _TestRailGetAttachment($attachment_id)

;	$response = cURL_easy($testrail_domain, "cookies.txt", 2, 0, "", "Content-Type: text/html", "name=sgriffin@janison.com&password=Gri01ffo&rememberme=1", 0, 1, 0, $testrail_username & ":" & $testrail_password)
;	Exit

	Local $iPID = Run('curl.exe -k https://janison.testrail.com/index.php?/attachments/get/' & $attachment_id & ' -b cookies.txt -D ' & $attachment_id & '.header -o ' & $attachment_id & '.image', @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
    ProcessWaitClose($iPID)
    $testrail_json = StdoutRead($iPID)

EndFunc


; Projects

Func _TestRailGetProjects()

;	$response = cURL_easy($testrail_domain & "/index.php?/api/v2/get_projects", "", 0, 0, "", "Content-Type: application/json", "", 0, 1, 0, $testrail_username & ":" & $testrail_password)
;	$testrail_json = $response[2]

;	Local $debug = 'curl.exe -k -H "Content-Type: application/json" -u ' & $testrail_username & ':' & $testrail_password & ' ' & $testrail_domain & '/index.php?/api/v2/get_projects'

	Local $iPID = Run('curl.exe -k -H "Content-Type: application/json" -u ' & $testrail_username & ':' & $testrail_password & ' ' & $testrail_domain & '/index.php?/api/v2/get_projects', @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
    ProcessWaitClose($iPID)
    $testrail_json = StdoutRead($iPID)



EndFunc

Func _TestRailGetProjectsIDAndNameArray()

	Local $output[0][2]

	_TestRailGetProjects()

	$testrail_json = "{""projects"":" & $testrail_json & "}"
	Local $decoded_json = Json_Decode($testrail_json)

	for $i = 0 to 99999

		Local $id = Json_Get($decoded_json, '.projects[' & $i & '].id')

		if StringLen($id) < 1 Then ExitLoop

		Local $name = Json_Get($decoded_json, '.projects[' & $i & '].name')
		Local $is_completed = Json_Get($decoded_json, '.projects[' & $i & '].is_completed')

		if StringCompare($is_completed, "false") = 0 Then

			_ArrayAdd($output, $id & Chr(28) & $name, 0, Chr(28), @CRLF, 1)
		EndIf
	Next

;	_ArrayDisplay($output)
	Return $output
EndFunc

; Suites

Func _TestRailGetSuitesIdName($project_id)

;	$response = cURL_easy($testrail_domain & "/index.php?/api/v2/get_suites/" & $project_id, "", 0, 0, "", "Content-Type: application/json", "", 0, 1, 0, $testrail_username & ":" & $testrail_password)


	Local $iPID = Run('curl.exe -k -H "Content-Type: application/json" -u ' & $testrail_username & ':' & $testrail_password & ' ' & $testrail_domain & '/index.php?/api/v2/get_suites/' & $project_id, @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
    ProcessWaitClose($iPID)
    $testrail_json = StdoutRead($iPID)

	$rr = StringRegExp($testrail_json, '(?U)"id":(.*),.*"name":"(.*)",', 3)
	Return $rr


EndFunc

Func _TestRailGetCases($project_id, $suite_id, $section_id)

	Local $iPID = Run('curl.exe -k -H "Content-Type: application/json" -u ' & $testrail_username & ':' & $testrail_password & ' ' & $testrail_domain & '/index.php?/api/v2/get_cases/' & $project_id & '&suite_id=' & $suite_id & '&section_id=' & $section_id, @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
    ProcessWaitClose($iPID)
    $testrail_json = StdoutRead($iPID)
EndFunc



Func _TestRailGetCasesIDTitleTestScenarioIDOwnerObjectivesPreconditionsNotesSteps($project_id, $suite_id, $section_id, $case_id_arr = Null, $user_arr = Null)

	Local $output[0][2]

	_TestRailGetCases($project_id, $suite_id, $section_id)

	$testrail_json = "{""cases"":" & $testrail_json & "}"
	Local $decoded_json = Json_Decode($testrail_json)

	for $i = 0 to 99999

		Local $id = Json_Get($decoded_json, '.cases[' & $i & '].id')

		if StringLen($id) < 1 Then ExitLoop

		if _ArraySearch($case_id_arr, $id) > -1 Then

			Local $title = Json_Get($decoded_json, '.cases[' & $i & '].title')
			Local $test_scenario_id = Json_Get($decoded_json, '.cases[' & $i & '].custom_test_scenario_id')
			Local $owner_id = Json_Get($decoded_json, '.cases[' & $i & '].custom_owner')

			Local $owner_name = ""
			Local $owner_name_index = _ArraySearch($user_arr, $owner_id, 0, 0, 0, 0, 1, 0)

			if $owner_name_index > -1 Then

				$owner_name = $user_arr[$owner_name_index][1]
			EndIf

			Local $objectives = Json_Get($decoded_json, '.cases[' & $i & '].custom_objectives')
			Local $preconditions = Json_Get($decoded_json, '.cases[' & $i & '].custom_preconds')
			Local $notes = Json_Get($decoded_json, '.cases[' & $i & '].custom_notes')

			_ArrayAdd($output, "Test Case Start" & Chr(28), 0, Chr(28), Chr(30), 1)
			_ArrayAdd($output, $id & Chr(28), 0, Chr(28), Chr(30), 1)
			_ArrayAdd($output, $title & Chr(28), 0, Chr(28), Chr(30), 1)
			_ArrayAdd($output, $test_scenario_id & Chr(28), 0, Chr(28), Chr(30), 1)
			_ArrayAdd($output, $owner_name & Chr(28), 0, Chr(28), Chr(30), 1)
			_ArrayAdd($output, $objectives & Chr(28), 0, Chr(28), Chr(30), 1)
			_ArrayAdd($output, $preconditions & Chr(28), 0, Chr(28), Chr(30), 1)
			_ArrayAdd($output, $notes & Chr(28), 0, Chr(28), Chr(30), 1)

			for $j = 0 to 99999

				Local $step = Json_Get($decoded_json, '.cases[' & $i & '].custom_steps_separated[' & $j & '].content')

				if StringLen($step) < 1 Then ExitLoop

				Local $expected_result = Json_Get($decoded_json, '.cases[' & $i & '].custom_steps_separated[' & $j & '].expected')
				_ArrayAdd($output, $step & Chr(28) & $expected_result, 0, Chr(28), Chr(30), 1)
			Next
		EndIf
	Next

;	_ArrayDisplay($output)
	Return $output

EndFunc




Func _TestRailGetCasesIdTitle($project_id, $suite_id, $section_id)

	Local $output[0][2]

	_TestRailGetCases($project_id, $suite_id, $section_id)

	$testrail_json = "{""cases"":" & $testrail_json & "}"
	Local $decoded_json = Json_Decode($testrail_json)

	for $i = 0 to 99999

		Local $id = Json_Get($decoded_json, '.cases[' & $i & '].id')

		if StringLen($id) < 1 Then ExitLoop

		Local $title = Json_Get($decoded_json, '.cases[' & $i & '].title')

		_ArrayAdd($output, $id & Chr(28) & $title, 0, Chr(28), @CRLF, 1)
	Next

	Return $output


EndFunc

Func _TestRailGetCasesIdTitleReferences($project_id, $suite_id)

;	$response = cURL_easy($testrail_domain & "/index.php?/api/v2/get_cases/" & $project_id & "&suite_id=" & $suite_id, "", 0, 0, "", "Content-Type: application/json", "", 0, 1, 0, $testrail_username & ":" & $testrail_password)


	Local $iPID = Run('curl.exe -k -H "Content-Type: application/json" -u ' & $testrail_username & ':' & $testrail_password & ' ' & $testrail_domain & '/index.php?/api/v2/get_cases/' & $project_id & '&suite_id=' & $suite_id, @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
    ProcessWaitClose($iPID)
    $testrail_json = StdoutRead($iPID)

	$rr = StringRegExp($testrail_json, '(?U)"id":(.*),.*"title":"(.*)",.*"refs":"(.*)"', 3)
	Return $rr
EndFunc

Func _TestRailGetCase($case_id)

	Local $iPID = Run('curl.exe -k -H "Content-Type: application/json" -u ' & $testrail_username & ':' & $testrail_password & ' ' & $testrail_domain & '/index.php?/api/v2/get_case/' & $case_id, @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
    ProcessWaitClose($iPID)
    $testrail_json = StdoutRead($iPID)
EndFunc



Func _TestRailGetCaseIDTitleTestScenarioIDOwnerObjectivesPreconditionsNotesSteps($case_id)

	Local $output[0][2]

	_TestRailGetCase($case_id)

	$testrail_json = "{""case"":" & $testrail_json & "}"
	Local $decoded_json = Json_Decode($testrail_json)
	Local $id = Json_Get($decoded_json, '.case.id')
	Local $title = Json_Get($decoded_json, '.case.title')
	Local $test_scenario_id = Json_Get($decoded_json, '.case.custom_test_scenario_id')
	Local $owner_id = Json_Get($decoded_json, '.case.custom_owner')
	Local $owner_name = _TestRailGetUserName($owner_id)
	Local $objectives = Json_Get($decoded_json, '.case.custom_objectives')
	Local $preconditions = Json_Get($decoded_json, '.case.custom_preconds')
	Local $notes = Json_Get($decoded_json, '.case.custom_notes')

	_ArrayAdd($output, $id & Chr(28), 0, Chr(28), Chr(30), 1)
	_ArrayAdd($output, $title & Chr(28), 0, Chr(28), Chr(30), 1)
	_ArrayAdd($output, $test_scenario_id & Chr(28), 0, Chr(28), Chr(30), 1)
	_ArrayAdd($output, $owner_name & Chr(28), 0, Chr(28), Chr(30), 1)
	_ArrayAdd($output, $objectives & Chr(28), 0, Chr(28), Chr(30), 1)
	_ArrayAdd($output, $preconditions & Chr(28), 0, Chr(28), Chr(30), 1)
	_ArrayAdd($output, $notes & Chr(28), 0, Chr(28), Chr(30), 1)

	for $i = 0 to 99999

		Local $step = Json_Get($decoded_json, '.case.custom_steps_separated[' & $i & '].content')

		if StringLen($step) < 1 Then ExitLoop

		Local $expected_result = Json_Get($decoded_json, '.case.custom_steps_separated[' & $i & '].expected')
		_ArrayAdd($output, $step & Chr(28) & $expected_result, 0, Chr(28), Chr(30), 1)
	Next

;	_ArrayDisplay($output)
	Return $output

EndFunc





Func _TestRailGetRunsIdName($project_id)

	Local $iPID = Run('curl.exe -k -H "Content-Type: application/json" -u ' & $testrail_username & ':' & $testrail_password & ' ' & $testrail_domain & '/index.php?/api/v2/get_runs/' & $project_id, @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
    ProcessWaitClose($iPID)
    $testrail_json = StdoutRead($iPID)

	$rr = StringRegExp($testrail_json, '(?U)"id":(.*),.*"name":"(.*)"', 3)
	Return $rr

EndFunc

Func _TestRailGetRun($run_id)

	$response = cURL_easy($testrail_domain & "/index.php?/api/v2/get_run/" & $run_id, "", 0, 0, "", "Content-Type: application/json", "", 0, 1, 0, $testrail_username & ":" & $testrail_password)
EndFunc

Func _TestRailGetRunIDFromPlanIDAndRunName($plan_id, $run_name)

	_TestRailGetPlan($plan_id)

	$rr = StringRegExp($testrail_json, '"runs":\[{"id":\d+,.*"name":"' & $run_name & '"', 1)
	$rr[0] = StringLeft($rr[0], StringInStr($rr[0], ",") - 1)
	$rr[0] = StringMid($rr[0], StringInStr($rr[0], ":", 0, 2) + 1)

	Return $rr[0]


EndFunc

Func _TestRailGetPlanRunsID($plan_id)

	_TestRailGetPlan($plan_id)

	$rr = StringRegExp($testrail_json, '(?U)"id":(\d+),"suite_id":\d+,"name":".*","description"', 3)

	return $rr
EndFunc

Func _TestRailGetPlanRunsIDAndNameArray($plan_id)

	_TestRailGetPlan($plan_id)

	$rr = StringRegExp($testrail_json, '(?U)"id":(\d+),"suite_id":\d+,"name":"(.*)","description"', 3)

	return $rr
EndFunc


Func _TestRailGetResults($test_id)

	$response = cURL_easy($testrail_domain & "/index.php?/api/v2/get_results/" & $test_id, "", 0, 0, "", "Content-Type: application/json", "", 0, 1, 0, $testrail_username & ":" & $testrail_password)
EndFunc

Func _TestRailGetResultsForCaseIdTestIdStatusIdDefects($run_id, $case_id)

	Local $iPID = Run('curl.exe -k -H "Content-Type: application/json" -u ' & $testrail_username & ':' & $testrail_password & ' ' & $testrail_domain & '/index.php?/api/v2/get_results_for_case/' & $run_id & '/' & $case_id, @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
    ProcessWaitClose($iPID)
    $testrail_json = StdoutRead($iPID)

	$rr = StringRegExp($testrail_json, '(?U)"defects":"(.*)","id":(.*),"status_id":(.*),"test_id":(.*),', 3)
	Return $rr


EndFunc

Func _TestRailGetResultsIdStatusIdDefects($test_id)

	Local $iPID = Run('curl.exe -k -H "Content-Type: application/json" -u ' & $testrail_username & ':' & $testrail_password & ' ' & $testrail_domain & '/index.php?/api/v2/get_results/' & $test_id & '&limit=1', @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
    ProcessWaitClose($iPID)
    $testrail_json = StdoutRead($iPID)
	filedelete("D:\dwn\fred.txt")
	filewrite("D:\dwn\fred.txt", $testrail_json)
;	Exit

	$rr = StringRegExp($testrail_json, '(?U)"id":(.*),.*"test_id":(.*),.*"status_id":(.*),.*"defects":(.*),"custom_step_results"', 3)
;	_ArrayDisplay($rr)
;	Exit
	Return $rr


EndFunc

Func _TestRailGetResultsForRunIdStatusIdCreatedOnDefects($run_id)

;	Local $iPID = Run('curl.exe -k -H "Content-Type: application/json" -u ' & $testrail_username & ':' & $testrail_password & ' ' & $testrail_domain & '/index.php?/api/v2/get_results_for_run/' & $run_id & '&limit=1', @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
	Local $iPID = Run('curl.exe -k -H "Content-Type: application/json" -u ' & $testrail_username & ':' & $testrail_password & ' ' & $testrail_domain & '/index.php?/api/v2/get_results_for_run/' & $run_id, @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
    ProcessWaitClose($iPID)
    $testrail_json = StdoutRead($iPID)
;	filedelete("D:\dwn\fred.txt")
;	filewrite("D:\dwn\fred.txt", $testrail_json)
;	Exit

	$rr = StringRegExp($testrail_json, '(?U)"id":(.*),.*"test_id":(.*),.*"status_id":(.*),.*"created_on":(.*),.*"defects":(.*),"custom_step_results"', 3)
;	_ArrayDisplay($rr)
;	Exit
	Return $rr


EndFunc

Func _TestRailGetTestsIdTitleCaseIdRunId($run_id)

	Local $iPID = Run('curl.exe -k -H "Content-Type: application/json" -u ' & $testrail_username & ':' & $testrail_password & ' ' & $testrail_domain & '/index.php?/api/v2/get_tests/' & $run_id, @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
    ProcessWaitClose($iPID)
    $testrail_json = StdoutRead($iPID)
;	filewrite("D:\dwn\fred.txt", $testrail_json)
;	Exit

	$rr = StringRegExp($testrail_json, '(?U)"id":(.*),.*"case_id":(.*),.*"run_id":(.*),.*"title":"(.*)",', 3)
	Return $rr


EndFunc

Func _TestRailGetTestsTitleAndIDFromRunID($run_id)

	Local $test_title_and_id_dict = ObjCreate("Scripting.Dictionary")

;	_TestRailGetTests($run_id)

	$rr = StringRegExp($testrail_json, '(?U)"id":\d+,.*"title":".*"', 3)

	for $each in $rr

		Local $id = $each
		Local $title = $each

		$id = StringLeft($id, StringInStr($id, ",") - 1)
		$id = StringMid($id, StringInStr($id, ":") + 1)
		$title = StringMid($title, StringInStr($title, ":", 0, -1) + 1)
		$title = StringReplace($title, """", "")
		$test_title_and_id_dict.Add($title, $id)
	Next

	Return $test_title_and_id_dict

EndFunc

Func _TestRailGetTestsReferenceAndIDFromRunID($run_id)

	Local $test_refs_and_id_dict = ObjCreate("Scripting.Dictionary")

;	_TestRailGetTests($run_id)

	$rr = StringRegExp($testrail_json, '(?U)"id":\d+,.*"refs":".*"', 3)

	for $each in $rr

		Local $id = $each
		Local $refs = $each

		$id = StringLeft($id, StringInStr($id, ",") - 1)
		$id = StringMid($id, StringInStr($id, ":") + 1)
		$refs = StringMid($refs, StringInStr($refs, ":", 0, -1) + 1)
		$refs = StringReplace($refs, """", "")
		$test_refs_and_id_dict.Add($refs, $id)
	Next

	Return $test_refs_and_id_dict

EndFunc

Func _TestRailGetPlan($plan_id)

	Local $iPID = Run('curl.exe -k -H "Content-Type: application/json" -u ' & $testrail_username & ':' & $testrail_password & ' ' & $testrail_domain & '/index.php?/api/v2/get_plan/' & $plan_id, @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
    ProcessWaitClose($iPID)
    $testrail_json = StdoutRead($iPID)

EndFunc

Func _TestRailGetPlans($project_id)

	Local $iPID = Run('curl.exe -k -H "Content-Type: application/json" -u ' & $testrail_username & ':' & $testrail_password & ' ' & $testrail_domain & '/index.php?/api/v2/get_plans/' & $project_id, @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
    ProcessWaitClose($iPID)
    $testrail_json = StdoutRead($iPID)


EndFunc

Func _TestRailGetPlansIDAndNameArray($project_id)

	Local $output[0][2]

	_TestRailGetPlans($project_id)

	$rr = StringRegExp($testrail_json, '(?U)"id":\d+,.*"name":".*"', 3)

	if IsArray($rr) Then

		for $each in $rr

			Local $id = $each
			Local $name = $each

			$id = StringLeft($id, StringInStr($id, ",") - 1)
			$id = StringMid($id, StringInStr($id, ":") + 1)
			$name = StringMid($name, StringInStr($name, ":", 0, -1) + 1)
			$name = StringReplace($name, """", "")
			Local $id_name = $id & "|" & $name
			_ArrayAdd($output, $id_name)
		Next
	EndIf

	Return $output
EndFunc

Func _TestRailGetPlanIDByName($project_id, $plan_name)

	_TestRailGetPlans($project_id)

	$rr = StringRegExp($testrail_json, '"id":\d+,"name":"' & $plan_name & '"', 1)
	$rr[0] = StringLeft($rr[0], StringInStr($rr[0], ",") - 1)
	$rr[0] = StringMid($rr[0], StringInStr($rr[0], ":") + 1)

	Return $rr[0]
EndFunc

Func _TestRailAddResult($test_id, $status_id)

	$response = cURL_easy($testrail_domain & "/index.php?/api/v2/add_result/" & $test_id, "", 0, 0, "", "Content-Type: application/json", '{"status_id":' & $status_id & '}', 0, 1, 0, $testrail_username & ":" & $testrail_password)
EndFunc

Func _TestRailAddResults($run_id, $results_arr)

	; create the results JSON to post from the results array

	Local $results_json = '{"results":['

	for $i = 0 to (UBound($results_arr) - 1) step 3

		if StringLen($results_json) > StringLen('{"results":[') Then

			$results_json = $results_json & ','
		EndIf

		$results_json = $results_json & '{"test_id":' & $results_arr[$i + 0] & ',"status_id":' & $results_arr[$i + 1] & ',"comment":"' & $results_arr[$i + 2] & '"}'
	Next

	$results_json = $results_json & ']}'

	FileDelete(@ScriptDir & "\curl_in.json")
	FileWrite(@ScriptDir & "\curl_in.json", $results_json)
	Local $iPID = Run('curl.exe -s -k -H "Content-Type: application/json" --data @curl_in.json -u ' & $testrail_username & ':' & $testrail_password & ' ' & $testrail_domain & '/index.php?/api/v2/add_results/' & $run_id, @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
    ProcessWaitClose($iPID)
    $testrail_json = StdoutRead($iPID)

	; unfortunately the below is unreliable.  Working intermittently
;	$response = cURL_easy($testrail_domain & "/index.php?/api/v2/add_results/" & $run_id, "", 0, 0, "", "Content-Type: application/json", $results_json, 0, 1, 0, $testrail_username & ":" & $testrail_password)
EndFunc


Func _TestRailGetIdFromTitle($json, $title)

	$rr = StringRegExp($json, '"id":.*"title":"' & $title & '"', 1)
	$tt = StringMid($rr[0], StringLen('"id":') + 1, StringInStr($rr[0], ",") - (StringLen('"id":') + 1))
	Return $tt

EndFunc

Func _TestRailGetStatusesIdLabel()

	Local $iPID = Run('curl.exe -k -H "Content-Type: application/json" -u ' & $testrail_username & ':' & $testrail_password & ' ' & $testrail_domain & '/index.php?/api/v2/get_statuses/', @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
    ProcessWaitClose($iPID)
    $testrail_json = StdoutRead($iPID)
;	filedelete("D:\dwn\fred.txt")
;	filewrite("D:\dwn\fred.txt", $testrail_json)
;	Exit

	$rr = StringRegExp($testrail_json, '(?U)"id":(.*),.*"label":"(.*)"', 3)
	Return $rr


EndFunc

Func _TestRailGetStatusLabelAndID()

	Local $status_label_and_id_dict = ObjCreate("Scripting.Dictionary")

;	_TestRailGetStatuses()

	$rr = StringRegExp($testrail_json, '(?U)"id":\d+,.*"label":".*"', 3)

	for $each in $rr

		Local $id = $each
		Local $label = $each

		$id = StringLeft($id, StringInStr($id, ",") - 1)
		$id = StringMid($id, StringInStr($id, ":") + 1)
		$label = StringMid($label, StringInStr($label, ":", 0, -1) + 1)
		$label = StringReplace($label, """", "")

		$status_label_and_id_dict.Add($label, $id)
	Next

	Return $status_label_and_id_dict

;	_ArrayDisplay($rr)

EndFunc


Func _TestRailGetSections($project_id)

	Local $iPID = Run('curl.exe -k -H "Content-Type: application/json" -u ' & $testrail_username & ':' & $testrail_password & ' ' & $testrail_domain & '/index.php?/api/v2/get_sections/' & $project_id, @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
    ProcessWaitClose($iPID)
    $testrail_json = StdoutRead($iPID)
EndFunc


Func _TestRailGetSectionNameAndDepth($project_id)

	Local $output[0][2]

	_TestRailGetSections($project_id)

	$testrail_json = "{""projects"":" & $testrail_json & "}"
	Local $decoded_json = Json_Decode($testrail_json)

	for $i = 0 to 99999

		Local $id = Json_Get($decoded_json, '.projects[' & $i & '].id')

		if @error > 0 Then ExitLoop

		Local $name = Json_Get($decoded_json, '.projects[' & $i & '].name')
		Local $depth = Json_Get($decoded_json, '.projects[' & $i & '].depth')

		Local $depth_str = _StringRepeat("-", $depth)

		if StringLen($depth_str) > 0 Then

			$depth_str = $depth_str & " "
		EndIf

		_ArrayAdd($output, $id & Chr(28) & $depth_str & $name, 0, Chr(28), @CRLF, 1)
	Next

;	_ArrayDisplay($output)
	Return $output


EndFunc



Func _TestRailGetSectionIDAndName($project_id)

	Local $output[0][2]

	_TestRailGetSections($project_id)

	$testrail_json = "{""sections"":" & $testrail_json & "}"
	Local $decoded_json = Json_Decode($testrail_json)

	for $i = 0 to 99999

		Local $id = Json_Get($decoded_json, '.sections[' & $i & '].id')

		if @error > 0 Then ExitLoop

		Local $name = Json_Get($decoded_json, '.sections[' & $i & '].name')
		Local $depth = Json_Get($decoded_json, '.sections[' & $i & '].depth')

		Local $depth_str = _StringRepeat("-", $depth)

		if StringLen($depth_str) > 0 Then

			$depth_str = $depth_str & " "
		EndIf

		_ArrayAdd($output, $id & Chr(28) & $depth_str & $name, 0, Chr(28), @CRLF, 1)
	Next

;	_ArrayDisplay($output)
	Return $output

EndFunc




Func _TestRailGetUser($user_id)

	Local $iPID = Run('curl.exe -k -H "Content-Type: application/json" -u ' & $testrail_username & ':' & $testrail_password & ' ' & $testrail_domain & '/index.php?/api/v2/get_user/' & $user_id, @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
    ProcessWaitClose($iPID)
    $testrail_json = StdoutRead($iPID)
EndFunc


Func _TestRailGetUserName($user_id)

	Local $output[0][2]

	_TestRailGetUser($user_id)

	$testrail_json = "{""user"":" & $testrail_json & "}"
	Local $decoded_json = Json_Decode($testrail_json)
	Local $name = Json_Get($decoded_json, '.user.name')

	Return $name

EndFunc


Func _TestRailGetUsers()

	Local $iPID = Run('curl.exe -k -H "Content-Type: application/json" -u ' & $testrail_username & ':' & $testrail_password & ' ' & $testrail_domain & '/index.php?/api/v2/get_users', @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
    ProcessWaitClose($iPID)
    $testrail_json = StdoutRead($iPID)
EndFunc


Func _TestRailGetUsersIDName()

	Local $output[0][2]

	_TestRailGetUsers()

	$testrail_json = "{""users"":" & $testrail_json & "}"
	Local $decoded_json = Json_Decode($testrail_json)

	for $i = 0 to 99999

		Local $id = Json_Get($decoded_json, '.users[' & $i & '].id')

		if StringLen($id) < 1 Then ExitLoop

		Local $name = Json_Get($decoded_json, '.users[' & $i & '].name')

		_ArrayAdd($output, $id & Chr(28) & $name, 0, Chr(28), @CRLF, 1)
	Next

;	_ArrayDisplay($output)
	Return $output

EndFunc


; TestRail Markdown



Func _TestRailMarkdownToHTML($markdown)

	; reformat testrail tables slightly
	$markdown_line = StringSplit($markdown, @CRLF, 3)

	for $i = 0 to (UBound($markdown_line) - 1)

		if StringRegExp($markdown_line[$i], "^\|\|\|") = 1 Then

			$markdown_line[$i] = "| " & StringReplace(StringMid($markdown_line[$i], 4), ":", " ")
			_ArrayInsert($markdown_line, $i + 1, StringRegExpReplace($markdown_line[$i], "[^|]", "-"), 0, chr(28), chr(28))
			$i = $i + 1
		EndIf

		$markdown_line[$i] = StringRegExpReplace($markdown_line[$i], "^\|\|", "|")
	Next

	$markdown = _ArrayToString($markdown_line, @CRLF)

	; adding a newline at the front of the markdown text, to enforce no metadata is read from line #1
	$markdown = @CRLF & $markdown

	FileDelete(@ScriptDir & "\tmp.html")
	FileDelete(@ScriptDir & "\tmp.md")
	FileWrite(@ScriptDir & "\tmp.md", $markdown)
	ShellExecuteWait("multimarkdown.exe", "tmp.md -o tmp.html", @ScriptDir, "", @SW_HIDE)
	FileDelete(@ScriptDir & "\tmp.md")
	Local $tmp_html = FileRead(@ScriptDir & "\tmp.html")
	FileDelete(@ScriptDir & "\tmp.html")

	; remove the additional paragraph tags multimarkdown adds at the start and end
	$tmp_html = StringRegExpReplace($tmp_html, "(?s)<p>(.*?)</p>", "\1")

	; add missing <br> newlines
	$tmp_html_line = StringSplit($tmp_html, @CRLF, 3)

	for $i = 0 to (UBound($tmp_html_line) - 1)

		if StringRegExp($tmp_html_line[$i], "[^>]$") = 1 or StringRegExp($tmp_html_line[$i], "</strong>$") = 1 Then

			$tmp_html_line[$i] = $tmp_html_line[$i] & "<br>"
		EndIf
	Next

	$tmp_html = _ArrayToString($tmp_html_line, @CRLF)

	; reformat <ol> tags to not indent or have margins

	$tmp_html = StringReplace($tmp_html, "<ol>", "<ol style=""list-style-position: inside; padding-left: 0; margin-top: 0em; margin-bottom: 0em;"">")

	; get all attachments

;	Local $attachment_id = StringRegExp($tmp_html, "<img src=""https://janison.testrail.com/index.php\?/attachments/get/(\d+)"" alt="""" />", 1)
	Local $attachment_id = StringRegExp($tmp_html, "<img src=""index.php\?/attachments/get/(\d+)"" alt="""" />", 1)

	if @error = 0 Then

		for $i = 0 to (UBound($attachment_id) - 1)

			_TestRailGetAttachment($attachment_id[$i])
			$attachment_header = FileRead($attachment_id[$i] & ".header")
			FileDelete(@ScriptDir & "\" & $attachment_id[$i] & ".header")
			Local $content_type = StringRegExp($attachment_header, "(?s)Content-Type: (.*?)\n", 1)
			Local $image_file_extension = ""

			Switch StringStripWS($content_type[0], 8)

				Case "image/gif"

					$image_file_extension = ".gif"

				Case "image/jpeg"

					$image_file_extension = ".jpg"

				Case "image/png"

					$image_file_extension = ".png"

			EndSwitch

			if StringLen($image_file_extension) > 0 Then

				FileDelete(@ScriptDir & "\" & $attachment_id[$i] & $image_file_extension)
				FileMove(@ScriptDir & "\" & $attachment_id[$i] & ".image", @ScriptDir & "\" & $attachment_id[$i] & $image_file_extension, 1)
				$tmp_html = StringReplace($tmp_html, "<img src=""index.php?/attachments/get/" & $attachment_id[$i] & """ alt="""" />", "<img src=""" & $attachment_id[$i] & $image_file_extension & """ alt="""" />")
			EndIf
		Next
	EndIf


;	$tmp_html = StringReplace($tmp_html, @CRLF & "<h1", "<h1")
;	$tmp_html = StringReplace($tmp_html, "</h1>" & @CRLF, "</h1>")
;	$tmp_html = StringReplace($tmp_html, @CRLF & "<h2", "<h2")
;	$tmp_html = StringReplace($tmp_html, "</h2>" & @CRLF, "</h2>")
;	$tmp_html = StringReplace($tmp_html, @CRLF & "<h3", "<h3")
;	$tmp_html = StringReplace($tmp_html, "</h3>" & @CRLF, "</h3>")
;	$tmp_html = StringReplace($tmp_html, @CRLF & "<hr />" & @CRLF, "<hr />")
;	$tmp_html = StringReplace($tmp_html, @CRLF & "<li>", "<li>")
;	$tmp_html = StringReplace($tmp_html, "</li>" & @CRLF, "</li>")
;	$tmp_html = StringReplace($tmp_html, @CRLF & "<blockquote>", "<blockquote>")
;	$tmp_html = StringReplace($tmp_html, "</blockquote>" & @CRLF, "</blockquote>")
;	$tmp_html = StringReplace($tmp_html, @CRLF & "<figure>", "<figure>")
;	$tmp_html = StringReplace($tmp_html, "</figure>" & @CRLF, "</figure>")


;	$tmp_html = StringReplace($tmp_html, @CRLF, "<br>")
;	$tmp_html = StringReplace($tmp_html, @LF, "<br>")

	Return $tmp_html
EndFunc



; Jira integration



Func _TestRailGetTestCases($key)

	$response = cURL_easy($testrail_domain & "/index.php?/ext/jira/render_panel&ae=connect&av=1&issue=" & $key & "&panel=references&login=button&frame=tr-frame-panel-references", "cookies.txt", 1, 0, "", "Content-Type: text/html; charset=UTF-8", "", 0, 1, 0, $testrail_username & ":" & $testrail_password)
	$testrail_html = $response[2]
EndFunc




