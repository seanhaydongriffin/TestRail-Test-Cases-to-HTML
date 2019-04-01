#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_UseUpx=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#RequireAdmin
;#AutoIt3Wrapper_usex64=n
#include <Date.au3>
#include <File.au3>
#include <Array.au3>
#Include "Json.au3"
#include "Jira.au3"
#include "Confluence.au3"
#include "TestRail.au3"
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <GuiListView.au3>
#include <ButtonConstants.au3>
#include <ComboConstants.au3>
#include <SQLite.au3>
#include <SQLite.dll.au3>
#include <Crypt.au3>
#include <GuiComboBox.au3>
#include <String.au3>



Global $html, $markup, $storage_format
Global $app_name = "TestRail Test Cases to HTML"
Global $ini_filename = @ScriptDir & "\" & $app_name & ".ini"
Global $aResult, $iRows, $iColumns, $iRval, $run_name = "", $max_num_defects = 0, $max_num_days = 0, $version_name = ""

Global $main_gui = GUICreate($app_name, 860, 600)

GUICtrlCreateGroup("TestRail Login", 10, 10, 360, 180)
GUICtrlCreateLabel("Username", 20, 30, 60, 20)
Global $testrail_username_input = GUICtrlCreateInput(IniRead($ini_filename, "main", "testrailusername", "sgriffin@janison.com"), 100, 30, 250, 20)
GUICtrlCreateLabel("Password", 20, 50, 60, 20)
Global $testrail_password_input = GUICtrlCreateInput("", 100, 50, 250, 20, $ES_PASSWORD)
Global $testrail_authenticate_button = GUICtrlCreateButton("Authenticate", 100, 70, 80, 20)
;GUICtrlCreateLabel("TestRail Project", 20, 70, 100, 20)
;Global $testrail_project_combo = GUICtrlCreateCombo("", 140, 70, 250, 20, BitOR($CBS_DROPDOWNLIST, $WS_VSCROLL))
GUICtrlCreateGroup("", -99, -99, 1, 1)

GUICtrlCreateGroup("TestRail Projects", 390, 10, 450, 180)
Global $testrail_project_listview = GUICtrlCreateListView("ID|Name", 410, 30, 410, 150, $LVS_SHOWSELALWAYS)
_GUICtrlListView_SetColumnWidth(-1, 0, 80)
_GUICtrlListView_SetColumnWidth(-1, 1, 300)
_GUICtrlListView_SetExtendedListViewStyle($testrail_project_listview, BitOR($LVS_EX_GRIDLINES, $LVS_EX_FULLROWSELECT, $LVS_EX_SUBITEMIMAGES))
GUICtrlCreateGroup("", -99, -99, 1, 1)

GUICtrlCreateGroup("TestRail Sections", 10, 200, 360, 360)
Global $testrail_section_listview = GUICtrlCreateListView("ID|Name", 30, 220, 320, 300, $LVS_SHOWSELALWAYS)
_GUICtrlListView_SetColumnWidth(-1, 0, 80)
_GUICtrlListView_SetColumnWidth(-1, 1, 300)
_GUICtrlListView_SetExtendedListViewStyle($testrail_section_listview, BitOR($LVS_EX_GRIDLINES, $LVS_EX_FULLROWSELECT, $LVS_EX_SUBITEMIMAGES))
GUICtrlCreateGroup("", -99, -99, 1, 1)

Local $testrail_encrypted_password = IniRead($ini_filename, "main", "testrailpassword", "")
Global $testrail_decrypted_password = ""

if stringlen($testrail_encrypted_password) > 0 Then

	$testrail_decrypted_password = _Crypt_DecryptData($testrail_encrypted_password, "applesauce", $CALG_AES_256)
	$testrail_decrypted_password = BinaryToString($testrail_decrypted_password)
	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $testrail_decrypted_password = ' & $testrail_decrypted_password & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
	GUICtrlSetData($testrail_password_input, $testrail_decrypted_password)
Else

	$testrail_decrypted_password = ""
EndIf


GUICtrlCreateGroup("TestRail Cases", 390, 200, 450, 360)
Global $testrail_case_listview = GUICtrlCreateListView("ID|Title", 410, 220, 410, 300, $LVS_SHOWSELALWAYS)
_GUICtrlListView_SetColumnWidth(-1, 0, 80)
_GUICtrlListView_SetColumnWidth(-1, 1, 300)
_GUICtrlListView_SetExtendedListViewStyle($testrail_case_listview, BitOR($LVS_EX_GRIDLINES, $LVS_EX_FULLROWSELECT, $LVS_EX_SUBITEMIMAGES))
GUICtrlCreateGroup("", -99, -99, 1, 1)

Global $export_button = GUICtrlCreateButton("Export", 410, 530, 100, 20, -1) ;, $BS_DEFPUSHBUTTON)

Global $status_input = GUICtrlCreateInput("Enter the ""Epic Key"" and click ""Start""", 10, 600 - 25, 360, 20, $ES_READONLY, $WS_EX_STATICEDGE)
Global $progress = GUICtrlCreateProgress(390, 600 - 25, 450, 20)


GUISetState(@SW_SHOW, $main_gui)
GUICtrlSetData($status_input, "")
GUIRegisterMsg($WM_NOTIFY, "WM_NOTIFY")

; Loop until the user exits.
While 1

	; GUI msg loop...
	$msg = GUIGetMsg()

	Switch $msg

		Case $GUI_EVENT_CLOSE

			IniWrite($ini_filename, "main", "testrailusername", GUICtrlRead($testrail_username_input))
			$testrail_encrypted_password = _Crypt_EncryptData(GUICtrlRead($testrail_password_input), "applesauce", $CALG_AES_256)
			IniWrite($ini_filename, "main", "testrailpassword", $testrail_encrypted_password)

			ExitLoop

		Case $testrail_authenticate_button

			disable_gui()

			; Startup TestRail

			GUICtrlSetData($status_input, "Starting the TestRail connection ... ")
			_TestRailDomainSet("https://janison.testrail.com")
			_TestRailLogin(GUICtrlRead($testrail_username_input), GUICtrlRead($testrail_password_input))

			if StringLen(GUICtrlRead($testrail_password_input)) > 0 Then

				_GUICtrlListView_DeleteAllItems($testrail_project_listview)
				_GUICtrlListView_DeleteAllItems($testrail_section_listview)
				_GUICtrlListView_DeleteAllItems($testrail_case_listview)

				GUICtrlSetData($status_input, "Getting the TestRail Projects ... ")
				Local $project_arr = _TestRailGetProjectsIDAndNameArray()
				_GUICtrlListView_AddArray($testrail_project_listview, $project_arr)
				GUICtrlSetData($status_input, "")
			EndIf

			enable_gui()


		Case $export_button

			Create_Case_HTML()


;			GUICtrlSetData($progress, 0)
;			GUICtrlSetState($epic_key_input, $GUI_DISABLE)
;			GUICtrlSetState($export_button, $GUI_DISABLE)
;			GUISetCursor(15, 1, $main_gui)
;			_GUICtrlListView_DeleteAllItems($listview)

;			$each = "blank"
;			Local $pid = ShellExecute(@ScriptDir & "\data_extractor.exe", """" & GUICtrlRead($testrail_username_input) & """ """ & GUICtrlRead($testrail_password_input) & """ """ & $run_ids & """ """ & GUICtrlRead($jira_username_input) & """ """ & GUICtrlRead($jira_password_input) & """ """ & GUICtrlRead($epic_key_input) & """", "", "", @SW_HIDE)

#cs
			; populate listview with epic keys

			Local $epic_key = StringSplit(GUICtrlRead($epic_key_input), ",;|", 2)

			for $each in $epic_key

				Local $pid = ShellExecute(@ScriptDir & "\data_extractor.exe", """" & GUICtrlRead($testrail_username_input) & """ """ & GUICtrlRead($testrail_password_input) & """ """ & $run_ids & """ """ & GUICtrlRead($jira_username_input) & """ """ & GUICtrlRead($jira_password_input) & """ """ & $each & """", "", "", @SW_HIDE)
				GUICtrlCreateListViewItem($each & "|" & $pid & "|In Progress", $listview)
			Next

			While True

				Local $all_epics_done = True

				for $index = 0 to (_GUICtrlListView_GetItemCount($listview) - 1)

					Local $pid = _GUICtrlListView_GetItemText($listview, $index, 1)
					Local $status = _GUICtrlListView_GetItemText($listview, $index, 2)

					if StringCompare($status, "In Progress") = 0 Then

						$all_epics_done = False

						if ProcessExists($pid) = False Then

							_GUICtrlListView_SetItemText($listview, $index, "Done", 2)

						EndIf
					EndIf
				Next

				if $all_epics_done = True Then

					ExitLoop
				EndIf

				Sleep(1000)
			WEnd
#ce


;			GUICtrlSetData($progress, 0)
;			GUICtrlSetData($status_input, "")
;			GUICtrlSetState($epic_key_input, $GUI_ENABLE)
;			GUICtrlSetState($export_button, $GUI_ENABLE)
;			GUISetCursor(2, 0, $main_gui)


	EndSwitch

WEnd

GUIDelete($main_gui)

; Shutdown Jira

GUICtrlSetData($status_input, "Closing TestRail ... ")
_TestRailShutdown()


Func WM_NOTIFY($hWnd, $iMsg, $wParam, $lParam)
    #forceref $hWnd, $iMsg, $wParam
    Local $hWndFrom, $iIDFrom, $iCode, $tNMHDR, $hWndListView, $tInfo
    ; Local $tBuffer
;    $hWndListView = $g_hListView
 ;   If Not IsHWnd($g_hListView) Then $hWndListView = GUICtrlGetHandle($testrail_project_listview)

    $tNMHDR = DllStructCreate($tagNMHDR, $lParam)
    $hWndFrom = HWnd(DllStructGetData($tNMHDR, "hWndFrom"))
    $iIDFrom = DllStructGetData($tNMHDR, "IDFrom")
    $iCode = DllStructGetData($tNMHDR, "Code")
    Switch $hWndFrom
        Case GUICtrlGetHandle($testrail_project_listview)
            Switch $iCode
				Case $NM_CLICK ; Sent by a list-view control when the user clicks an item with the left mouse button

					disable_gui()
					_GUICtrlListView_DeleteAllItems($testrail_section_listview)
					_GUICtrlListView_DeleteAllItems($testrail_case_listview)

					Local $selected_project_id = _GUICtrlListView_GetItemText($testrail_project_listview, Number(_GUICtrlListView_GetSelectedIndices($testrail_project_listview)), 0)
					Local $selected_project_name = _GUICtrlListView_GetItemText($testrail_project_listview, Number(_GUICtrlListView_GetSelectedIndices($testrail_project_listview)), 1)
					ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $selected_project_name = ' & $selected_project_name & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console

					GUICtrlSetData($status_input, "Getting the TestRail Sections ... ")
					Local $section_arr = _TestRailGetSectionIDAndName($selected_project_id)
					_GUICtrlListView_AddArray($testrail_section_listview, $section_arr)
					GUICtrlSetData($status_input, "")

					enable_gui()

            EndSwitch

        Case GUICtrlGetHandle($testrail_section_listview)
            Switch $iCode
				Case $NM_CLICK ; Sent by a list-view control when the user clicks an item with the left mouse button

					disable_gui()
					_GUICtrlListView_DeleteAllItems($testrail_case_listview)

					Local $selected_project_id = _GUICtrlListView_GetItemText($testrail_project_listview, Number(_GUICtrlListView_GetSelectedIndices($testrail_project_listview)), 0)
					Local $selected_section_id = _GUICtrlListView_GetItemText($testrail_section_listview, Number(_GUICtrlListView_GetSelectedIndices($testrail_section_listview)), 0)

					GUICtrlSetData($status_input, "Getting the TestRail Cases ... ")
					Local $case_arr = _TestRailGetCasesIdTitle($selected_project_id, "", $selected_section_id)
					_GUICtrlListView_AddArray($testrail_case_listview, $case_arr)
					GUICtrlSetData($status_input, "")

					enable_gui()

            EndSwitch

    EndSwitch
    Return $GUI_RUNDEFMSG
EndFunc   ;==>WM_NOTIFY




Func Create_Case_HTML($confluence_html = False)

	Local $double_quotes = """"

	$selected_case = _GUICtrlListView_GetSelectedIndices($testrail_case_listview, true)
	_ArrayDelete($selected_case, 0)

	if UBound($selected_case) = 0 Then

		SplashTextOn($app_name, "You must select at least one case first.", 400, 50)
		Sleep(3000)
		SplashOff()
	Else

		Local $num_selected_cases = UBound($selected_case)
		Local $selected_case_num = 0

		disable_gui()

		for $each_case_index In $selected_case

			$selected_case_num = $selected_case_num + 1
			GUICtrlSetData($progress, ($selected_case_num / $num_selected_cases) * 100)

			Local $id = _GUICtrlListView_GetItemText($testrail_case_listview, $each_case_index, 0)
			Local $report_filename = @ScriptDir & "\TestRail " & $id & ".html"
			GUICtrlSetData($status_input, "Exporting TestRail " & $id & ".html ... ")

			Local $case_arr = _TestRailGetCaseIDTitleObjectivesPreconditionsNotesSteps($id)
;			_ArrayDisplay($case_arr)

			$case_arr[1][0] = _TestRailMarkdownToHTML($case_arr[1][0])
			$case_arr[2][0] = _TestRailMarkdownToHTML($case_arr[2][0])
			$case_arr[3][0] = _TestRailMarkdownToHTML($case_arr[3][0])
			$case_arr[4][0] = _TestRailMarkdownToHTML($case_arr[4][0])

			$html = 			""

			$html = $html &		"<!DOCTYPE html>" & @CRLF & _
								"<html>" & @CRLF & _
								"<head>" & @CRLF & _
								"<style>" & @CRLF & _
								"table, th, td {border: 1px solid black; border-collapse: collapse; font-size: 12px; font-family: Arial;}" & @CRLF & _
								"table {table-layout: fixed;}" & @CRLF & _
								"table.b {width: 2050px;}" & @CRLF & _
								".ds {min-width: 400px; text-align: left;}" & @CRLF & _
								".tes {min-width: 800px; text-align: left;}" & @CRLF & _
								".mti {min-width: 110px; text-align: center;}" & @CRLF & _
								".tt {width: 500px; text-align: left;}" & @CRLF & _
								".ttbc {width: 500px; text-align: left; font-weight: bold; text-align: center;}" & @CRLF & _
								".ati {min-width: 150px; text-align: center;}" & @CRLF & _
								".sd {min-width: 1000px; text-align: left;}" & @CRLF & _
								".tc {width: 50px; text-align: left;}" & @CRLF & _
								".tcbc {width: 50px; text-align: left; font-weight: bold; text-align: center;}" & @CRLF & _
								".tr {width: 150px; text-align: left;}" & @CRLF & _
								".trb {width: 150px; text-align: left; font-weight: bold;}" & @CRLF & _
								".ng {text-align: left; background-color: #008080; color: white;}" & @CRLF & _
								".ngc {text-align: left; background-color: #008080; color: white; text-align: center;}" & @CRLF & _
								".ts {width: 100px; text-align: left;}" & @CRLF & _
								".tsb {width: 100px; text-align: left; font-weight: bold;}" & @CRLF & _
								".mp {background-color: yellow;}" & @CRLF & _
								".rh {background-color: seagreen; color: white;}" & @CRLF & _
								".rhr {background-color: seagreen; color: white; text-align:center; white-space:nowrap; transform-origin:50% 50%; transform: rotate(-90deg);}" & @CRLF & _
								".rhr:before {background-color: seagreen; color: white; content:''; padding-top:100%; display:inline-block; vertical-align:middle;}" & @CRLF & _
								".i {background-color: deepskyblue;}" & @CRLF & _
								"</style>" & @CRLF & _
								"</head>" & @CRLF & _
								"<body>" & @CRLF & _
								"<table>" & @CRLF & _
								"<tr class=" & $double_quotes & "ng" & $double_quotes & "><td class=" & $double_quotes & "trb" & $double_quotes & ">Document Name:</td><td colspan=" & $double_quotes & "5" & $double_quotes & "></td></tr>" & @CRLF & _
								"<tr><td class=" & $double_quotes & "trb" & $double_quotes & ">Product Owner:</td><td class=" & $double_quotes & "tr" & $double_quotes & "></td><td class=" & $double_quotes & "tsb" & $double_quotes & ">Release Date:</td><td colspan=" & $double_quotes & "3" & $double_quotes & "></td></tr>" & @CRLF & _
								"<tr><td class=" & $double_quotes & "trb" & $double_quotes & ">Developer:</td><td class=" & $double_quotes & "tr" & $double_quotes & "></td><td class=" & $double_quotes & "tsb" & $double_quotes & ">Unit(s):</td><td class=" & $double_quotes & "ts" & $double_quotes & "></td><td class=" & $double_quotes & "trb" & $double_quotes & ">Version / Build:</td><td></td></tr>" & @CRLF & _
								"<tr><td class=" & $double_quotes & "trb" & $double_quotes & ">Test Objective:</td><td colspan=" & $double_quotes & "5" & $double_quotes & ">" & $case_arr[2][0] & "</td></tr>" & @CRLF & _
								"<tr><td class=" & $double_quotes & "trb" & $double_quotes & ">Prepared By / Date:</td><td colspan=" & $double_quotes & "5" & $double_quotes & "></td></tr>" & @CRLF & _
								"<tr><td class=" & $double_quotes & "trb" & $double_quotes & ">Tested By / Test Date:</td><td colspan=" & $double_quotes & "5" & $double_quotes & "></td></tr>" & @CRLF & _
								"<tr><td class=" & $double_quotes & "trb" & $double_quotes & ">Description:</td><td colspan=" & $double_quotes & "5" & $double_quotes & ">" & $case_arr[1][0] & "</td></tr>" & @CRLF & _
								"<tr><td class=" & $double_quotes & "trb" & $double_quotes & ">Prerequisite:</td><td colspan=" & $double_quotes & "5" & $double_quotes & ">" & $case_arr[3][0] & "</td></tr>" & @CRLF & _
								"<tr><td class=" & $double_quotes & "trb" & $double_quotes & ">Role:</td><td colspan=" & $double_quotes & "5" & $double_quotes & ">" & $case_arr[4][0] & "</td></tr></table>" & @CRLF & _
								"<br>" & @CRLF & _
								"<table class=" & $double_quotes & "b" & $double_quotes & ">" & @CRLF & _
								"<tr class=" & $double_quotes & "ng" & $double_quotes & "><td class=" & $double_quotes & "tcbc" & $double_quotes & ">Step #</td><td class=" & $double_quotes & "ttbc" & $double_quotes & ">Steps</td><td class=" & $double_quotes & "ttbc" & $double_quotes & ">Expected Result</td><td class=" & $double_quotes & "ttbc" & $double_quotes & ">Actual Result</td><td class=" & $double_quotes & "ttbc" & $double_quotes & ">Remarks</td></tr>" & @CRLF




			Local $step_num = 0

			for $i = 5 to (UBound($case_arr) - 1)

				$step_num = $step_num + 1
				$case_arr[$i][0] = _TestRailMarkdownToHTML($case_arr[$i][0])
				$case_arr[$i][1] = _TestRailMarkdownToHTML($case_arr[$i][1])
				$html = $html &		"<tr><td class=" & $double_quotes & "tcbc" & $double_quotes & ">" & $step_num & "</td><td>" & $case_arr[$i][0] & "</td><td>" & $case_arr[$i][1] & "</td><td></td><td></td></tr>" & @CRLF
			Next

			$html = $html &			"</table></body></html>" & @CRLF
			FileDelete($report_filename)
			FileWrite($report_filename, $html)
			GUICtrlSetData($status_input, "")

		Next

		GUICtrlSetData($progress, 0)
		enable_gui()
	EndIf


EndFunc


Func Update_Confluence_Page($url, $jira_username, $jira_password, $space_key, $ancestor_key, $page_key, $page_title, $page_body)

	_ConfluenceSetup()
	_ConfluenceDomainSet($url)
	_ConfluenceLogin($jira_username, $jira_password)
	_ConfluenceUpdatePage($space_key, $ancestor_key, $page_key, $page_title, $page_body)
	_ConfluenceShutdown()

EndFunc

Func disable_gui()

	GUISetCursor(15, 1, $main_gui)
	GUICtrlSetState($testrail_username_input, $GUI_DISABLE)
	GUICtrlSetState($testrail_password_input, $GUI_DISABLE)
	GUICtrlSetState($testrail_authenticate_button, $GUI_DISABLE)
	GUICtrlSetState($testrail_project_listview, $GUI_DISABLE)
	GUICtrlSetState($testrail_section_listview, $GUI_DISABLE)
	GUICtrlSetState($testrail_case_listview, $GUI_DISABLE)
	GUICtrlSetState($export_button, $GUI_DISABLE)
EndFunc

Func enable_gui()

	GUISetCursor(2, 0, $main_gui)
	GUICtrlSetState($testrail_username_input, $GUI_ENABLE)
	GUICtrlSetState($testrail_password_input, $GUI_ENABLE)
	GUICtrlSetState($testrail_authenticate_button, $GUI_ENABLE)
	GUICtrlSetState($testrail_project_listview, $GUI_ENABLE)
	GUICtrlSetState($testrail_section_listview, $GUI_ENABLE)
	GUICtrlSetState($testrail_case_listview, $GUI_ENABLE)
	GUICtrlSetState($export_button, $GUI_ENABLE)
EndFunc
