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
#include <Excel.au3>
#include <Word.au3>
#include <MsgBoxConstants.au3>



Global $html, $markup, $storage_format
Global $app_name = "TestRail Test Cases to HTML"
Global $ini_filename = @ScriptDir & "\" & $app_name & ".ini"
Global $log_filename = @ScriptDir & "\" & $app_name & ".log"
Global $aResult, $iRows, $iColumns, $iRval, $run_name = "", $max_num_defects = 0, $max_num_days = 0, $version_name = ""
Global $user_arr

Global $main_gui = GUICreate($app_name, 860, 600)

GUICtrlCreateGroup("TestRail Login", 10, 10, 360, 180)
GUICtrlCreateLabel("Username", 20, 30, 60, 20)
Global $testrail_username_input = GUICtrlCreateInput(IniRead($ini_filename, "main", "testrailusername", "sgriffin@janison.com"), 100, 30, 250, 20)
GUICtrlCreateLabel("Password", 20, 50, 60, 20)
Global $testrail_password_input = GUICtrlCreateInput("", 100, 50, 250, 20, $ES_PASSWORD)
Global $testrail_authenticate_button = GUICtrlCreateButton("Authenticate", 100, 70, 80, 20)
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

FileDelete($log_filename)

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

			; the following is required for _TestRailGetAttachment()
			_TestRailAuth()

			if StringLen(GUICtrlRead($testrail_password_input)) > 0 Then

				_GUICtrlListView_DeleteAllItems($testrail_project_listview)
				_GUICtrlListView_DeleteAllItems($testrail_section_listview)
				_GUICtrlListView_DeleteAllItems($testrail_case_listview)

				GUICtrlSetData($status_input, "Getting the TestRail Users ... ")
				_FileWriteLog($log_filename, "Getting the TestRail Users ... ")
				$user_arr = _TestRailGetUsersIDName()

				GUICtrlSetData($status_input, "Getting the TestRail Projects ... ")
				_FileWriteLog($log_filename, "Getting the TestRail Projects ... ")
				Local $project_arr = _TestRailGetProjectsIDAndNameArray()
				_GUICtrlListView_AddArray($testrail_project_listview, $project_arr)

				GUICtrlSetData($status_input, "")
			EndIf

			enable_gui()

		Case $export_button

			Create_Case_HTML()

	EndSwitch

WEnd

GUIDelete($main_gui)

; Shutdown Jira

GUICtrlSetData($status_input, "Closing TestRail ... ")
_TestRailShutdown()


Func WM_NOTIFY($hWnd, $iMsg, $wParam, $lParam)
    #forceref $hWnd, $iMsg, $wParam
    Local $hWndFrom, $iIDFrom, $iCode, $tNMHDR, $hWndListView, $tInfo

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

	Local $selected_case_index = _GUICtrlListView_GetSelectedIndices($testrail_case_listview, True)

	Local $case_id_arr[$selected_case_index[0]]

	for $i = 1 to $selected_case_index[0]

		$case_id_arr[$i - 1] = _GUICtrlListView_GetItemText($testrail_case_listview, $selected_case_index[$i], 0)
	Next

	Local $selected_project_id = _GUICtrlListView_GetItemText($testrail_project_listview, Number(_GUICtrlListView_GetSelectedIndices($testrail_project_listview)), 0)
	Local $selected_section_id = _GUICtrlListView_GetItemText($testrail_section_listview, Number(_GUICtrlListView_GetSelectedIndices($testrail_section_listview)), 0)

	GUICtrlSetData($status_input, "Getting the TestRail Cases ... ")
	_FileWriteLog($log_filename, "Getting the TestRail Cases ... ")
	Local $cases_arr = _TestRailGetCasesIDTitleTestScenarioIDOwnerObjectivesPreconditionsNotesSteps($selected_project_id, "", $selected_section_id, $case_id_arr, $user_arr)

;	_ArrayDisplay($cases_arr)

	Local $num_cases = 0

	for $i = 0 to (UBound($cases_arr) - 1)
		ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $i = ' & $i & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console

		if StringCompare($cases_arr[$i][0], "Test Case Start") = 0 Then

			$num_cases = $num_cases + 1
		EndIf
	Next

	disable_gui()

	for $i = 1 to (UBound($cases_arr) - 1)

		Local $id = $cases_arr[$i][0]
		_FileWriteLog($log_filename, "case id = " & $id)
		$i = $i + 1
		Local $title = $cases_arr[$i][0]
		_FileWriteLog($log_filename, "case title = " & $title)
		$i = $i + 1
		Local $test_scenario_id = $cases_arr[$i][0]
		_FileWriteLog($log_filename, "case test scenario id = " & $test_scenario_id)
		$i = $i + 1

		if StringLen($test_scenario_id) = 0 Then

			_FileWriteLog($log_filename, "skipping this test case as the test scenario is blank")

			While StringCompare($cases_arr[$i][0], "Test Case Start") <> 0 and $i < (UBound($cases_arr) - 1)

				$i = $i + 1
			WEnd
		Else

			Local $owner_name = $cases_arr[$i][0]
			_FileWriteLog($log_filename, "case owner name = " & $owner_name)
			$i = $i + 1
			ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $cases_arr[$i][0] = ' & $cases_arr[$i][0] & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
			Local $objectives = _TestRailMarkdownToHTML($cases_arr[$i][0])
			_FileWriteLog($log_filename, "case objectives = " & $objectives)
			$i = $i + 1
			Local $preconditions = _TestRailMarkdownToHTML($cases_arr[$i][0])
			_FileWriteLog($log_filename, "case preconditions = " & $preconditions)
			$i = $i + 1
			Local $notes = _TestRailMarkdownToHTML($cases_arr[$i][0])

			; Special case - removing "Role :" if the text appears at the beginning
			$notes = StringRegExpReplace($notes, "^Role: ", "")

			_FileWriteLog($log_filename, "case notes = " & $notes)
			$i = $i + 1

			Local $report_filename = $id & ".html"
			GUICtrlSetData($status_input, "Exporting " & $report_filename & " ... ")
			_FileWriteLog($log_filename, "Exporting " & $report_filename & " ... ")
			$report_filename = $test_scenario_id & ".html"

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
								"<tr class=" & $double_quotes & "ng" & $double_quotes & "><td class=" & $double_quotes & "trb" & $double_quotes & ">Document Name:</td><td colspan=" & $double_quotes & "5" & $double_quotes & ">" & $test_scenario_id & "</td></tr>" & @CRLF & _
								"<tr><td class=" & $double_quotes & "trb" & $double_quotes & ">Product Owner:</td><td class=" & $double_quotes & "tr" & $double_quotes & ">" & $owner_name & "</td><td class=" & $double_quotes & "tsb" & $double_quotes & ">Release Date:</td><td colspan=" & $double_quotes & "3" & $double_quotes & "></td></tr>" & @CRLF & _
								"<tr><td class=" & $double_quotes & "trb" & $double_quotes & ">Developer:</td><td class=" & $double_quotes & "tr" & $double_quotes & "></td><td class=" & $double_quotes & "tsb" & $double_quotes & ">Unit(s):</td><td class=" & $double_quotes & "ts" & $double_quotes & "></td><td class=" & $double_quotes & "trb" & $double_quotes & ">Version / Build:</td><td></td></tr>" & @CRLF & _
								"<tr><td class=" & $double_quotes & "trb" & $double_quotes & ">Test Objective:</td><td colspan=" & $double_quotes & "5" & $double_quotes & ">" & $objectives & "</td></tr>" & @CRLF & _
								"<tr><td class=" & $double_quotes & "trb" & $double_quotes & ">Prepared By / Date:</td><td colspan=" & $double_quotes & "5" & $double_quotes & "></td></tr>" & @CRLF & _
								"<tr><td class=" & $double_quotes & "trb" & $double_quotes & ">Tested By / Test Date:</td><td colspan=" & $double_quotes & "5" & $double_quotes & "></td></tr>" & @CRLF & _
								"<tr><td class=" & $double_quotes & "trb" & $double_quotes & ">Description:</td><td colspan=" & $double_quotes & "5" & $double_quotes & ">" & $title & "</td></tr>" & @CRLF & _
								"<tr><td class=" & $double_quotes & "trb" & $double_quotes & ">Prerequisite:</td><td colspan=" & $double_quotes & "5" & $double_quotes & ">" & $preconditions & "</td></tr>" & @CRLF & _
								"<tr><td class=" & $double_quotes & "trb" & $double_quotes & ">Role:</td><td colspan=" & $double_quotes & "5" & $double_quotes & ">" & $notes & "</td></tr></table>" & @CRLF & _
								"<br>" & @CRLF & _
								"<table class=" & $double_quotes & "b" & $double_quotes & ">" & @CRLF & _
								"<tr class=" & $double_quotes & "ng" & $double_quotes & "><td class=" & $double_quotes & "tcbc" & $double_quotes & ">Step #</td><td class=" & $double_quotes & "ttbc" & $double_quotes & ">Steps</td><td class=" & $double_quotes & "ttbc" & $double_quotes & ">Expected Result</td><td class=" & $double_quotes & "ttbc" & $double_quotes & ">Actual Result</td><td class=" & $double_quotes & "ttbc" & $double_quotes & ">Remarks</td></tr>" & @CRLF

			Local $step_num = 0

			While $i < UBound($cases_arr) and StringCompare($cases_arr[$i][0], "Test Case Start") <> 0

				$step_num = $step_num + 1
				$cases_arr[$i][0] = _TestRailMarkdownToHTML($cases_arr[$i][0])
				_FileWriteLog($log_filename, "step " & $step_num & " steps = " & $cases_arr[$i][0])
				$cases_arr[$i][1] = _TestRailMarkdownToHTML($cases_arr[$i][1])
				_FileWriteLog($log_filename, "step " & $step_num & " expected result = " & $cases_arr[$i][1])
				$html = $html &		"<tr><td class=" & $double_quotes & "tcbc" & $double_quotes & ">" & $step_num & "</td><td>" & $cases_arr[$i][0] & "</td><td>" & $cases_arr[$i][1] & "</td><td></td><td></td></tr>" & @CRLF
				$i = $i + 1
			WEnd

			$html = $html &			"</table></body></html>" & @CRLF

			_FileWriteLog($log_filename, "creating file " & @ScriptDir & "\" & $report_filename)
			FileDelete(@ScriptDir & "\" & $report_filename)
			FileWrite(@ScriptDir & "\" & $report_filename, $html)

			_FileWriteLog($log_filename, "creating excel file for " & @ScriptDir & "\" & $report_filename)
			html_file_to_excel_file(@ScriptDir & "\" & $report_filename, "", "20,50,50,30,30")
			html_file_to_word_file(@ScriptDir & "\" & $report_filename, "")

		EndIf


		GUICtrlSetData($status_input, "")

	Next

	GUICtrlSetData($progress, 0)
	enable_gui()


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


Func html_file_to_excel_file($html_file_path, $excel_file_path = "", $comma_separated_column_widths = "")

	if StringLen($excel_file_path) = 0 Then

		$excel_file_path = StringReplace($html_file_path, ".html", ".xlsx")
	EndIf

	FileDelete($excel_file_path)

	; Create application object
	Local $oExcel = _Excel_Open(False, False, False, False, False)
	If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_BookOpen Example", "Error creating the Excel application object." & @CRLF & "@error = " & @error & ", @extended = " & @extended)

	; Open an existing workbook and return its object identifier.
	Local $oWorkbook = _Excel_BookOpen($oExcel, $html_file_path) ;, False, False)
	If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_BookOpen Example 1", "Error opening '" & $html_file_path & "'." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
	;MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_BookOpen Example 1", "Workbook '" & $sWorkbook & "' has been opened successfully." & @CRLF & @CRLF & "Creation Date: " & $oWorkbook.BuiltinDocumentProperties("Creation Date").Value)

	Local $sResult = _Excel_RangeRead($oWorkbook, Default, "A1")
	If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_RangeRead Example 1", "Error reading from workbook." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
	;MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_RangeRead Example 1", "Data successfully read." & @CRLF & "Value of cell A1: " & $sResult)
	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $sResult = ' & $sResult & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console

	if StringLen($comma_separated_column_widths) > 0 Then

		Local $column_width = StringSplit($comma_separated_column_widths, ",", 3)

		for $i = 0 to (UBound($column_width) - 1)

			Local $range_str = _Excel_ColumnToLetter($i + 1) & ":" & _Excel_ColumnToLetter($i + 1)
			$oWorkbook.Sheets(1).Range($range_str).ColumnWidth = $column_width[$i]
		Next
	EndIf

	$oWorkbook.Sheets(1).Range("1:100").Rows.AutoFit


	_Excel_BookSaveAs($oWorkbook, $excel_file_path, $xlWorkbookDefault, True)
	If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_BookSaveAs Example 1", "Error saving workbook to '" & $excel_file_path & "'." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
	;MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_BookSaveAs Example 1", "Workbook successfully saved as '" & $excel_workbook & "'.")

	_Excel_Close($oExcel)
	If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_Close Example 1", "Error closing the Excel application." & @CRLF & "@error = " & @error & ", @extended = " & @extended)

EndFunc


Func html_file_to_word_file($html_file_path, $word_file_path = "")

	if StringLen($word_file_path) = 0 Then

		$word_file_path = StringReplace($html_file_path, ".html", ".docx")
	EndIf

	FileDelete($word_file_path)


	; Create application object
	Local $oWord = _Word_Create(False)
	If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Word UDF: _Word_DocOpen Example", "Error creating a new Word application object." & @CRLF & "@error = " & @error & ", @extended = " & @extended)

	; Open a document read-only
	Local $oDoc = _Word_DocOpen($oWord, $html_file_path)
	If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Word UDF: _Word_DocOpen Example 1", "Error opening '.\Extras\Test.doc'." & @CRLF & "@error = " & @error & ", @extended = " & @extended)


	; Save document
	_Word_DocSaveAs($oDoc, $word_file_path, $WdFormatDocumentDefault)
	If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Word UDF: _Word_DocSaveAs Example", "Error saving the Word document." & @CRLF & "@error = " & @error & ", @extended = " & @extended)

;	_Word_DocSaveAsEx($oDoc, $word_file_path, $WdFormatDocumentDefault)
;	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $word_file_path = ' & $word_file_path & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console

	; Close a Word coument
	_Word_DocClose($oDoc)
	If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Word UDF: _Word_DocClose Example", "Error closing document '.\Extras\Test.doc'." & @CRLF & "@error = " & @error & ", @extended = " & @extended)

	_Word_Quit($oWord)
	If @error Then MsgBox($MB_SYSTEMMODAL, "Word UDF: _Word_Quit Example", "Error closing the Word application object." & @CRLF & "@error = " & @error & ", @extended = " & @extended)


	; make paths relative not absolute

	DirRemove(@ScriptDir & "\word", 1)
	ShellExecuteWait("7z.exe", "x """ & $word_file_path & """ word\document.xml", @ScriptDir, "", @SW_HIDE)
	Local $str = FileRead(@ScriptDir & "\word\document.xml")
	$str = StringReplace($str, @ScriptDir & "\", "")
	FileDelete(@ScriptDir & "\word\document.xml")
	FileWrite(@ScriptDir & "\word\document.xml", $str)
	ShellExecuteWait("7z.exe", "u """ & $word_file_path & """ word\document.xml", @ScriptDir, "", @SW_HIDE)
	DirRemove(@ScriptDir & "\word", 1)

EndFunc

; #FUNCTION# ====================================================================================================================
; Author ........: water (based on the Word UDF written by Bob Anthony)
; Modified ......:
; ===============================================================================================================================
Func _Word_DocSaveAsEx($oDoc, $sFileName = Default, $iFileFormat = Default, $bReadOnlyRecommended = Default, $bAddToRecentFiles = Default, $sPassword = Default, $sWritePassword = Default)
    ; Error handler, automatic cleanup at end of function
    Local $oError = ObjEvent("AutoIt.Error", "__Word_COMErrFuncEx")
    #forceref $oError

    If $bReadOnlyRecommended = Default Then $bReadOnlyRecommended = False
    If $bAddToRecentFiles = Default Then $bAddToRecentFiles = 0
    If $sPassword = Default Then $sPassword = ""
    If $sWritePassword = Default Then $sWritePassword = ""
    If Not IsObj($oDoc) Then Return SetError(1, 0, 0)
    $oDoc.SaveAs2($sFileName, $iFileFormat, False, $sPassword, $bAddToRecentFiles, $sWritePassword, $bReadOnlyRecommended) ; Try to save for >= Word 2010
    If @error = 0x80020006 Then $oDoc.SaveAs($sFileName, $iFileFormat, False, $sPassword, $bAddToRecentFiles, $sWritePassword, $bReadOnlyRecommended) ; COM error "Unknown Name" hence save for <= Word 2007
    If @error Then Return SetError(2, @error, 0)
    Return 1
EndFunc   ;==>_Word_DocSaveAsEx

; #INTERNAL_USE_ONLY#============================================================================================================
; Name...........: __Word_COMErrFunc
; Description ...: Dummy function for silently handling COM errors.
; Syntax.........:
; Parameters ....:
; Return values .:
;
; Author ........:
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......:
; ===============================================================================================================================
Func __Word_COMErrFuncEx($oError)
    ; Do anything here.
    ConsoleWrite(@ScriptName & " (" & $oError.scriptline & ") : ==> COM Error intercepted !" & @CRLF & _
            @TAB & "err.number is: " & @TAB & @TAB & "0x" & Hex($oError.number) & @CRLF & _
            @TAB & "err.windescription:" & @TAB & $oError.windescription & @CRLF & _
            @TAB & "err.description is: " & @TAB & $oError.description & @CRLF & _
            @TAB & "err.source is: " & @TAB & @TAB & $oError.source & @CRLF & _
            @TAB & "err.helpfile is: " & @TAB & $oError.helpfile & @CRLF & _
            @TAB & "err.helpcontext is: " & @TAB & $oError.helpcontext & @CRLF & _
            @TAB & "err.lastdllerror is: " & @TAB & $oError.lastdllerror & @CRLF & _
            @TAB & "err.scriptline is: " & @TAB & $oError.scriptline & @CRLF & _
            @TAB & "err.retcode is: " & @TAB & "0x" & Hex($oError.retcode) & @CRLF & @CRLF)
EndFunc   ;==>__Word_COMErrFuncEx
