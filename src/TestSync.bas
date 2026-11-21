Attribute VB_Name = "TestSync"
Option Explicit

' Tests for GCV Sync. Run RunSyncTests from the VBA editor (F5) in the DEV
' copy. Uses FakeSyncTarget (no network) and a temporary workbook closed
' without saving. Do NOT ship this module or FakeSyncTarget.

Private passedCount As Long
Private failedCount As Long
Private currentTest As String

Public Sub RunSyncTests()
    Dim wb As Workbook

    passedCount = 0
    failedCount = 0
    Debug.Print String$(60, "-")
    Debug.Print "Sync tests"
    On Error GoTo Crash

    Test_JsonString_EscapesAndKeepsArabic
    Test_JsonValue_NumbersDatesBooleansErrors
    Test_ParseAnswer_ReadsSuccessRefusalAndErrors
    Test_Codes_And_Urls
    Test_SchoolNumberTyped_BecomesGcvCode
    Test_DeviceKey_IsRandomHex

    Set wb = TestWorkbook()
    Test_SheetJson_KeepsTextIdsDatesAndTrimsEmptyCells wb
    Test_SheetsToSync_LeavesOutVeryHiddenSheets wb
    Test_RunSync_RefusesNonGcvWorkbooks wb
    Test_RunSync_SendsSchoolKeyNameAndSheets wb
    Test_RunSync_ReportsOfflineAndRefusal wb

Finish:
    Application.StatusBar = False
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Debug.Print "Passed: " & passedCount & "   Failed: " & failedCount
    MsgBox "Passed: " & passedCount & vbCrLf & "Failed: " & failedCount & vbCrLf & _
           "Details in the Immediate window (Ctrl+G).", _
           IIf(failedCount = 0, vbInformation, vbCritical), "Sync tests"
    Exit Sub

Crash:
    failedCount = failedCount + 1
    Debug.Print "CRASH in " & currentTest & ": " & Err.Number & " " & Err.Description
    Resume Finish
End Sub

' ----------------------------------------------------------------------------
' JSON and answers
' ----------------------------------------------------------------------------

Private Sub Test_JsonString_EscapesAndKeepsArabic()
    StartTest "JSON text: quotes, backslashes and line breaks escaped, Arabic kept"
    AssertEqual """a\""b\\c\nd""", JsonString("a""b\c" & vbLf & "d"), "escapes"
    AssertEqual """" & FromCodePoints("0645 062D 0645 062F") & """", JsonString(FromCodePoints("0645 062D 0645 062F")), "Arabic"
    AssertEqual """""", JsonString(""), "empty"
    AssertEqual """\u0001""", JsonString(Chr$(1)), "control character"
End Sub

Private Sub Test_JsonValue_NumbersDatesBooleansErrors()
    StartTest "JSON values"
    AssertEqual "42", JsonValue(42), "integer"
    AssertEqual "0.5", JsonValue(0.5), "decimal point, leading zero"
    AssertEqual "-0.25", JsonValue(-0.25), "negative"
    AssertEqual "true", JsonValue(True), "boolean"
    AssertEqual "null", JsonValue(Empty), "empty"
    AssertEqual "{""$d"": ""2015-03-15""}", JsonValue(DateSerial(2015, 3, 15)), "date"
    AssertEqual "{""$d"": ""2026-11-19T10:30:00""}", JsonValue(DateSerial(2026, 11, 19) + TimeSerial(10, 30, 0)), "date and time"
    AssertEqual """#N/A""", JsonValue(CVErr(xlErrNA)), "error"
    AssertEqual """0123""", JsonValue("0123"), "text that looks like a number"
End Sub

Private Sub Test_ParseAnswer_ReadsSuccessRefusalAndErrors()
    Dim r As SyncResult
    StartTest "Answers of the web app"
    Set r = ParseAnswer(200, "{""ok"": true, ""sheets"": 8, ""rows"": 420, ""url"": ""https://docs.google.com/s""}")
    AssertEqual soOk, r.Outcome, "ok"
    AssertEqual False, r.Registered, "not a first sync"
    AssertEqual 8, r.SheetCount, "sheets"
    AssertEqual 420, r.RowCount, "rows"
    AssertEqual "https://docs.google.com/s", r.SheetUrl, "url"
    Set r = ParseAnswer(200, "{""ok"": false, ""error"": ""Unknown school or wrong key.""}")
    AssertEqual soRejected, r.Outcome, "refused"
    AssertEqual "Unknown school or wrong key.", r.Message, "reason"
    Set r = ParseAnswer(200, "{""ok"": true, ""registered"": true, ""sheets"": 1, ""rows"": 1}")
    AssertEqual True, r.Registered, "first sync links the PC"
    Set r = ParseAnswer(200, "{""ok"": false, ""code"": ""school-taken"", ""error"": ""School GCV05 is linked to another computer.""}")
    AssertEqual "school-taken", r.ErrorCode, "refusal code"
    AssertEqual soHttpError, ParseAnswer(500, "").Outcome, "HTTP error"
    AssertEqual soHttpError, ParseAnswer(200, "<html>sign in</html>").Outcome, "not JSON"
End Sub

Private Sub Test_Codes_And_Urls()
    StartTest "School codes and web-app links"
    AssertEqual True, IsSchoolCode("GCV05"), "code"
    AssertEqual True, IsSchoolCode("GCV21"), "code 21"
    AssertEqual False, IsSchoolCode("GCV00"), "no school 00"
    AssertEqual False, IsSchoolCode("GCV5"), "one digit"
    AssertEqual False, IsSchoolCode("S03"), "old style"
    AssertEqual True, IsWebAppUrl("https://script.google.com/macros/s/AKfy123/exec"), "link"
    AssertEqual False, IsWebAppUrl("https://script.google.com/macros/s/AKfy123/dev"), "test link"
    AssertEqual False, IsWebAppUrl("http://example.org/exec"), "other site"
End Sub

Private Sub Test_SchoolNumberTyped_BecomesGcvCode()
    StartTest "The secretary types the school number"
    AssertEqual "GCV05", SchoolCodeFromInput("5"), "5"
    AssertEqual "GCV05", SchoolCodeFromInput(" 05 "), "05"
    AssertEqual "GCV07", SchoolCodeFromInput("gcv7"), "gcv7"
    AssertEqual "GCV21", SchoolCodeFromInput("GCV21"), "GCV21"
    AssertEqual "", SchoolCodeFromInput("0"), "0"
    AssertEqual "", SchoolCodeFromInput("100"), "100"
    AssertEqual "", SchoolCodeFromInput("5.5"), "5.5"
    AssertEqual "", SchoolCodeFromInput("abc"), "text"
    AssertEqual "", SchoolCodeFromInput(""), "empty"
End Sub

Private Sub Test_DeviceKey_IsRandomHex()
    Dim a As String
    Dim b As String
    StartTest "The PC key is 64 random hex characters"
    a = NewDeviceKey()
    b = NewDeviceKey()
    AssertEqual 64, Len(a), "length"
    AssertEqual True, IsDeviceKey(a), "hex"
    AssertEqual False, (a = b), "different every time"
    AssertEqual False, IsDeviceKey(UCase$(a)), "lower case only"
    AssertEqual False, IsDeviceKey("abc"), "too short"
End Sub

' ----------------------------------------------------------------------------
' Workbook
' ----------------------------------------------------------------------------

Private Sub Test_SheetJson_KeepsTextIdsDatesAndTrimsEmptyCells(ByVal wb As Workbook)
    Dim json As String
    StartTest "Sheet: values in rows, text IDs stay text, empty cells at the ends left out"
    json = SheetJson(wb.Worksheets(FromCodePoints("0627 0644 0637 0644 0627 0628")))
    AssertEqual True, InStr(json, """name"": """ & FromCodePoints("0627 0644 0637 0644 0627 0628") & """") > 0, "name"
    AssertEqual True, InStr(json, """rightToLeft"": true") > 0, "right to left"
    AssertEqual True, InStr(json, "[""" & FromCodePoints("0627 0644 0627 0633 0645") & """, ""ID""]") > 0, "header row"
    AssertEqual True, InStr(json, "[""" & FromCodePoints("0645 062D 0645 062F") & """, ""0123456789"", {""$d"": ""2015-03-15""}]") > 0, "data row"
    AssertEqual True, InStr(json, "[]") > 0, "empty row inside the data"
    AssertEqual 0, InStr(json, "null]"), "no empty cells at row ends"
    AssertEqual True, Right$(json, 3) = "]]}", "no empty rows at the end"
End Sub

Private Sub Test_SheetsToSync_LeavesOutVeryHiddenSheets(ByVal wb As Workbook)
    StartTest "Very hidden sheets (Settings) are not sent"
    AssertEqual 2, SheetsToSync(wb).Count, "two of three sheets"
End Sub

Private Sub Test_RunSync_RefusesNonGcvWorkbooks(ByVal wb As Workbook)
    Dim fake As FakeSyncTarget
    Dim other As Workbook
    StartTest "Only GCV workbooks are synced"
    Set fake = New FakeSyncTarget
    Set other = Workbooks.Add(xlWBATWorksheet)
    AssertEqual soNotGcvWorkbook, RunSync(other, fake, "GCV03", "k").Outcome, "refused"
    AssertEqual 0, fake.Calls, "nothing sent"
    other.Close SaveChanges:=False
    AssertEqual soNotConfigured, RunSync(wb, fake, "", "").Outcome, "no settings"
End Sub

Private Sub Test_RunSync_SendsSchoolKeyNameAndSheets(ByVal wb As Workbook)
    Dim fake As FakeSyncTarget
    Dim payload As Object
    StartTest "The payload names the school, its key, the school name and every sheet"
    Set fake = New FakeSyncTarget
    AssertEqual soOk, RunSync(wb, fake, "GCV03", "secret-key").Outcome, "ok"
    Set payload = JsonParse(fake.LastPayload)
    AssertEqual "GCV03", JsonStr(payload, "school"), "school"
    AssertEqual "secret-key", JsonStr(payload, "key"), "key"
    AssertEqual FromCodePoints("0645 062F 0631 0633 0629 0020 0627 0644 0627 062E 062A 0628 0627 0631"), JsonStr(payload, "schoolName"), "school name"
    AssertEqual 2, JsonObj(payload, "sheets").Count, "sheets"
    AssertEqual GcvAddin.ADDIN_VERSION, JsonStr(payload, "addinVersion"), "version"
End Sub

Private Sub Test_RunSync_ReportsOfflineAndRefusal(ByVal wb As Workbook)
    Dim fake As FakeSyncTarget
    StartTest "Offline and refused syncs are reported"
    Set fake = New FakeSyncTarget
    fake.Offline = True
    AssertEqual soOffline, RunSync(wb, fake, "GCV03", "k").Outcome, "offline"
    fake.Offline = False
    fake.AnswerText = "{""ok"": false, ""error"": ""This file belongs to another school.""}"
    AssertEqual soRejected, RunSync(wb, fake, "GCV03", "k").Outcome, "refused"
End Sub

' ----------------------------------------------------------------------------
' Helpers
' ----------------------------------------------------------------------------

' GCV workbook: students sheet (RTL, text ID, date, empty row, empty cells at
' the end), a summary sheet with SCHOOL_NAME, a very hidden Settings sheet.
Private Function TestWorkbook() As Workbook
    Dim wb As Workbook
    Dim ws As Worksheet
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    ws.Name = FromCodePoints("0627 0644 0637 0644 0627 0628")
    ws.DisplayRightToLeft = True
    ws.Range("B2").NumberFormat = "@"
    ws.Range("A1:B1").Value = Array(FromCodePoints("0627 0644 0627 0633 0645"), "ID")
    ws.Range("A2").Value = FromCodePoints("0645 062D 0645 062F")
    ws.Range("B2").Value = "0123456789"
    ws.Range("C2").Value = DateSerial(2015, 3, 15)
    ws.Range("A4").Value = FromCodePoints("0623 062D 0645 062F")
    ws.Range("F10").Interior.Color = vbYellow                      ' formatting only: UsedRange grows
    Set ws = wb.Worksheets.Add(After:=ws)
    ws.Name = "Summary"
    ws.Range("A1").Value = FromCodePoints("0645 062F 0631 0633 0629 0020 0627 0644 0627 062E 062A 0628 0627 0631")
    wb.Names.Add Name:="SCHOOL_NAME", RefersTo:="=Summary!$A$1"
    wb.Names.Add Name:="GCV_TEMPLATE", RefersTo:="=""GCV Administration 1.1""", Visible:=False
    Set ws = wb.Worksheets.Add(After:=ws)
    ws.Name = "Settings"
    ws.Range("A1").Value = "secret setting"
    ws.Visible = xlSheetVeryHidden
    Set TestWorkbook = wb
End Function

Private Sub StartTest(ByVal name As String)
    currentTest = name
End Sub

Private Sub AssertEqual(ByVal expected As Variant, ByVal actual As Variant, ByVal label As String)
    Dim same As Boolean
    If VarType(expected) = vbBoolean Or VarType(actual) = vbBoolean Then
        same = (CStr(expected) = CStr(actual))
    ElseIf IsNumeric(expected) And IsNumeric(actual) And VarType(expected) <> vbString And VarType(actual) <> vbString Then
        same = (CDbl(expected) = CDbl(actual))
    Else
        same = (CStr(expected) = CStr(actual))
    End If
    If same Then
        passedCount = passedCount + 1
    Else
        failedCount = failedCount + 1
        Debug.Print "FAIL: " & currentTest & " / " & label & " - expected [" & CStr(expected) & "] got [" & CStr(actual) & "]"
    End If
End Sub

Private Function FromCodePoints(ByVal codePoints As String) As String
    Dim parts() As String
    Dim i As Long
    parts = Split(codePoints, " ")
    For i = LBound(parts) To UBound(parts)
        FromCodePoints = FromCodePoints & ChrW$(CLng("&H" & parts(i)))
    Next i
End Function
