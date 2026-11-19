Attribute VB_Name = "SyncCore"
Option Explicit

' ============================================================================
' GCV Sync - building the request (no user interface; tested).
'
' Payload (UTF-8 JSON):
'   {"school": "S03", "key": "...", "schoolName": "<SCHOOL_NAME of the workbook>",
'    "workbook": "file.xlsx", "addinVersion": "0.1.0",
'    "sheets": [{"name": "...", "rightToLeft": true, "rows": [[...], ...]}, ...]}
' Values only (formula results). Dates as {"$d": "yyyy-mm-dd[Thh:nn:ss]"},
' errors as their text (#N/A). Empty cells at the end of rows and empty rows
' at the end of a sheet are left out. Very hidden sheets are not sent.
' ============================================================================

Public Const SETTINGS_APP As String = "GCV Tools"
Public Const SETTINGS_SECTION As String = "Sync"
Private Const TEMPLATE_NAME As String = "GCV_TEMPLATE"
Private Const SCHOOL_NAME_NAME As String = "SCHOOL_NAME"

' ----------------------------------------------------------------------------
' One sync
' ----------------------------------------------------------------------------

Public Function RunSync(ByVal wb As Workbook, ByVal target As ISyncTarget, _
                        ByVal schoolCode As String, ByVal key As String) As SyncResult
    Dim result As SyncResult

    If Not IsGcvWorkbook(wb) Then
        Set result = New SyncResult
        result.Outcome = soNotGcvWorkbook
        Set RunSync = result
        Exit Function
    End If
    If Len(schoolCode) = 0 Or Len(key) = 0 Then
        Set result = New SyncResult
        result.Outcome = soNotConfigured
        Set RunSync = result
        Exit Function
    End If
    Set RunSync = target.Send(BuildPayload(wb, schoolCode, key))
End Function

Public Function BuildPayload(ByVal wb As Workbook, ByVal schoolCode As String, ByVal key As String) As String
    Dim parts As Collection
    Dim ws As Variant
    Dim sheetParts() As String
    Dim i As Long

    Dim sheetsJson As String

    Set parts = SheetsToSync(wb)
    If parts.Count > 0 Then
        ReDim sheetParts(1 To parts.Count)
        For Each ws In parts
            i = i + 1
            Application.StatusBar = "GCV Sync: " & i & " / " & parts.Count
            sheetParts(i) = SheetJson(ws)
        Next ws
        sheetsJson = Join(sheetParts, ", ")
    End If
    BuildPayload = "{""school"": " & JsonString(schoolCode) & _
                   ", ""key"": " & JsonString(key) & _
                   ", ""schoolName"": " & JsonString(SchoolNameOf(wb)) & _
                   ", ""workbook"": " & JsonString(wb.Name) & _
                   ", ""addinVersion"": " & JsonString(GcvAddin.ADDIN_VERSION) & _
                   ", ""sheets"": [" & sheetsJson & "]}"
End Function

' Worksheets in tab order, without very hidden ones (e.g. Settings).
Public Function SheetsToSync(ByVal wb As Workbook) As Collection
    Dim result As Collection
    Dim ws As Worksheet
    Set result = New Collection
    For Each ws In wb.Worksheets
        If ws.Visible <> xlSheetVeryHidden Then result.Add ws
    Next ws
    Set SheetsToSync = result
End Function

Public Function SheetJson(ByVal ws As Worksheet) As String
    Dim lastCell As Range
    Dim values As Variant
    Dim rowParts() As String
    Dim cellParts() As String
    Dim r As Long
    Dim c As Long
    Dim lastRow As Long
    Dim lastCol As Long

    SheetJson = "{""name"": " & JsonString(ws.Name) & ", ""rightToLeft"": " & LCase$(CStr(ws.DisplayRightToLeft)) & _
                ", ""rows"": ["
    If Application.WorksheetFunction.CountA(ws.UsedRange) = 0 Then
        SheetJson = SheetJson & "]}"
        Exit Function
    End If
    Set lastCell = ws.UsedRange.Cells(ws.UsedRange.Rows.Count, ws.UsedRange.Columns.Count)
    If lastCell.Row = 1 And lastCell.Column = 1 Then
        ReDim values(1 To 1, 1 To 1)
        values(1, 1) = ws.Range("A1").Value
    Else
        values = ws.Range(ws.Cells(1, 1), lastCell).Value
    End If

    ' Leave out empty rows at the end.
    lastRow = UBound(values, 1)
    Do While lastRow > 1 And RowIsEmpty(values, lastRow)
        lastRow = lastRow - 1
    Loop
    ReDim rowParts(1 To lastRow)
    For r = 1 To lastRow
        lastCol = UBound(values, 2)
        Do While lastCol > 0
            If Not IsEmpty(values(r, lastCol)) Then Exit Do
            lastCol = lastCol - 1
        Loop
        If lastCol = 0 Then
            rowParts(r) = "[]"
        Else
            ReDim cellParts(1 To lastCol)
            For c = 1 To lastCol
                cellParts(c) = JsonValue(values(r, c))
            Next c
            rowParts(r) = "[" & Join(cellParts, ", ") & "]"
        End If
    Next r
    SheetJson = SheetJson & Join(rowParts, "," & vbLf) & "]}"
End Function

Private Function RowIsEmpty(ByRef values As Variant, ByVal r As Long) As Boolean
    Dim c As Long
    For c = 1 To UBound(values, 2)
        If Not IsEmpty(values(r, c)) Then Exit Function
    Next c
    RowIsEmpty = True
End Function

' ----------------------------------------------------------------------------
' JSON
' ----------------------------------------------------------------------------

Public Function JsonValue(ByVal value As Variant) As String
    Select Case VarType(value)
        Case vbEmpty, vbNull
            JsonValue = "null"
        Case vbBoolean
            JsonValue = IIf(value, "true", "false")
        Case vbDate
            If value = Int(value) Then
                JsonValue = "{""$d"": """ & Format$(value, "yyyy-mm-dd") & """}"
            Else
                JsonValue = "{""$d"": """ & Format$(value, "yyyy-mm-dd") & "T" & Format$(value, "hh:nn:ss") & """}"
            End If
        Case vbError
            JsonValue = JsonString(ErrorText(value))
        Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
            JsonValue = Trim$(Str$(value))                ' Str$ always uses "." as decimal separator
            If Left$(JsonValue, 1) = "." Then JsonValue = "0" & JsonValue
            If Left$(JsonValue, 2) = "-." Then JsonValue = "-0" & Mid$(JsonValue, 2)
        Case Else
            JsonValue = JsonString(CStr(value))
    End Select
End Function

' Escapes quotes, backslashes and control characters; other characters
' (Arabic) are kept and sent as UTF-8.
Public Function JsonString(ByVal text As String) As String
    Dim i As Long
    Dim ch As String
    Dim code As Long
    Dim parts() As String

    If Len(text) = 0 Then
        JsonString = """"""
        Exit Function
    End If
    ReDim parts(1 To Len(text))
    For i = 1 To Len(text)
        ch = Mid$(text, i, 1)
        code = AscW(ch) And &HFFFF&
        Select Case code
            Case 34: parts(i) = "\"""
            Case 92: parts(i) = "\\"
            Case 10: parts(i) = "\n"
            Case 13: parts(i) = "\r"
            Case 9: parts(i) = "\t"
            Case Is < 32: parts(i) = "\u" & Right$("000" & Hex$(code), 4)
            Case Else: parts(i) = ch
        End Select
    Next i
    JsonString = """" & Join(parts, vbNullString) & """"
End Function

Public Function ErrorText(ByVal value As Variant) As String
    Select Case CLng(value)
        Case 2000: ErrorText = "#NULL!"
        Case 2007: ErrorText = "#DIV/0!"
        Case 2015: ErrorText = "#VALUE!"
        Case 2023: ErrorText = "#REF!"
        Case 2029: ErrorText = "#NAME?"
        Case 2036: ErrorText = "#NUM!"
        Case 2042: ErrorText = "#N/A"
        Case Else: ErrorText = "#ERROR"
    End Select
End Function

Public Function Utf8Bytes(ByVal text As String) As Byte()
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2                                     ' text
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText text
    stream.Position = 0
    stream.Type = 1                                     ' binary
    stream.Position = 3                                 ' skip the byte order mark
    Utf8Bytes = stream.Read
    stream.Close
End Function

' Answer of the web app -> SyncResult.
Public Function ParseAnswer(ByVal status As Long, ByVal text As String) As SyncResult
    Dim result As SyncResult
    Dim answer As Variant

    Set result = New SyncResult
    result.HttpStatus = status
    If status <> 200 Then
        result.Outcome = soHttpError
        result.Message = "HTTP " & status
        Set ParseAnswer = result
        Exit Function
    End If
    On Error Resume Next
    JsonParseInto text, answer
    If Err.Number <> 0 Or Not IsObject(answer) Then
        On Error GoTo 0
        result.Outcome = soHttpError
        result.Message = "Unexpected answer from the sync service."
        Set ParseAnswer = result
        Exit Function
    End If
    On Error GoTo 0
    If JsonBool(answer, "ok") Then
        result.Outcome = soOk
        result.SheetUrl = JsonStr(answer, "url")
        result.SheetCount = Val(JsonStr(answer, "sheets"))
        result.RowCount = Val(JsonStr(answer, "rows"))
    Else
        result.Outcome = soRejected
        result.Message = JsonStr(answer, "error")
    End If
    Set ParseAnswer = result
End Function

' ----------------------------------------------------------------------------
' Workbook and settings
' ----------------------------------------------------------------------------

Public Function IsGcvWorkbook(ByVal wb As Workbook) As Boolean
    Dim definedName As Name
    If wb Is Nothing Then Exit Function
    On Error Resume Next
    Set definedName = wb.Names(TEMPLATE_NAME)
    On Error GoTo 0
    IsGcvWorkbook = Not definedName Is Nothing
End Function

' Value of the workbook's SCHOOL_NAME (a cell or a constant), "" when absent.
Public Function SchoolNameOf(ByVal wb As Workbook) As String
    Dim value As Variant
    On Error Resume Next
    value = wb.Names(SCHOOL_NAME_NAME).RefersToRange.Cells(1, 1).Value
    If Err.Number <> 0 Then
        Err.Clear
        value = wb.Worksheets(1).Evaluate(wb.Names(SCHOOL_NAME_NAME).RefersTo)
    End If
    If Err.Number <> 0 Or IsError(value) Then value = vbNullString
    On Error GoTo 0
    SchoolNameOf = Trim$(CStr(value))
End Function

' Per PC (registry: HKCU\Software\VB and VBA Program Settings\GCV Tools\Sync).
Public Function SettingValue(ByVal key As String) As String
    SettingValue = GetSetting(SETTINGS_APP, SETTINGS_SECTION, key, vbNullString)
End Function

Public Sub SaveSettingValue(ByVal key As String, ByVal value As String)
    SaveSetting SETTINGS_APP, SETTINGS_SECTION, key, value
End Sub

Public Function IsSchoolCode(ByVal code As String) As Boolean
    If Len(code) = 0 Or Len(code) > 20 Then Exit Function
    IsSchoolCode = Not (code Like "*[!A-Za-z0-9-]*")
End Function

Public Function IsWebAppUrl(ByVal url As String) As Boolean
    IsWebAppUrl = (url Like "https://script.google.com/macros/s/*/exec")
End Function
