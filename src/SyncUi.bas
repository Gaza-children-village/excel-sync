Attribute VB_Name = "SyncUi"
Option Explicit

' ============================================================================
' GCV Sync add-in: start-up, ribbon group "sync" and messages.
'   Sync now   - sends the active GCV workbook to the school's Google Sheet
'   Settings   - web-app link, school code and key (once per PC)
' ============================================================================

Private Declare PtrSafe Function MessageBoxW Lib "user32" (ByVal hWnd As LongPtr, _
    ByVal lpText As LongPtr, ByVal lpCaption As LongPtr, ByVal uType As Long) As Long

Private Const LOG_SOURCE As String = "sync"
Private Const MB_YESNO As Long = &H4
Private Const MB_ICONERROR As Long = &H10
Private Const MB_ICONWARNING As Long = &H30
Private Const MB_ICONINFO As Long = &H40
Private Const MB_RIGHT As Long = &H80000
Private Const MB_RTLREADING As Long = &H100000
Private Const ID_YES As Long = 6

' ----------------------------------------------------------------------------
' Add-in life cycle
' ----------------------------------------------------------------------------

Public Sub AddinOpened()
    On Error Resume Next
    LogInfo LOG_SOURCE, GcvAddin.ADDIN_NAME & " " & GcvAddin.ADDIN_VERSION & " opened."
End Sub

Public Sub AddinClosing()
End Sub

Public Sub AddinSelfCheck()
    SyncCore.JsonString "x"
    SyncCore.IsSchoolCode "S01"
End Sub

' ----------------------------------------------------------------------------
' Ribbon callbacks (customUI14.xml)
' ----------------------------------------------------------------------------

Public Sub Ribbon_SyncNow(control As IRibbonControl)
    SyncNow
End Sub

Public Sub Ribbon_SyncSettings(control As IRibbonControl)
    EditSettings
End Sub

Public Sub SyncNow()
    Dim wb As Workbook
    Dim target As GoogleScriptTarget
    Dim result As SyncResult
    Dim started As Single

    On Error Resume Next
    Set wb = ActiveWorkbook
    On Error GoTo Failed
    If wb Is Nothing Then
        Show T_OpenWorkbook(), MB_ICONINFO
        Exit Sub
    End If
    If Not SyncCore.IsGcvWorkbook(wb) Then
        Show T_NotGcv(), MB_ICONINFO
        Exit Sub
    End If
    If Not IsConfigured() Then
        If Not EditSettings() Then Exit Sub
    End If

    Set target = New GoogleScriptTarget
    target.Init SyncCore.SettingValue("Url")
    started = Timer
    LogInfo LOG_SOURCE, "Sync started (" & SyncCore.SheetsToSync(wb).Count & " sheet(s))."
    Application.Cursor = xlWait
    Set result = SyncCore.RunSync(wb, target, SyncCore.SettingValue("School"), SyncCore.SettingValue("Key"))
    Application.Cursor = xlDefault
    Application.StatusBar = False
    Report result, Timer - started
    Exit Sub

Failed:
    Application.Cursor = xlDefault
    Application.StatusBar = False
    LogError LOG_SOURCE, "Sync stopped by an unexpected error"
    Show T_Failed(Err.Description), MB_ICONERROR
End Sub

Private Sub Report(ByVal result As SyncResult, ByVal seconds As Single)
    Select Case result.Outcome
        Case soOk
            LogInfo LOG_SOURCE, "Sync finished: " & result.SheetCount & " sheet(s), " & result.RowCount & _
                                " row(s) in " & Format$(seconds, "0.0") & " s."
            If Ask(T_Done(result.SheetCount, result.RowCount)) And Len(result.SheetUrl) > 0 Then
                ThisWorkbook.FollowHyperlink result.SheetUrl
            End If
        Case soOffline
            LogWarn LOG_SOURCE, "Sync: no connection (" & result.Message & ")."
            Show T_Offline(), MB_ICONWARNING
        Case soRejected
            LogWarn LOG_SOURCE, "Sync refused by the server: " & result.Message
            Show T_Rejected(result.Message), MB_ICONERROR
        Case soHttpError
            LogWarn LOG_SOURCE, "Sync: " & result.Message & "."
            Show T_Failed(result.Message), MB_ICONERROR
        Case soNotConfigured
            Show T_NotConfigured(), MB_ICONINFO
        Case soNotGcvWorkbook
            Show T_NotGcv(), MB_ICONINFO
    End Select
End Sub

' ----------------------------------------------------------------------------
' Settings (asked once per PC; the technical manager gives the values)
' ----------------------------------------------------------------------------

Public Function IsConfigured() As Boolean
    IsConfigured = SyncCore.IsWebAppUrl(SyncCore.SettingValue("Url")) And _
                   SyncCore.IsSchoolCode(SyncCore.SettingValue("School")) And _
                   Len(SyncCore.SettingValue("Key")) > 0
End Function

' Returns True when all three values were entered and saved.
Public Function EditSettings() As Boolean
    Dim url As Variant
    Dim school As Variant
    Dim key As Variant

    url = Application.InputBox(T_AskUrl(), T_Title(), SyncCore.SettingValue("Url"), Type:=2)
    If VarType(url) = vbBoolean Then Exit Function
    url = Trim$(url)
    If Not SyncCore.IsWebAppUrl(url) Then
        Show T_BadUrl(), MB_ICONWARNING
        Exit Function
    End If
    school = Application.InputBox(T_AskSchool(), T_Title(), SyncCore.SettingValue("School"), Type:=2)
    If VarType(school) = vbBoolean Then Exit Function
    school = Trim$(school)
    If Not SyncCore.IsSchoolCode(school) Then
        Show T_BadSchool(), MB_ICONWARNING
        Exit Function
    End If
    key = Application.InputBox(T_AskKey(), T_Title(), vbNullString, Type:=2)
    If VarType(key) = vbBoolean Then Exit Function
    key = Trim$(key)
    If Len(key) = 0 Then
        If Len(SyncCore.SettingValue("Key")) = 0 Then Exit Function
        key = SyncCore.SettingValue("Key")                ' empty = keep the saved key
    End If

    SyncCore.SaveSettingValue "Url", url
    SyncCore.SaveSettingValue "School", school
    SyncCore.SaveSettingValue "Key", key
    LogInfo LOG_SOURCE, "Sync settings saved (school " & school & ")."
    Show T_Saved(), MB_ICONINFO
    EditSettings = True
End Function

' ----------------------------------------------------------------------------
' Messages (Unicode, right-to-left)
' ----------------------------------------------------------------------------

Private Sub Show(ByVal text As String, ByVal icon As Long)
    MessageBoxW Application.hWnd, StrPtr(text), StrPtr(T_Title()), icon Or MB_RIGHT Or MB_RTLREADING
End Sub

Private Function Ask(ByVal text As String) As Boolean
    Ask = (MessageBoxW(Application.hWnd, StrPtr(text), StrPtr(T_Title()), _
                       MB_YESNO Or MB_ICONINFO Or MB_RIGHT Or MB_RTLREADING) = ID_YES)
End Function

Public Function T_Title() As String
    T_Title = FromCodePoints("0645 0632 0627 0645 0646 0629 0020 0627 0644 0645 0644 0641")
End Function

Public Function T_Done(ByVal sheets As Long, ByVal rows As Long) As String
    T_Done = FromCodePoints("062A 0645 062A 0020 0627 0644 0645 0632 0627 0645 0646 0629 0020 0628 0646 062C 0627 062D 002E") & vbCrLf & _
             FromCodePoints("0627 0644 0623 0648 0631 0627 0642 003A 0020") & sheets & "    " & FromCodePoints("0627 0644 0635 0641 0648 0641 003A 0020") & rows & vbCrLf & vbCrLf & _
             FromCodePoints("0647 0644 0020 062A 0631 064A 062F 0020 0641 062A 062D 0020 0645 0644 0641 0020 0047 006F 006F 0067 006C 0065 0020 0627 0644 0622 0646 061F")
End Function

Public Function T_Offline() As String
    T_Offline = FromCodePoints("0644 0627 0020 064A 0648 062C 062F 0020 0627 062A 0635 0627 0644 0020 0628 0627 0644 0625 0646 062A 0631 0646 062A 002E 0020 0644 0645 0020 064A 062A 063A 064A 0631 0020 0634 064A 0621 061B 0020 0623 0639 062F 0020 0627 0644 0645 062D 0627 0648 0644 0629 0020 0639 0646 062F 0020 062A 0648 0641 0631 0020 0627 0644 0627 062A 0635 0627 0644 002E")
End Function

Public Function T_Rejected(ByVal details As String) As String
    T_Rejected = FromCodePoints("0631 0641 0636 0020 0627 0644 062E 0627 062F 0645 0020 0627 0644 0645 0632 0627 0645 0646 0629 003A") & vbCrLf & details & vbCrLf & vbCrLf & _
                 FromCodePoints("062A 062D 0642 0642 0020 0645 0646 0020 0625 0639 062F 0627 062F 0627 062A 0020 0627 0644 0645 0632 0627 0645 0646 0629 0020 0623 0648 0020 0623 0628 0644 063A 0020 0627 0644 0645 0633 0624 0648 0644 0020 0627 0644 062A 0642 0646 064A 002E")
End Function

Public Function T_Failed(ByVal details As String) As String
    T_Failed = FromCodePoints("062A 0639 0630 0631 062A 0020 0627 0644 0645 0632 0627 0645 0646 0629 003A") & vbCrLf & details & vbCrLf & vbCrLf & FromCodePoints("0644 0645 0020 064A 062A 063A 064A 0631 0020 0634 064A 0621 0020 0641 064A 0020 0627 0644 0645 0644 0641 002E")
End Function

Public Function T_NotGcv() As String
    T_NotGcv = FromCodePoints("0647 0630 0627 0020 0644 064A 0633 0020 0645 0644 0641 0020 0625 062F 0627 0631 0629 0020 0645 062F 0631 0633 0629 0020 0645 0646 0020 0642 0631 064A 0629 0020 0627 0644 0623 0637 0641 0627 0644 060C 0020 0644 0630 0644 0643 0020 0644 0627 0020 062A 062A 0645 0020 0645 0632 0627 0645 0646 062A 0647 002E")
End Function

Public Function T_OpenWorkbook() As String
    T_OpenWorkbook = FromCodePoints("0627 0641 062A 062D 0020 0645 0644 0641 0020 0625 062F 0627 0631 0629 0020 0627 0644 0645 062F 0631 0633 0629 0020 0623 0648 0644 0627 064B 002E")
End Function

Public Function T_NotConfigured() As String
    T_NotConfigured = FromCodePoints("0625 0639 062F 0627 062F 0627 062A 0020 0627 0644 0645 0632 0627 0645 0646 0629 0020 063A 064A 0631 0020 0645 0643 062A 0645 0644 0629 002E 0020 0627 0636 063A 0637 0020 0632 0631 0020 0625 0639 062F 0627 062F 0627 062A 0020 0627 0644 0645 0632 0627 0645 0646 0629 002E")
End Function

Public Function T_AskUrl() As String
    T_AskUrl = FromCodePoints("0631 0627 0628 0637 0020 062E 062F 0645 0629 0020 0627 0644 0645 0632 0627 0645 0646 0629 0020 0028 0645 0646 0020 0627 0644 0645 0633 0624 0648 0644 0020 0627 0644 062A 0642 0646 064A 0029 003A")
End Function

Public Function T_AskSchool() As String
    T_AskSchool = FromCodePoints("0631 0645 0632 0020 0627 0644 0645 062F 0631 0633 0629 0020 0028 0645 062B 0644 0020 0053 0030 0033 0029 003A")
End Function

Public Function T_AskKey() As String
    T_AskKey = FromCodePoints("0645 0641 062A 0627 062D 0020 0627 0644 0645 0632 0627 0645 0646 0629 0020 0028 0627 062A 0631 0643 0647 0020 0641 0627 0631 063A 0627 064B 0020 0644 0644 0625 0628 0642 0627 0621 0020 0639 0644 0649 0020 0627 0644 0645 0641 062A 0627 062D 0020 0627 0644 0645 062D 0641 0648 0638 0029 003A")
End Function

Public Function T_BadUrl() As String
    T_BadUrl = FromCodePoints("0627 0644 0631 0627 0628 0637 0020 063A 064A 0631 0020 0635 062D 064A 062D 002E 0020 064A 062C 0628 0020 0623 0646 0020 064A 0628 062F 0623 0020 0628 0640 0020 0068 0074 0074 0070 0073 003A 002F 002F 0073 0063 0072 0069 0070 0074 002E 0067 006F 006F 0067 006C 0065 002E 0063 006F 006D 002F 006D 0061 0063 0072 006F 0073 002F 0073 002F 0020 0648 064A 0646 062A 0647 064A 0020 0628 0640 0020 002F 0065 0078 0065 0063")
End Function

Public Function T_BadSchool() As String
    T_BadSchool = FromCodePoints("0631 0645 0632 0020 0627 0644 0645 062F 0631 0633 0629 0020 063A 064A 0631 0020 0635 062D 064A 062D 003A 0020 0623 062D 0631 0641 0020 0625 0646 062C 0644 064A 0632 064A 0629 0020 0648 0623 0631 0642 0627 0645 0020 0648 0634 0631 0637 0629 0020 0641 0642 0637 002E")
End Function

Public Function T_Saved() As String
    T_Saved = FromCodePoints("062A 0645 0020 062D 0641 0638 0020 0625 0639 062F 0627 062F 0627 062A 0020 0627 0644 0645 0632 0627 0645 0646 0629 0020 0639 0644 0649 0020 0647 0630 0627 0020 0627 0644 062C 0647 0627 0632 002E")
End Function

Private Function FromCodePoints(ByVal codePoints As String) As String
    Dim parts() As String
    Dim i As Long
    parts = Split(codePoints, " ")
    For i = LBound(parts) To UBound(parts)
        FromCodePoints = FromCodePoints & ChrW$(CLng("&H" & parts(i)))
    Next i
End Function
