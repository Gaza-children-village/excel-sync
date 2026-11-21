Attribute VB_Name = "SyncUi"
Option Explicit

' ============================================================================
' GCV Sync add-in: start-up, ribbon group "sync" and messages.
'   Sync now   - sends the active GCV workbook to the school's Google Sheet
'   Settings   - the school number (once per PC). The web-app link is built
'                in (SyncCore.SYNC_URL); the PC key is made automatically.
' ============================================================================

Private Declare PtrSafe Function MessageBoxW Lib "user32" (ByVal hWnd As LongPtr, _
    ByVal lpText As LongPtr, ByVal lpCaption As LongPtr, ByVal uType As Long) As Long

Private Const LOG_SOURCE As String = "sync"
Private Const MB_ICONERROR As Long = &H10
Private Const MB_ICONWARNING As Long = &H30
Private Const MB_ICONINFO As Long = &H40
Private Const MB_RIGHT As Long = &H80000
Private Const MB_RTLREADING As Long = &H100000

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
    SyncCore.IsSchoolCode "GCV01"
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
    Dim school As String

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
    If Not SyncCore.IsWebAppUrl(SyncCore.SyncUrl()) Then
        LogError LOG_SOURCE, "Sync: the add-in has no valid web-app link (SYNC_URL)."
        Show T_NoServiceLink(), MB_ICONERROR
        Exit Sub
    End If
    If Not IsConfigured() Then
        If Not EditSettings() Then Exit Sub
    End If

    school = SyncCore.SettingValue("School")
    Set target = New GoogleScriptTarget
    target.Init SyncCore.SyncUrl()
    started = Timer
    LogInfo LOG_SOURCE, "Sync of " & school & " started (" & SyncCore.SheetsToSync(wb).Count & " sheet(s))."
    Application.Cursor = xlWait
    Set result = SyncCore.RunSync(wb, target, school, SyncCore.DeviceKey())
    Application.Cursor = xlDefault
    Application.StatusBar = False
    Report result, school, Timer - started
    Exit Sub

Failed:
    Application.Cursor = xlDefault
    Application.StatusBar = False
    LogError LOG_SOURCE, "Sync stopped by an unexpected error"
    Show T_Failed(Err.Description), MB_ICONERROR
End Sub

Private Sub Report(ByVal result As SyncResult, ByVal school As String, ByVal seconds As Single)
    Select Case result.Outcome
        Case soOk
            LogInfo LOG_SOURCE, "Sync finished: " & result.SheetCount & " sheet(s), " & result.RowCount & _
                                " row(s) in " & Format$(seconds, "0.0") & " s." & _
                                IIf(result.Registered, " This PC is now linked to " & school & ".", "")
            If result.Registered Then
                Show T_Linked(school) & vbCrLf & vbCrLf & T_Done(result.SheetCount, result.RowCount), MB_ICONINFO
            Else
                Show T_Done(result.SheetCount, result.RowCount), MB_ICONINFO
            End If
        Case soOffline
            LogWarn LOG_SOURCE, "Sync: no connection (" & result.Message & ")."
            Show T_Offline(), MB_ICONWARNING
        Case soRejected
            LogWarn LOG_SOURCE, "Sync refused by the server (" & result.ErrorCode & "): " & result.Message
            Select Case result.ErrorCode
                Case "school-taken": Show T_SchoolTaken(school), MB_ICONERROR
                Case "registration-closed": Show T_RegistrationClosed(school), MB_ICONERROR
                Case "wrong-school": Show T_WrongSchool(result.Message), MB_ICONERROR
                Case "busy": Show T_Busy(), MB_ICONWARNING
                Case Else: Show T_Rejected(result.Message), MB_ICONERROR
            End Select
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
' Settings: only the school number (once per PC)
' ----------------------------------------------------------------------------

Public Function IsConfigured() As Boolean
    IsConfigured = SyncCore.IsSchoolCode(SyncCore.SettingValue("School"))
End Function

' Returns True when a valid school number was entered and saved.
Public Function EditSettings() As Boolean
    Dim answer As Variant
    Dim current As String
    Dim school As String

    current = SyncCore.SettingValue("School")
    If SyncCore.IsSchoolCode(current) Then current = CStr(CLng(Mid$(current, 4))) Else current = vbNullString
    answer = Application.InputBox(T_AskSchool(), T_Title(), current, Type:=2)
    If VarType(answer) = vbBoolean Then Exit Function
    school = SyncCore.SchoolCodeFromInput(CStr(answer))
    If Len(school) = 0 Then
        Show T_BadSchool(), MB_ICONWARNING
        Exit Function
    End If
    SyncCore.SaveSettingValue "School", school
    LogInfo LOG_SOURCE, "Sync settings saved (school " & school & ")."
    Show T_Saved(school), MB_ICONINFO
    EditSettings = True
End Function

' ----------------------------------------------------------------------------
' Messages (Unicode, right-to-left)
' ----------------------------------------------------------------------------

Private Sub Show(ByVal text As String, ByVal icon As Long)
    MessageBoxW Application.hWnd, StrPtr(text), StrPtr(T_Title()), icon Or MB_RIGHT Or MB_RTLREADING
End Sub

Public Function T_Title() As String
    T_Title = FromCodePoints("0645 0632 0627 0645 0646 0629 0020 0627 0644 0645 0644 0641")
End Function

Public Function T_Done(ByVal sheets As Long, ByVal rows As Long) As String
    T_Done = FromCodePoints("062A 0645 062A 0020 0627 0644 0645 0632 0627 0645 0646 0629 0020 0628 0646 062C 0627 062D 002E") & vbCrLf & _
             FromCodePoints("0627 0644 0623 0648 0631 0627 0642 003A 0020") & sheets & "    " & FromCodePoints("0627 0644 0635 0641 0648 0641 003A 0020") & rows
End Function

Public Function T_Linked(ByVal school As String) As String
    T_Linked = FromCodePoints("062A 0645 0020 0631 0628 0637 0020 0647 0630 0627 0020 0627 0644 062C 0647 0627 0632 0020 0628 0627 0644 0645 062F 0631 0633 0629 0020") & school & "."
End Function

Public Function T_Offline() As String
    T_Offline = FromCodePoints("0644 0627 0020 064A 0648 062C 062F 0020 0627 062A 0635 0627 0644 0020 0628 0627 0644 0625 0646 062A 0631 0646 062A 002E 0020 0644 0645 0020 064A 062A 063A 064A 0631 0020 0634 064A 0621 061B 0020 0623 0639 062F 0020 0627 0644 0645 062D 0627 0648 0644 0629 0020 0639 0646 062F 0020 062A 0648 0641 0631 0020 0627 0644 0627 062A 0635 0627 0644 002E")
End Function

Public Function T_SchoolTaken(ByVal school As String) As String
    T_SchoolTaken = FromCodePoints("0627 0644 0645 062F 0631 0633 0629 0020") & school & FromCodePoints("0020 0645 0631 0628 0648 0637 0629 0020 0628 062C 0647 0627 0632 0020 0622 062E 0631 060C 0020 0644 0630 0644 0643 0020 0644 0645 0020 062A 062A 0645 0020 0627 0644 0645 0632 0627 0645 0646 0629 002E") & vbCrLf & vbCrLf & _
                    FromCodePoints("062A 0623 0643 062F 0020 0645 0646 0020 0631 0642 0645 0020 0627 0644 0645 062F 0631 0633 0629 0020 0641 064A 0020 0625 0639 062F 0627 062F 0627 062A 0020 0627 0644 0645 0632 0627 0645 0646 0629 002E 0020 0625 0630 0627 0020 062A 063A 064A 0651 0631 0020 062C 0647 0627 0632 0020 0627 0644 0645 062F 0631 0633 0629 0020 0641 0623 0628 0644 063A 0020 0627 0644 0645 0633 0624 0648 0644 0020 0627 0644 062A 0642 0646 064A 002E")
End Function

Public Function T_RegistrationClosed(ByVal school As String) As String
    T_RegistrationClosed = FromCodePoints("0627 0644 0645 062F 0631 0633 0629 0020") & school & FromCodePoints("0020 063A 064A 0631 0020 0645 0633 062C 0644 0629 0020 0641 064A 0020 062E 062F 0645 0629 0020 0627 0644 0645 0632 0627 0645 0646 0629 002E") & vbCrLf & vbCrLf & _
                           FromCodePoints("062A 0623 0643 062F 0020 0645 0646 0020 0631 0642 0645 0020 0627 0644 0645 062F 0631 0633 0629 0020 0641 064A 0020 0625 0639 062F 0627 062F 0627 062A 0020 0627 0644 0645 0632 0627 0645 0646 0629 0020 0623 0648 0020 0623 0628 0644 063A 0020 0627 0644 0645 0633 0624 0648 0644 0020 0627 0644 062A 0642 0646 064A 002E")
End Function

Public Function T_WrongSchool(ByVal details As String) As String
    T_WrongSchool = FromCodePoints("0647 0630 0627 0020 0627 0644 0645 0644 0641 0020 064A 062E 0635 0020 0645 062F 0631 0633 0629 0020 0623 062E 0631 0649 060C 0020 0644 0630 0644 0643 0020 0644 0645 0020 062A 062A 0645 0020 0627 0644 0645 0632 0627 0645 0646 0629 002E") & vbCrLf & details & vbCrLf & vbCrLf & _
                    FromCodePoints("062A 0623 0643 062F 0020 0623 0646 0643 0020 0641 062A 062D 062A 0020 0645 0644 0641 0020 0645 062F 0631 0633 062A 0643 060C 0020 0623 0648 0020 0623 0628 0644 063A 0020 0627 0644 0645 0633 0624 0648 0644 0020 0627 0644 062A 0642 0646 064A 002E")
End Function

Public Function T_Busy() As String
    T_Busy = FromCodePoints("062E 062F 0645 0629 0020 0627 0644 0645 0632 0627 0645 0646 0629 0020 0645 0634 063A 0648 0644 0629 0020 0628 0645 0632 0627 0645 0646 0629 0020 0623 062E 0631 0649 002E 0020 0623 0639 062F 0020 0627 0644 0645 062D 0627 0648 0644 0629 0020 0628 0639 062F 0020 062F 0642 064A 0642 0629 002E")
End Function

Public Function T_Rejected(ByVal details As String) As String
    T_Rejected = FromCodePoints("0631 0641 0636 0020 0627 0644 062E 0627 062F 0645 0020 0627 0644 0645 0632 0627 0645 0646 0629 003A") & vbCrLf & details & vbCrLf & vbCrLf & _
                 FromCodePoints("0623 0628 0644 063A 0020 0627 0644 0645 0633 0624 0648 0644 0020 0627 0644 062A 0642 0646 064A 002E")
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
    T_NotConfigured = FromCodePoints("0644 0645 0020 064A 062A 0645 0020 0625 062F 062E 0627 0644 0020 0631 0642 0645 0020 0627 0644 0645 062F 0631 0633 0629 0020 0628 0639 062F 002E 0020 0627 0636 063A 0637 0020 0632 0631 0020 0625 0639 062F 0627 062F 0627 062A 0020 0627 0644 0645 0632 0627 0645 0646 0629 002E")
End Function

Public Function T_NoServiceLink() As String
    T_NoServiceLink = FromCodePoints("0631 0627 0628 0637 0020 062E 062F 0645 0629 0020 0627 0644 0645 0632 0627 0645 0646 0629 0020 063A 064A 0631 0020 0645 0636 0628 0648 0637 0020 0641 064A 0020 0647 0630 0647 0020 0627 0644 0623 062F 0627 0629 002E 0020 0623 0628 0644 063A 0020 0627 0644 0645 0633 0624 0648 0644 0020 0627 0644 062A 0642 0646 064A 002E")
End Function

Public Function T_AskSchool() As String
    T_AskSchool = FromCodePoints("0631 0642 0645 0020 0627 0644 0645 062F 0631 0633 0629 0020 0028 0645 062B 0644 0627 064B 0020 0035 0020 0644 0644 0645 062F 0631 0633 0629 0020 0047 0043 0056 0030 0035 0029 003A")
End Function

Public Function T_BadSchool() As String
    T_BadSchool = FromCodePoints("0631 0642 0645 0020 0627 0644 0645 062F 0631 0633 0629 0020 063A 064A 0631 0020 0635 062D 064A 062D 002E 0020 0627 0643 062A 0628 0020 0631 0642 0645 0627 064B 0020 0645 062B 0644 0020 0035 0020 0623 0648 0020 0047 0043 0056 0030 0035 002E")
End Function

Public Function T_Saved(ByVal school As String) As String
    T_Saved = FromCodePoints("062A 0645 0020 062D 0641 0638 0020 0631 0642 0645 0020 0627 0644 0645 062F 0631 0633 0629 0020 0639 0644 0649 0020 0647 0630 0627 0020 0627 0644 062C 0647 0627 0632 003A 0020") & school
End Function

Private Function FromCodePoints(ByVal codePoints As String) As String
    Dim parts() As String
    Dim i As Long
    parts = Split(codePoints, " ")
    For i = LBound(parts) To UBound(parts)
        FromCodePoints = FromCodePoints & ChrW$(CLng("&H" & parts(i)))
    Next i
End Function
