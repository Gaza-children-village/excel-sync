VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FakeSyncTarget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ISyncTarget

' Test double: keeps the payload and answers with Answer (or "offline").

Public Offline As Boolean
Public LastPayload As String
Public Calls As Long
Public AnswerStatus As Long
Public AnswerText As String

Private Sub Class_Initialize()
    AnswerStatus = 200
    AnswerText = "{""ok"": true, ""sheets"": 1, ""rows"": 1, ""url"": ""https://docs.google.com/x""}"
End Sub

Private Function ISyncTarget_Send(ByVal payloadJson As String) As SyncResult
    Dim result As SyncResult
    Calls = Calls + 1
    LastPayload = payloadJson
    If Offline Then
        Set result = New SyncResult
        result.Outcome = soOffline
        result.Message = "fake: offline"
        Set ISyncTarget_Send = result
    Else
        Set ISyncTarget_Send = SyncCore.ParseAnswer(AnswerStatus, AnswerText)
    End If
End Function
