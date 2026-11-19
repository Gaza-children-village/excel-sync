VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GoogleScriptTarget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ISyncTarget

' POSTs the payload (UTF-8 JSON) to the GCV Sync Apps Script web app.
' Apps Script answers a POST with a redirect that WinHTTP follows as GET;
' the JSON answer is read from there.

Private Const CONNECT_TIMEOUT_MS As Long = 15000
Private Const TRANSFER_TIMEOUT_MS As Long = 180000

Private mUrl As String

Public Sub Init(ByVal webAppUrl As String)
    mUrl = webAppUrl
End Sub

Private Function ISyncTarget_Send(ByVal payloadJson As String) As SyncResult
    Dim http As Object
    Dim result As SyncResult

    On Error GoTo Offline
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.SetTimeouts CONNECT_TIMEOUT_MS, CONNECT_TIMEOUT_MS, TRANSFER_TIMEOUT_MS, TRANSFER_TIMEOUT_MS
    http.Open "POST", mUrl, False
    http.SetRequestHeader "Content-Type", "application/json; charset=utf-8"
    http.Send SyncCore.Utf8Bytes(payloadJson)
    On Error GoTo 0
    Set ISyncTarget_Send = SyncCore.ParseAnswer(http.Status, http.ResponseText)
    Exit Function

Offline:
    Set result = New SyncResult
    result.Outcome = soOffline
    result.Message = Err.Description
    Set ISyncTarget_Send = result
End Function
