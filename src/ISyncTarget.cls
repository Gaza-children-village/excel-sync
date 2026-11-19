VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ISyncTarget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Where a sync goes. GoogleScriptTarget sends to the GCV Google Apps Script
' web app; a future GCV server target implements the same method.

Public Function Send(ByVal payloadJson As String) As SyncResult
End Function
