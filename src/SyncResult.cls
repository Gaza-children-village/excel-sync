VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SyncResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Result of one sync.

Public Enum SyncOutcome
    soOk = 1
    soRejected = 2          ' the server refused (wrong key, other school, ...): Message says why
    soHttpError = 3
    soOffline = 4
    soNotGcvWorkbook = 5
    soNotConfigured = 6
End Enum

Public Outcome As SyncOutcome
Public Message As String
Public SheetUrl As String
Public SheetCount As Long
Public RowCount As Long
Public HttpStatus As Long
