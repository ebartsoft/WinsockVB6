VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public ImplName As String
Public Passed As Boolean
Public Failure As String
Public Trace As String
Public Snapshot As String
Public Details As String
Public KnownGap As Boolean
Public Skipped As String

Public Sub Reset(ByVal name As String)
    ImplName = name
    Passed = False
    Failure = ""
    Trace = ""
    Snapshot = ""
    Details = ""
    KnownGap = False
    Skipped = ""
End Sub
