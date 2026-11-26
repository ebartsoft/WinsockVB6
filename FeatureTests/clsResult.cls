VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Group As String
Public Scenario As String
Public W2 As clsRun
Public Ocx As clsRun
Public Status As String
Public Notes As String
Public Diff As String

Public Function StatusPrefix() As String
    StatusPrefix = "[" & Status & "]"
End Function
