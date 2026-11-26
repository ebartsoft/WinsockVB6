VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPerfRow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Operation As String
Public Iterations As Long
Public W2Seconds As Double
Public OcxSeconds As Double
Public W2Ops As Double
Public OcxOps As Double
Public W2BytesPerSec As Double
Public OcxBytesPerSec As Double
Public Notes As String

Public Function RatioText() As String
    If OcxSeconds <= 0 Or W2Seconds <= 0 Then
        RatioText = "-"
    Else
        RatioText = Format$(W2Seconds / OcxSeconds, "0.00") & "x"
    End If
End Function
