Attribute VB_Name = "modMain"
Option Explicit

Public gReqID As Long

Public Sub Log(ByVal s As String)
    Dim f As Integer
    On Error Resume Next
    f = FreeFile
    Open App.Path & "\form.log" For Append As #f
    Print #f, Format$(Now, "hh:nn:ss") & "  " & s
    Close #f
End Sub

Public Sub Main()
    Log "main start"
    On Error Resume Next
    Err.Clear
    Load Form1
    Log "load err=" & Err.Number & " " & Err.Description
    Err.Clear
    Form1.Show
    Log "show err=" & Err.Number & " " & Err.Description
End Sub
