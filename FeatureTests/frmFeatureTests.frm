VERSION 5.00
Begin VB.Form frmFeatureTests 
   Caption         =   "modWinsock2 Feature & Performance Tests"
   ClientHeight    =   8340
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12960
   LinkTopic       =   "Form1"
   ScaleHeight     =   8340
   ScaleWidth      =   12960
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Caption         =   "View Raw Log"
      Height          =   360
      Left            =   11040
      TabIndex        =   4
      Top             =   120
      Width           =   1680
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Save Reports"
      Height          =   360
      Left            =   9240
      TabIndex        =   3
      Top             =   120
      Width           =   1680
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Performance"
      Height          =   360
      Left            =   7440
      TabIndex        =   2
      Top             =   120
      Width           =   1680
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Functional"
      Height          =   360
      Left            =   5640
      TabIndex        =   1
      Top             =   120
      Width           =   1680
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Run All"
      Height          =   360
      Left            =   3840
      TabIndex        =   0
      Top             =   120
      Width           =   1680
   End
   Begin VB.Label Label1 
      Caption         =   "Ready. Press Run All (or launch with /autorun)."
      Height          =   600
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   12600
   End
   Begin VB.ListBox ListBox1 
      Height          =   7380
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   4440
   End
   Begin VB.TextBox Text1 
      Height          =   7380
      Left            =   4680
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   7
      Top             =   1200
      Width           =   8040
   End
End
Attribute VB_Name = "frmFeatureTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mAutoRun As Boolean

Private Sub Form_Load()
    Dim cmd As String
    Dim t0 As Double

    cmd = LCase$(Command$)

    ' Child mode: one matrix cell, written to a file, then exit.
    If InStr(1, cmd, "/cell:", vbTextCompare) > 0 Then
        modMatrix.RunCell ArgValue("/cell:"), ArgValue("/impl:"), ArgValue("/out:")
        Unload Me
        Exit Sub
    End If

    ' Parent mode: run the whole differential matrix.
    If InStr(1, cmd, "/matrix", vbTextCompare) > 0 Then
        mAutoRun = True
        Label1.Caption = "Running matrix..."
        DoEvents
        modMatrix.RunMatrix Me, ArgValue("/matrix:")
        Unload Me
        Exit Sub
    End If

    If InStr(1, cmd, "/autorun", vbTextCompare) = 0 Then
        Text1.Text = "Legend:" & vbCrLf & _
            "  [PASS]      both implementations behave the same" & vbCrLf & _
            "  [DIFF]      both pass, but the event traces differ" & vbCrLf & _
            "  [FAIL]      a scenario failed on at least one implementation" & vbCrLf & _
            "  [EXPECTED]  known v1 limitation (classified, not a failure)" & vbCrLf & _
            "  [PERF]      performance row" & vbCrLf & vbCrLf & _
            "Select a row on the left to see its details." & vbCrLf & vbCrLf & _
            "Command-line switches:" & vbCrLf & _
            "  /autorun      run and save reports, then exit" & vbCrLf & _
            "  /func         functional scenarios only" & vbCrLf & _
            "  /perf         performance suite only" & vbCrLf & _
            "  /only:<id>    run a single scenario (defaults, props, tcp_str," & vbCrLf & _
            "                tcp_bytes, tcp_multi, tcp_large, udp, dns, errors," & vbCrLf & _
            "                lifecycle, latebinding, interop)"
        Exit Sub
    End If

    mAutoRun = True
    Label1.Caption = "Running..."
    DoEvents

    t0 = Now
    On Error GoTo EH
    If InStr(1, cmd, "/perf", vbTextCompare) > 0 Then
        modTestEngine.RunPerformance Me
    ElseIf InStr(1, cmd, "/func", vbTextCompare) > 0 Then
        modTestEngine.RunFunctional Me
    Else
        modTestEngine.RunAll Me
    End If
    modTestEngine.SaveReports Me
    Label1.Caption = modTestEngine.SummaryLine() & "  (" & Format$(Now - t0, "hh:nn:ss") & ")"
    DoEvents
    Unload Me
    Exit Sub
EH:
    Label1.Caption = "Autorun error " & Err.Number & ": " & Err.Description
End Sub

Private Function ArgValue(ByVal name As String) As String
    Dim cmd As String
    Dim p As Long
    Dim rest As String
    Dim q As Long

    cmd = Command$
    p = InStr(1, cmd, name, vbTextCompare)
    If p = 0 Then Exit Function
    rest = Mid$(cmd, p + Len(name))
    q = InStr(1, rest, " /")
    If InStr(1, rest, """") = 1 Then
        rest = Mid$(rest, 2)
        q = InStr(1, rest, """")
        ArgValue = Left$(rest, q - 1)
        Exit Function
    End If
    If q > 0 Then rest = Left$(rest, q - 1)
    ArgValue = Trim$(rest)
End Function

Public Sub SetSummary(ByVal s As String)
    Label1.Caption = s
End Sub

Public Sub SetReportText(ByVal s As String)
    Text1.Text = s
    Text1.SelStart = 0
End Sub

Public Sub SetDetailText(ByVal s As String)
    Text1.Text = s
    Text1.SelStart = 0
End Sub

Private Sub Command1_Click()
    On Error GoTo EH
    Label1.Caption = "Running all scenarios..."
    DoEvents
    modTestEngine.RunAll Me
    modTestEngine.SaveReports Me
    Exit Sub
EH:
    Label1.Caption = "Run error " & Err.Number & ": " & Err.Description
End Sub

Private Sub Command2_Click()
    On Error GoTo EH
    Label1.Caption = "Running functional scenarios..."
    DoEvents
    modTestEngine.RunFunctional Me
    Exit Sub
EH:
    Label1.Caption = "Functional error " & Err.Number & ": " & Err.Description
End Sub

Private Sub Command3_Click()
    On Error GoTo EH
    Label1.Caption = "Running performance suite..."
    DoEvents
    modTestEngine.RunPerformance Me
    Exit Sub
EH:
    Label1.Caption = "Performance error " & Err.Number & ": " & Err.Description
End Sub

Private Sub Command4_Click()
    On Error Resume Next
    modTestEngine.SaveReports Me
    If Err.Number <> 0 Then Label1.Caption = "Save error " & Err.Number & ": " & Err.Description
End Sub

Private Sub Command5_Click()
    Text1.Text = modTestEngine.BuildRawLog()
    Text1.SelStart = 0
End Sub

Private Sub ListBox1_Click()
    modTestEngine.ShowItem Me, ListBox1.ListIndex
End Sub
