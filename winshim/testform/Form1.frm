VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form Form1 
   Caption         =   "winshim form test"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock Winsock2 
      Left            =   1320
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   120
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   400
      Left            =   2400
      Top             =   1680
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    ' The replacement does not receive design-time property values (see
    ' winshim/README.md), so the control is configured at run time.
    On Error Resume Next
    Winsock1.Protocol = 0
    Winsock1.LocalPort = 65000
    modMain.Log "Form_Load state=" & Winsock1.State & " proto=" & Winsock1.Protocol & _
                " localport=" & Winsock1.LocalPort & " handle=" & Winsock1.SocketHandle
    Winsock1.Listen
    modMain.Log "listen state=" & Winsock1.State & " err=" & Err.Number & " " & Err.Description
End Sub

Private Sub Timer1_Timer()
    Static n As Long
    On Error Resume Next
    n = n + 1
    If n = 1 Then
        modMain.Log "timer: connecting"
        Winsock2.Connect "127.0.0.1", 65000
        modMain.Log "connect state=" & Winsock2.State & " err=" & Err.Number & " " & Err.Description
    ElseIf n >= 8 Then
        Timer1.Enabled = False
        modMain.Log "timer: unload, states " & Winsock1.State & "/" & Winsock2.State
        Unload Me
    End If
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
    On Error Resume Next
    modMain.Log "server connreq " & requestID
    Winsock1.Close
    Winsock1.Accept requestID
    modMain.Log "server accepted state=" & Winsock1.State & " err=" & Err.Number
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim s As String
    On Error Resume Next
    Winsock1.GetData s, vbString
    modMain.Log "server got '" & s & "' bytes=" & bytesTotal & " err=" & Err.Number
    Winsock1.SendData "world"
End Sub

Private Sub Winsock1_Close()
    modMain.Log "server close event"
End Sub

Private Sub Winsock2_Connect()
    On Error Resume Next
    modMain.Log "client connect event state=" & Winsock2.State
    Winsock2.SendData "hello"
    modMain.Log "client send err=" & Err.Number
End Sub

Private Sub Winsock2_DataArrival(ByVal bytesTotal As Long)
    Dim s As String
    On Error Resume Next
    Winsock2.GetData s, vbString
    modMain.Log "client got '" & s & "' bytes=" & bytesTotal & " err=" & Err.Number
    Winsock2.Close
    Winsock1.Close
    modMain.Log "closed states " & Winsock2.State & "/" & Winsock1.State
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, _
                            ByVal Source As String, ByVal HelpFile As String, _
                            ByVal HelpContext As Long, CancelDisplay As Boolean)
    CancelDisplay = True
    modMain.Log "server error " & Number & " " & Scode & " " & Description
End Sub

Private Sub Winsock2_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, _
                            ByVal Source As String, ByVal HelpFile As String, _
                            ByVal HelpContext As Long, CancelDisplay As Boolean)
    CancelDisplay = True
    modMain.Log "client error " & Number & " " & Scode & " " & Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    modMain.Log "Form_Unload"
End Sub