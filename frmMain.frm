VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Winsock Demo"
   ClientHeight    =   6060
   ClientLeft      =   -30
   ClientTop       =   255
   ClientWidth     =   7170
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   7.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   404
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   478
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Send Request"
      Height          =   300
      Left            =   5160
      TabIndex        =   3
      Top             =   600
      Width           =   1815
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   600
      Width           =   4815
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   4695
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1080
      Width           =   6615
   End
   Begin VB.Label Label1 
      Caption         =   "Socket Provider"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mSock As Winsock
Attribute mSock.VB_VarHelpID = -1

'================================================
' Form_Load
'================================================
Private Sub Form_Load()
    Debug_Print "Form_Load"
    Combo1.AddItem "Microsoft"
    Combo1.AddItem "EBArtSoft + DeepSeek V4.1 Flash 1017"
    'Combo1.AddItem "EBArtSoft"
    Combo1.ListIndex = 0
End Sub

'================================================
' Form_Unload
'================================================
Private Sub Form_Unload(cancel As Integer)
    Debug_Print "Form_Unload"
    Set mSock = Nothing
End Sub

'================================================
' Command1_Click
'================================================
Private Sub Command1_Click()
    If Not mSock Is Nothing Then
        mSock.Close
        Set mSock = Nothing
    End If
    Text1.Text = vbNullString
    Set mSock = CreateNewWinsock(Combo1.ListIndex)
    If mSock Is Nothing Then Exit Sub
    mSock.Protocol = sckTCPProtocol
    mSock.Connect "jsonplaceholder.typicode.com", 80
End Sub

'================================================
' mSock_Close
'================================================
Private Sub mSock_Close()
    Debug_Print "mSock_Close"
    mSock.Close
End Sub

'================================================
' mSock_error
'================================================
Private Sub mSock_error(ByVal number As Integer, Description As String, ByVal scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Debug_Print "mSock_error " & number & " " & Description & " " & scode & " " & Source & " " & HelpFile & " " & HelpContext & " " & CancelDisplay
    mSock.Close
End Sub

'================================================
' mSock_SendComplete
'================================================
Private Sub mSock_SendComplete()
    Debug_Print "mSock_SendComplete"
End Sub

'================================================
' mSock_SendProgress
'================================================
Private Sub mSock_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
    Debug_Print "mSock_SendProgress " & bytesSent & " " & bytesRemaining
End Sub

'================================================
' mSock_connect
'================================================
Private Sub mSock_connect()
    Debug_Print "mSock_connect"
    mSock.SendData _
    "GET /posts/1 HTTP/1.1" & vbCrLf & _
    "Host: jsonplaceholder.typicode.com" & vbCrLf & _
    "Connection: close" & vbCrLf & _
    vbCrLf
End Sub

'================================================
' mSock_DataArrival
'================================================
Private Sub mSock_DataArrival(ByVal bytesTotal As Long)
    Debug_Print "mSock_connect " & bytesTotal
    Dim s As String
    mSock.GetData s
    Debug_Print s
End Sub


'================================================
' CreateNewWinsock
'================================================
Public Function CreateNewWinsock(ByVal pType As Long) As Winsock
    Select Case pType
    Case 0: Set CreateNewWinsock = New Winsock
    Case 1: Set CreateNewWinsock = New_Winsock2
    'Case 2: Set CreateNewWinsock = New_Winsock
    End Select
End Function

'================================================
' Debug_Print
'================================================
Public Sub Debug_Print(ByVal pText As String, Optional ByVal pLogType As Long = 1)
    Select Case pLogType
    Case 0
        Debug.Print pText
        
    Case 1
        frmMain.Text1.Text = frmMain.Text1.Text & pText & vbCrLf
    
    Case 2
        Open App.Path & "\output.log" For Append As #1
        Print #1, pText & vbCrLf
        Close #1
        
    End Select
End Sub

