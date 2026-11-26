VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEndpoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mSock As Winsock
Attribute mSock.VB_VarHelpID = -1

Private mTrace As Collection
Private mLastBytes() As Byte

Public UseOcx       As Boolean
Public EndpointName As String
Public Echo         As Boolean
Public AutoRead     As Boolean
Public CloseOnData  As Boolean
Public CountOnly    As Boolean
Public ExpectedBytes As Long
Public ManualAccept As Boolean
Public RefuseAccept As Boolean
Public LastRequestID As Long

' Stream capture: accumulate everything received into one buffer so a scenario
' can compare it byte for byte with what was sent.
Public CaptureMode  As Boolean
Public CaptureLimit As Long
Public CaptureSize  As Long
Public CaptureOverflow As Boolean
Public ChildCaptureLimit As Long
Private mCapture() As Byte

Public DoneConnect  As Boolean
Public DoneData     As Boolean
Public DoneClose    As Boolean
Public DoneError    As Boolean
Public DataCount    As Long
Public LastDataSize As Long
Public LastString   As String
Public TotalReceived As Long
Public LastErrorNo  As Long
Public LastErrorText As String
Public LastErrorSource As String
Public LastErrorHelp As String
Public StateAfterConnectCall As Long
Public LastEventSeconds As Double
Public IsClosed As Boolean

'------------------------------------------------------------------------------
' Init
'------------------------------------------------------------------------------
Public Sub Init(ByVal ocx As Boolean, ByVal nm As String)
    UseOcx = ocx
    EndpointName = nm
    Set mTrace = New Collection
    If ocx Then
        Set mSock = New Winsock
    Else
        Set mSock = New_Winsock2
    End If
End Sub

'------------------------------------------------------------------------------
' Methods
'------------------------------------------------------------------------------
Public Sub SetProtocol(ByVal p As Long)
    mSock.Protocol = p
End Sub

Public Sub SetRemote(ByVal host As String, ByVal port As Long)
    mSock.RemoteHost = host
    mSock.RemotePort = port
End Sub

Public Sub SetLocalPort(ByVal p As Long)
    mSock.LocalPort = p
End Sub

Public Sub ConnectTo(ByVal host As String, ByVal port As Long)
    mSock.Connect host, port
    StateAfterConnectCall = mSock.State
End Sub

' Lets a cell omit one of the optional arguments.
Public Sub ConnectOptional(ByVal host As Variant, ByVal port As Variant)
    mSock.Connect host, port
    StateAfterConnectCall = mSock.State
End Sub

Public Sub SetRemotePort(ByVal p As Long)
    mSock.RemotePort = p
End Sub

Public Sub ListenOn(ByVal port As Long)
    mSock.LocalPort = port
    mSock.Listen
End Sub

Public Sub BindTo(ByVal port As Long)
    mSock.LocalPort = port
    mSock.Bind
End Sub

Public Sub AcceptReq(ByVal requestID As Long)
    mSock.Accept requestID
End Sub

Public Sub SendString(ByVal s As String)
    mSock.SendData s
End Sub

Public Sub SendBytes(ByRef b() As Byte)
    mSock.SendData b
End Sub

Public Sub GetStringData()
    Dim s As String
    mSock.GetData s, vbString
    LastString = s
End Sub

Public Sub PeekStringData()
    Dim s As String
    mSock.PeekData s, vbString
    LastString = s
End Sub

Public Sub GetByteData()
    Dim b() As Byte
    mSock.GetData b
    mLastBytes = b
End Sub

Public Sub PeekByteData()
    Dim b() As Byte
    mSock.PeekData b
    mLastBytes = b
End Sub

Public Sub GetByteDataMax(ByVal maxLen As Long)
    Dim b() As Byte
    mSock.GetData b, vbArray + vbByte, maxLen
    mLastBytes = b
End Sub

Public Sub PeekByteDataMax(ByVal maxLen As Long)
    Dim b() As Byte
    mSock.PeekData b, vbArray + vbByte, maxLen
    mLastBytes = b
End Sub

Public Sub CloseSock()
    mSock.Close
    IsClosed = True
End Sub

' Abrupt close: sets SO_LINGER(1,0) so the peer sees a reset, not a FIN.
Public Sub KillSock()
    modTestUtil.SetLingerZero mSock.SocketHandle
    mSock.Close
    IsClosed = True
End Sub

Public Sub ShrinkReceiveBuffer(ByVal bytes As Long)
    modTestUtil.SetRecvBuf mSock.SocketHandle, bytes
End Sub

Public Sub ShrinkSendBuffer(ByVal bytes As Long)
    modTestUtil.SetSendBuf mSock.SocketHandle, bytes
End Sub

Public Sub ShutdownSend()
    modTestUtil.ShutdownSend mSock.SocketHandle
End Sub

Public Function GetAny(ByVal vType As Variant, ByVal vMax As Variant) As Variant
    Dim v As Variant
    mSock.GetData v, vType, vMax
    GetAny = v
End Function

Public Function PeekAny(ByVal vType As Variant, ByVal vMax As Variant) As Variant
    Dim v As Variant
    mSock.PeekData v, vType, vMax
    PeekAny = v
End Function

Public Sub SendVariant(ByVal v As Variant)
    mSock.SendData v
End Sub

'------------------------------------------------------------------------------
' Capture
'------------------------------------------------------------------------------
Public Sub CaptureReset(ByVal limit As Long)
    CaptureMode = True
    CaptureLimit = limit
    CaptureSize = 0
    CaptureOverflow = False
    ReDim mCapture(0 To limit - 1)
End Sub

Public Sub CaptureStop()
    CaptureMode = False
End Sub

Public Function CaptureDone() As Boolean
    CaptureDone = (CaptureSize >= CaptureLimit)
End Function

Public Sub CopyCapture(ByRef b() As Byte)
    b = mCapture
End Sub

Public Sub AboutBox()
    mSock.AboutBox
End Sub

Public Function LateGet(ByVal prop As String) As Variant
    LateGet = CallByName(mSock, prop, VbGet)
End Function

Public Function LateLet(ByVal prop As String, ByVal val As Variant)
    CallByName mSock, prop, VbLet, val
End Function

'------------------------------------------------------------------------------
' Property snapshots
'------------------------------------------------------------------------------
Public Property Get ProtocolValue() As Long
    ProtocolValue = mSock.Protocol
End Property

Public Property Get StateValue() As Long
    StateValue = mSock.State
End Property

Public Property Get SocketHandleValue() As Long
    SocketHandleValue = mSock.SocketHandle
End Property

Public Property Get LocalPortValue() As Long
    LocalPortValue = mSock.LocalPort
End Property

Public Property Get LocalIPValue() As String
    LocalIPValue = mSock.LocalIP
End Property

Public Property Get LocalHostNameValue() As String
    LocalHostNameValue = mSock.LocalHostName
End Property

Public Property Get RemotePortValue() As Long
    RemotePortValue = mSock.RemotePort
End Property

Public Property Get RemoteHostValue() As String
    RemoteHostValue = mSock.RemoteHost
End Property

Public Property Get RemoteHostIPValue() As String
    RemoteHostIPValue = mSock.RemoteHostIP
End Property

Public Property Get BytesReceivedValue() As Long
    BytesReceivedValue = mSock.BytesReceived
End Property

' Snapshot of the public property surface.
' Note: MSWINSCK.OCX 6.1.97.82 raises an access violation (0xC0000005) when
' BytesReceived is read on a socket that has already been closed, so that
' single property is skipped on closed sockets for both implementations.
Public Function SnapshotText() As String
    Dim s As String
    Dim st As Long
    On Error Resume Next
    Err.Clear
    st = mSock.State
    s = "Protocol=" & mSock.Protocol & ";State=" & st & _
        ";Handle=" & mSock.SocketHandle & _
        ";LocalPort=" & mSock.LocalPort & _
        ";LocalIP=" & mSock.LocalIP & _
        ";LocalHost=" & mSock.LocalHostName & _
        ";RemotePort=" & mSock.RemotePort & _
        ";RemoteHost=" & mSock.RemoteHost & _
        ";RemoteIP=" & mSock.RemoteHostIP
    If st <> sckClosed Then
        s = s & ";BytesReceived=" & mSock.BytesReceived
    Else
        s = s & ";BytesReceived=<closed>"
    End If
    If Err.Number <> 0 Then
        s = s & ";ERR=" & Err.Number
        Err.Clear
    End If
    SnapshotText = s
End Function

'------------------------------------------------------------------------------
' Trace
'------------------------------------------------------------------------------
Public Sub ClearTrace()
    Set mTrace = New Collection
End Sub

Public Sub LogTrace(ByVal s As String)
    If mTrace Is Nothing Then Set mTrace = New Collection
    LastEventSeconds = NowSeconds()
    mTrace.Add s
End Sub

Public Function TraceText() As String
    Dim i As Long
    Dim s As String
    If mTrace Is Nothing Then Exit Function
    For i = 1 To mTrace.Count
        If i > 1 Then s = s & vbCrLf
        s = s & CStr(mTrace.Item(i))
    Next
    TraceText = s
End Function

Public Function TraceCount() As Long
    If mTrace Is Nothing Then Exit Function
    TraceCount = mTrace.Count
End Function

'------------------------------------------------------------------------------
' Bytes
'------------------------------------------------------------------------------
Public Sub CopyLastBytes(ByRef b() As Byte)
    b = mLastBytes
End Sub

Public Function LastByteCount() As Long
    On Error Resume Next
    LastByteCount = UBound(mLastBytes) - LBound(mLastBytes) + 1
    If Err.Number <> 0 Then LastByteCount = 0
End Function

Public Sub ResetFlags()
    DoneConnect = False
    DoneData = False
    DoneClose = False
    DoneError = False
    DataCount = 0
    LastDataSize = 0
    LastString = ""
    TotalReceived = 0
    LastErrorNo = 0
    LastErrorText = ""
    Erase mLastBytes
End Sub

'------------------------------------------------------------------------------
' Events
'------------------------------------------------------------------------------
Private Sub mSock_Connect()
    DoneConnect = True
    LogTrace EndpointName & "|Connect|State=" & mSock.State
End Sub

Private Sub mSock_DataArrival(ByVal bytesTotal As Long)
    Dim b() As Byte
    Dim n As Long

    DataCount = DataCount + 1
    LastDataSize = bytesTotal
    LogTrace EndpointName & "|DataArrival|bytes=" & bytesTotal & "|BytesReceived=" & mSock.BytesReceived

    If CaptureMode Then
        On Error Resume Next
        Err.Clear
        mSock.GetData mLastBytes, vbArray + vbByte
        If Err.Number <> 0 Then
            Err.Clear
            mSock.GetData mLastBytes
        End If
        Err.Clear
        On Error GoTo 0
        n = LastByteCount()
        If n > 0 Then
            If CaptureSize + n <= CaptureLimit Then
                MemoryCopy mCapture(CaptureSize), mLastBytes(0), n
                CaptureSize = CaptureSize + n
            Else
                CaptureOverflow = True
            End If
        End If
        If CaptureSize >= CaptureLimit Then DoneData = True
        If CloseOnData Then CloseSock
        Exit Sub
    End If

    If CountOnly Then
        On Error Resume Next
        Err.Clear
        mSock.GetData b
        If Err.Number = 0 Then
            If UBound(b) >= LBound(b) Then
                n = UBound(b) - LBound(b) + 1
                TotalReceived = TotalReceived + n
            End If
            If TotalReceived >= ExpectedBytes Then DoneData = True
        End If
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If

    DoneData = True
    If CloseOnData Then
        CloseSock
        Exit Sub
    End If
    If AutoRead Then
        ' Read the payload inside the event, which is the documented MSWinsock
        ' usage. MSWINSCK.OCX discards unread UDP payloads once the event ends.
        On Error Resume Next
        Err.Clear
        mSock.GetData mLastBytes, vbArray + vbByte
        If Err.Number = 0 Then
            LastString = BytesToText(mLastBytes)
        Else
            Err.Clear
            mSock.GetData LastString, vbString
        End If
        Err.Clear
        On Error GoTo 0
        If Echo Then
            On Error Resume Next
            mSock.SendData mLastBytes
            Err.Clear
            On Error GoTo 0
        End If
        Exit Sub
    End If
    If Echo Then
        On Error Resume Next
        Err.Clear
        mSock.GetData b
        If Err.Number = 0 Then mSock.SendData b
        Err.Clear
        On Error GoTo 0
    End If
End Sub

Private Sub mSock_ConnectionRequest(ByVal requestID As Long)
    LogTrace EndpointName & "|ConnectionRequest"
    LastRequestID = requestID
    If RefuseAccept Then Exit Sub
    If ManualAccept Then Exit Sub
    modTestEngine.AcceptNew Me, requestID
End Sub

Private Sub mSock_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
    LogTrace EndpointName & "|SendProgress|sent=" & bytesSent & "|remaining=" & bytesRemaining
End Sub

Private Sub mSock_SendComplete()
    LogTrace EndpointName & "|SendComplete"
End Sub

Private Sub mSock_Close()
    DoneClose = True
    LogTrace EndpointName & "|Close"
End Sub

Private Sub mSock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    DoneError = True
    LastErrorNo = Number
    LastErrorText = Description
    LastErrorSource = Source
    LastErrorHelp = HelpFile & "|" & HelpContext
    CancelDisplay = True
    ' Source is deliberately not part of the compared trace: the OCX reports its
    ' own module path there, which is machine specific.
    LogTrace EndpointName & "|Error|Number=" & Number & "|Scode=" & Scode
End Sub

