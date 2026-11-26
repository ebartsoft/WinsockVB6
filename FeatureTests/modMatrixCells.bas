Attribute VB_Name = "modMatrixCells"
Option Explicit

'==============================================================================
' Matrix cells: fixtures, helpers and the state/property families.
' The transfer, event and abort families live in modMatrixCells2.
'==============================================================================

'------------------------------------------------------------------------------
' Fixtures
'------------------------------------------------------------------------------
Public Function MxEp(ByVal nm As String) As clsEndpoint
    modMatrix.gMxStage = "new:" & nm
    Set MxEp = modMatrix.MxNew(nm)
End Function

Public Function MxServerOnNewPort(ByVal nm As String) As clsEndpoint
    Dim s As clsEndpoint
    Set s = MxEp(nm)
    s.Echo = False
    s.ListenOn modMatrix.MxNextPort()
    modMatrix.MxWaitState s, sckListening, 4000
    Set MxServerOnNewPort = s
End Function

Public Function MxClientTo(ByRef srv As clsEndpoint, ByVal nm As String) As clsEndpoint
    Dim c As clsEndpoint
    Set c = MxEp(nm)
    c.ConnectTo "127.0.0.1", srv.LocalPortValue
    modMatrix.MxWaitConnect c, 5000
    modMatrix.MxWaitAccepted 1, 5000
    Set MxClientTo = c
End Function

' Puts the subject into one of the states the matrix asks for.
Public Function MxFixture(ByVal state As String) As clsEndpoint
    Dim ep As clsEndpoint
    Dim srv As clsEndpoint

    modMatrix.gMxStage = "fixture:" & state
    Select Case LCase$(state)
    Case "closed"
        modMatrix.gMxStage = "fx-closed"
        Set ep = MxEp("subject")
        modMatrix.gMxStage = "fx-closed-ok"

    Case "listening"
        Set ep = MxEp("subject")
        ep.ListenOn modMatrix.MxNextPort()
        modMatrix.MxWaitState ep, sckListening, 4000

    Case "connected"
        Set srv = MxServerOnNewPort("srv")
        Set ep = MxClientTo(srv, "subject")
        Set modMatrix.gMxServer = srv
        If modTestEngine.MxChildCount >= 1 Then Set modMatrix.gMxConn = modTestEngine.MxChild(1)

    Case "open"
        Set ep = MxEp("subject")
        ep.SetProtocol sckUDPProtocol
        ep.BindTo modMatrix.MxNextPort()
        modMatrix.MxWaitState ep, sckOpen, 3000

    Case "connecting"
        Set ep = MxEp("subject")
        ep.ConnectTo "10.255.255.1", 81      ' black hole: stays in sckConnecting

    Case "error"
        Set ep = MxEp("subject")
        ep.ConnectTo "127.0.0.1", 1          ' refused
        modMatrix.MxWaitError ep, 5000

    Case Else
        Set ep = MxEp("subject")
    End Select
    modMatrix.gMxStage = "fx-end"
    Set MxFixture = ep
End Function

'------------------------------------------------------------------------------
' Property access by name
'------------------------------------------------------------------------------
Public Function MxGetProp(ByRef ep As clsEndpoint, ByVal prop As String) As String
    On Error Resume Next
    Err.Clear
    Select Case LCase$(prop)
    Case "protocol":       MxGetProp = CStr(ep.ProtocolValue)
    Case "sockethandle":   MxGetProp = CStr(ep.SocketHandleValue)
    Case "localport":      MxGetProp = CStr(ep.LocalPortValue)
    Case "localip":        MxGetProp = ep.LocalIPValue
    Case "localhostname":  MxGetProp = ep.LocalHostNameValue
    Case "remoteport":     MxGetProp = CStr(ep.RemotePortValue)
    Case "remotehost":     MxGetProp = ep.RemoteHostValue
    Case "remotehostip":   MxGetProp = ep.RemoteHostIPValue
    Case "bytesreceived":  MxGetProp = CStr(ep.BytesReceivedValue)
    Case "state":          MxGetProp = CStr(ep.StateValue)
    Case Else:             MxGetProp = "<unknown>"
    End Select
    If Err.Number <> 0 Then
        MxGetProp = "raise:" & Err.Number
        Err.Clear
    End If
End Function

Public Function MxSetProp(ByRef ep As clsEndpoint, ByVal name As String, ByVal value As String) As Long
    Dim e As Long
    On Error Resume Next
    Err.Clear
    Select Case LCase$(name)
    Case "protocol":    ep.SetProtocol CLng("0" & value)
    Case "localport":   ep.SetLocalPort CLng("0" & value)
    Case "remoteport":  ep.SetRemotePort CLng("0" & value)
    Case "remotehost":  ep.SetRemote value, 0
    End Select
    e = Err.Number
    Err.Clear
    On Error GoTo 0
    MxSetProp = e
End Function

Public Sub MxRecordEvents()
    Dim s As String
    On Error Resume Next
    If Not modMatrix.gMxSubject Is Nothing Then s = modTestUtil.NormalizeTrace(modMatrix.gMxSubject.TraceText)
    If Not modMatrix.gMxServer Is Nothing Then s = s & modTestUtil.NormalizeTrace(modMatrix.gMxServer.TraceText)
    If Not modMatrix.gMxConn Is Nothing Then s = s & modTestUtil.NormalizeTrace(modMatrix.gMxConn.TraceText)
    If Len(s) = 0 Then s = "(none)"
    modMatrix.MxObserve "events", s
End Sub

Public Sub MxRecordState(ByVal label As String)
    On Error Resume Next
    If modMatrix.gMxSubject Is Nothing Then
        modMatrix.MxObserve label, "none"
    Else
        modMatrix.MxObserve label, CStr(modMatrix.gMxSubject.StateValue) & "/" & _
                                    CStr(modMatrix.gMxSubject.BytesReceivedValue)
    End If
End Sub

'------------------------------------------------------------------------------
' Dispatcher
'------------------------------------------------------------------------------
Public Sub ExecuteCell(ByVal cellID As String, ByVal ocx As Boolean)
    Dim parts() As String
    Dim fam As String

    parts = Split(cellID, ":")
    fam = LCase$(parts(0))
    Select Case fam
    Case "t":     ExecT parts
    Case "pg":    ExecPG parts
    Case "ps":    ExecPS parts
    Case "conn":  ExecConn parts
    Case "listen": ExecListen parts
    Case "accept": ExecAccept parts
    Case "bind":  ExecBind parts
    Case "send":  ExecSend part(parts, 1)
    Case "get":   ExecGet part(parts, 1)
    Case "close": ExecClose part(parts, 1)
    Case "dns":   ExecDns part(parts, 1)
    Case "evt":   ExecEvt part(parts, 1)
    Case "abort": ExecAbort part(parts, 1)
    Case "conc":  ExecConc part(parts, 1)
    Case Else
        modMatrix.MxObserve "unknown", cellID
    End Select
End Sub

Public Function part(ByRef parts() As String, ByVal i As Long) As String
    If i <= UBound(parts) Then part = parts(i)
End Function

Private Sub ExecT(ByRef parts() As String)
    Dim ep As clsEndpoint
    modMatrix.MxObserve "a", "1"
    modMatrix.gMxStage = "t-new"
    Set ep = MxEp("t")
    If ep Is Nothing Then modMatrix.MxObserve "b", "nothing" Else modMatrix.MxObserve "b", "object"
    modMatrix.MxObserve "c", CStr(modTestEngine.MxChildCount)
    modMatrix.gMxStage = "t-prop"
    modMatrix.MxObserve "d", MxGetProp(ep, "State")
    modMatrix.MxObserve "e", CStr(ep.StateValue)
    modMatrix.gMxStage = "t-done"
End Sub

'------------------------------------------------------------------------------
' pg:<state>:<Prop>
'------------------------------------------------------------------------------
Private Sub ExecPG(ByRef parts() As String)
    Dim ep As clsEndpoint
    modMatrix.gMxStage = "pg1"
    Set ep = MxFixture(part(parts, 1))
    modMatrix.gMxStage = "pg2"
    Set modMatrix.gMxSubject = ep
    modMatrix.gMxStage = "pg3"
    modMatrix.MxObserve "value", modMatrix.MxMask(part(parts, 2), MxGetProp(ep, part(parts, 2)))
    MxRecordState "state"
    MxRecordEvents
End Sub

'------------------------------------------------------------------------------
' ps:<state>:<Prop=Value>
'------------------------------------------------------------------------------
Private Sub ExecPS(ByRef parts() As String)
    Dim ep As clsEndpoint
    Dim kv As String
    Dim p As Long
    Dim name As String
    Dim value As String
    Dim e As Long

    modMatrix.gMxStage = "pg1"
    Set ep = MxFixture(part(parts, 1))
    modMatrix.gMxStage = "pg2"
    Set modMatrix.gMxSubject = ep
    modMatrix.gMxStage = "pg3"
    kv = part(parts, 2)
    p = InStr(1, kv, "=")
    If p > 0 Then
        name = Left$(kv, p - 1)
        value = Mid$(kv, p + 1)
    Else
        name = kv
    End If
    modMatrix.MxObserve "set", name & "=" & value
    e = MxSetProp(ep, name, value)
    modMatrix.MxObserve "err", CStr(e)
    modMatrix.MxObserve "after", modMatrix.MxMask(name, MxGetProp(ep, name))
    MxRecordState "state"
    MxRecordEvents
End Sub

'------------------------------------------------------------------------------
' conn:<variant>
'------------------------------------------------------------------------------
Private Sub ExecConn(ByRef parts() As String)
    Dim ep As clsEndpoint
    Dim e As Long
    Dim v As String

    v = part(parts, 1)
    Select Case v
    Case "whileConnected", "whileListening", "afterClose"
        Set ep = MxFixture(IIf(v = "whileListening", "listening", "connected"))
        Set modMatrix.gMxSubject = ep
    Case Else
        Set ep = MxEp("subject")
        Set modMatrix.gMxSubject = ep
    End Select

    On Error Resume Next
    Err.Clear
    Select Case v
    Case "nohost":        ep.ConnectTo "", 0
    Case "hostonly":      ep.ConnectOptional "127.0.0.1", Empty
    Case "portonly":      ep.ConnectOptional Empty, 80
    Case "port0":         ep.ConnectTo "127.0.0.1", 0
    Case "port65536":     ep.ConnectTo "127.0.0.1", 65536
    Case "refused":       ep.ConnectTo "127.0.0.1", 1
    Case "blackhole":     ep.ConnectTo "10.255.255.1", 81
    Case "name":          ep.ConnectTo "localhost", modMatrix.MxNextPort()
    Case "ownlistener":   ep.ListenOn modMatrix.MxNextPort(): ep.ConnectTo "127.0.0.1", ep.LocalPortValue
    Case "whileConnected": ep.ConnectTo "127.0.0.1", modMatrix.MxNextPort()
    Case "whileListening": ep.ConnectTo "127.0.0.1", modMatrix.MxNextPort()
    Case "afterClose":    ep.CloseSock: ep.ConnectTo "127.0.0.1", modMatrix.MxNextPort()
    Case "emptyName":     ep.ConnectTo "", 80
    End Select
    e = Err.Number
    Err.Clear
    On Error GoTo 0

    modMatrix.MxObserve "err", CStr(e)
    If v = "ownlistener" Or v = "name" Then modMatrix.MxSettle 300
    MxRecordState "state"
    modMatrix.MxObserve "connectedFlag", CStr(ep.DoneConnect)
    modMatrix.MxObserve "errorFlag", CStr(ep.DoneError) & "/" & ep.LastErrorNo
    MxRecordEvents
End Sub

'------------------------------------------------------------------------------
' listen:<variant>
'------------------------------------------------------------------------------
Private Sub ExecListen(ByRef parts() As String)
    Dim ep As clsEndpoint
    Dim srv As clsEndpoint
    Dim e As Long
    Dim v As String
    Dim p As Long

    v = part(parts, 1)
    On Error Resume Next
    Err.Clear
    Select Case v
    Case "port0"
        Set ep = MxEp("subject")
        Set modMatrix.gMxSubject = ep
        ep.ListenOn 0
    Case "inuse"
        Set srv = MxServerOnNewPort("srv")
        Set modMatrix.gMxServer = srv
        Set ep = MxEp("subject")
        Set modMatrix.gMxSubject = ep
        ep.ListenOn srv.LocalPortValue
    Case "twice"
        Set ep = MxEp("subject")
        Set modMatrix.gMxSubject = ep
        p = modMatrix.MxNextPort()
        ep.ListenOn p
        modMatrix.MxWaitState ep, sckListening, 3000
        ep.ListenOn p
    Case "whileConnected"
        Set ep = MxFixture("connected")
        Set modMatrix.gMxSubject = ep
        ep.ListenOn modMatrix.MxNextPort()
    Case "afterBind"
        Set ep = MxEp("subject")
        Set modMatrix.gMxSubject = ep
        ep.BindTo modMatrix.MxNextPort()
        ep.ListenOn modMatrix.MxNextPort()
    Case "closeWithPending"
        Set ep = MxEp("subject")
        ep.RefuseAccept = True
        Set modMatrix.gMxSubject = ep
        p = modMatrix.MxNextPort()
        ep.ListenOn p
        modMatrix.MxWaitState ep, sckListening, 3000
        Dim cli As clsEndpoint
        Set cli = MxEp("pending")
        cli.ConnectTo "127.0.0.1", p
        modMatrix.MxWaitConnect cli, 3000
        modMatrix.MxSettle 300
        ep.CloseSock
    End Select
    e = Err.Number
    Err.Clear
    On Error GoTo 0

    modMatrix.MxObserve "err", CStr(e)
    If Not (v = "closeWithPending") Then modMatrix.MxWaitState ep, sckListening, 1000
    MxRecordState "state"
    MxRecordEvents
End Sub

'------------------------------------------------------------------------------
' accept:<variant>
'------------------------------------------------------------------------------
Private Sub ExecAccept(ByRef parts() As String)
    Dim ep As clsEndpoint
    Dim cli As clsEndpoint
    Dim e As Long
    Dim v As String
    Dim rid As Long
    Dim p As Long

    v = part(parts, 1)
    Set ep = MxEp("subject")
    ep.ManualAccept = True
    Set modMatrix.gMxSubject = ep
    p = modMatrix.MxNextPort()
    ep.ListenOn p
    modMatrix.MxWaitState ep, sckListening, 3000

    If v <> "notListening" Then
        Set cli = MxEp("pending")
        cli.ConnectTo "127.0.0.1", p
        modMatrix.MxWaitConnect cli, 3000
        modMatrix.MxSettle 300
        rid = ep.LastRequestID
    End If

    On Error Resume Next
    Err.Clear
    Select Case v
    Case "valid":       ep.AcceptReq rid
    Case "bogus":       ep.AcceptReq 999999
    Case "duplicate":   ep.AcceptReq rid: ep.AcceptReq rid
    Case "foreign":     ep.AcceptReq rid + 1
    Case "afterClose":  ep.CloseSock: ep.AcceptReq rid
    Case "notListening": ep.AcceptReq 1
    End Select
    e = Err.Number
    Err.Clear
    On Error GoTo 0

    modMatrix.MxObserve "requestID", IIf(rid = 0, "0", "<id>")
    modMatrix.MxObserve "err", CStr(e)
    MxRecordState "state"
    modMatrix.MxObserve "children", CStr(modTestEngine.MxChildCount)
    MxRecordEvents
End Sub

'------------------------------------------------------------------------------
' bind:<variant>
'------------------------------------------------------------------------------
Private Sub ExecBind(ByRef parts() As String)
    Dim ep As clsEndpoint
    Dim srv As clsEndpoint
    Dim e As Long
    Dim v As String
    Dim p As Long

    v = part(parts, 1)
    On Error Resume Next
    Err.Clear
    Select Case v
    Case "port0"
        Set ep = MxEp("subject"): ep.SetProtocol sckUDPProtocol: ep.BindTo 0
        Set modMatrix.gMxSubject = ep
    Case "specific"
        Set ep = MxEp("subject"): ep.SetProtocol sckUDPProtocol: ep.BindTo modMatrix.MxNextPort()
        Set modMatrix.gMxSubject = ep
    Case "inuse"
        Set srv = MxEp("holder"): srv.SetProtocol sckUDPProtocol: srv.BindTo modMatrix.MxNextPort()
        modMatrix.MxWaitState srv, sckOpen, 2000
        Set ep = MxEp("subject"): ep.SetProtocol sckUDPProtocol: ep.BindTo srv.LocalPortValue
        Set modMatrix.gMxSubject = ep
        Set modMatrix.gMxServer = srv
    Case "twice"
        Set ep = MxEp("subject"): ep.SetProtocol sckUDPProtocol
        p = modMatrix.MxNextPort()
        ep.BindTo p
        modMatrix.MxWaitState ep, sckOpen, 2000
        ep.BindTo p
        Set modMatrix.gMxSubject = ep
    Case "afterConnect"
        Set ep = MxFixture("connected")
        ep.SetProtocol sckUDPProtocol
        ep.BindTo modMatrix.MxNextPort()
        Set modMatrix.gMxSubject = ep
    Case "localIP"
        Set ep = MxEp("subject"): ep.SetProtocol sckUDPProtocol
        ep.SetLocalPort modMatrix.MxNextPort()
        ep.BindTo ep.LocalPortValue
        Set modMatrix.gMxSubject = ep
    Case "onTCP"
        Set ep = MxEp("subject"): ep.BindTo modMatrix.MxNextPort()
        Set modMatrix.gMxSubject = ep
    End Select
    e = Err.Number
    Err.Clear
    On Error GoTo 0

    modMatrix.MxObserve "err", CStr(e)
    MxRecordState "state"
    MxRecordEvents
End Sub
