Attribute VB_Name = "modMatrixCells2"
Option Explicit

'==============================================================================
' Matrix cells: transfer, timing, abort and concurrency families.
'==============================================================================

Private Sub SendPattern(ByRef ep As clsEndpoint, ByVal n As Long)
    Dim b() As Byte
    ReDim b(0 To n - 1)
    modTestUtil.FillPattern b, n, modTestUtil.PAT_RANDOM
    ep.SendBytes b
End Sub

Private Sub FillBytes(ByRef b() As Byte, ByVal n As Long)
    ReDim b(0 To n - 1)
    modTestUtil.FillPattern b, n, modTestUtil.PAT_RANDOM
End Sub

'------------------------------------------------------------------------------
' send:<variant>
'------------------------------------------------------------------------------
Public Sub ExecSend(ByVal v As String)
    Dim ep As clsEndpoint
    Dim peer As clsEndpoint
    Dim conn As clsEndpoint
    Dim e As Long
    Dim b() As Byte
    Dim sVar As Variant

    Select Case v
    Case "closed"
        Set ep = modMatrixCells.MxFixture("closed")
    Case "connecting"
        Set ep = modMatrixCells.MxFixture("connecting")
    Case "listening"
        Set ep = modMatrixCells.MxFixture("listening")
    Case "udpNoRemoteHost", "udpNoRemotePort", "udp65507", "udp65508"
        Set ep = modMatrixCells.MxEp("subject")
        ep.SetProtocol sckUDPProtocol
        ep.BindTo modMatrix.MxNextPort()
        modMatrix.MxWaitState ep, sckOpen, 2000
        If v = "udpNoRemoteHost" Then
            ep.SetRemotePort 9999
        ElseIf v = "udpNoRemotePort" Then
            ep.SetRemote "127.0.0.1", 0
        Else
            Set peer = modMatrixCells.MxEp("rcv")
            peer.SetProtocol sckUDPProtocol
            peer.AutoRead = True
            peer.BindTo modMatrix.MxNextPort()
            modMatrix.MxWaitState peer, sckOpen, 2000
            ep.SetRemote "127.0.0.1", peer.LocalPortValue
            Set modMatrix.gMxServer = peer
        End If
    Case Else
        Set ep = modMatrixCells.MxFixture("connected")
        Set conn = modMatrix.gMxConn
    End Select
    Set modMatrix.gMxSubject = ep

    On Error Resume Next
    Err.Clear
    Select Case v
    Case "emptyStr":   ep.SendString ""
    Case "emptyArr":   ep.SendBytes b
    Case "one":        FillBytes b, 1: ep.SendBytes b
    Case "64k":        SendPattern ep, 65536
    Case "1m":         SendPattern ep, 1048576
    Case "unicode":    ep.SendString "caf" & Chr$(233) & Chr$(255) & "x"
    Case "longs":      sVar = Array(1&, 2&, 3&): ep.SendVariant sVar
    Case "null":       sVar = Null: ep.SendVariant sVar
    Case "nothing":    Set sVar = Nothing: ep.SendVariant sVar
    Case "closed", "connecting", "listening": ep.SendString "x"
    Case "udpNoRemoteHost", "udpNoRemotePort": ep.SendString "x"
    Case "udp65507":   SendPattern ep, 65507
    Case "udp65508":   SendPattern ep, 65508
    Case "afterReset"
        conn.CloseSock
        modMatrix.MxSettle 300
        ep.SendString "after-reset"
    Case "whileGetPending"
        ep.SendString "sent-while-unread"
    End Select
    e = Err.Number
    Err.Clear
    On Error GoTo 0

    modMatrix.MxObserve "err", CStr(e)
    MxRecordState "state"
    ' SendProgress/SendComplete arrive through the message loop.
    modMatrix.MxSettle 300
    MxRecordEvents
End Sub

'------------------------------------------------------------------------------
' get:<variant>
'------------------------------------------------------------------------------
Public Sub ExecGet(ByVal v As String)
    Dim ep As clsEndpoint
    Dim peer As clsEndpoint
    Dim conn As clsEndpoint
    Dim e As Long
    Dim res As Variant
    Dim b() As Byte
    Dim n As Long

    If v = "closed" Then
        Set ep = modMatrixCells.MxFixture("closed")
        Set modMatrix.gMxSubject = ep
    ElseIf v = "twoDatagrams" Then
        Set ep = modMatrixCells.MxEp("rcv")
        ep.SetProtocol sckUDPProtocol
        ep.AutoRead = True
        ep.BindTo modMatrix.MxNextPort()
        modMatrix.MxWaitState ep, sckOpen, 2000
        Set peer = modMatrixCells.MxEp("snd")
        peer.SetProtocol sckUDPProtocol
        peer.SetRemote "127.0.0.1", ep.LocalPortValue
        FillBytes b, 8: peer.SendBytes b
        FillBytes b, 16: peer.SendBytes b
        modMatrix.MxWaitData ep, 3000
        Set modMatrix.gMxSubject = ep
        Set modMatrix.gMxServer = peer
    ElseIf v = "inDataArrival" Then
        Set ep = modMatrixCells.MxFixture("connected")
        Set modMatrix.gMxSubject = ep
        modMatrix.gMxConn.AutoRead = True
    Else
        Set ep = modMatrixCells.MxFixture("connected")
        Set modMatrix.gMxSubject = ep
        modMatrix.gMxConn.Echo = False
    End If

    If (v <> "nodata") And (v <> "closed") And (v <> "inDataArrival") Then
        ep.ResetFlags
        ep.SendString "0123456789ABCDEF"
        modMatrix.MxWaitData ep, 3000
    ElseIf v = "inDataArrival" Then
        ep.ResetFlags
        ep.SendString "0123456789ABCDEF"
        modMatrix.MxSettle 400
    End If

    On Error Resume Next
    Err.Clear
    Select Case v
    Case "nodata":     res = ep.GetAny(vbString, Empty)
    Case "closed":     res = ep.GetAny(vbString, Empty)
    Case "maxlen0":    res = ep.GetAny(vbArray + vbByte, CLng(0))
    Case "maxlen1":    res = ep.GetAny(vbArray + vbByte, CLng(1))
    Case "maxlenHuge": res = ep.GetAny(vbArray + vbByte, CLng(1000000))
    Case "maxlenNeg":  res = ep.GetAny(vbArray + vbByte, CLng(-1))
    Case "typeString": res = ep.GetAny(vbString, Empty)
    Case "typeByte":   res = ep.GetAny(vbByte, Empty)
    Case "typeArray":  res = ep.GetAny(vbArray + vbByte, Empty)
    Case "typeLong":   res = ep.GetAny(vbLong, Empty)
    Case "twoDatagrams"
        res = ep.PeekAny(vbArray + vbByte, Empty)
        If IsArray(res) Then modMatrix.MxObserve "peek", CStr(UBound(res) - LBound(res) + 1)
        res = ep.GetAny(vbArray + vbByte, Empty)
    Case "peekThenGet"
        res = ep.PeekAny(vbString, Empty)
        modMatrix.MxObserve "peeked", CStr(Len(CStr(res)))
        res = ep.GetAny(vbString, Empty)
    Case "inDataArrival"
        modMatrix.MxObserve "autoread", CStr(Len(ep.LastString))
        res = ep.GetAny(vbString, Empty)
    End Select
    e = Err.Number
    Err.Clear
    On Error GoTo 0

    modMatrix.MxObserve "err", CStr(e)
    If IsArray(res) Then
        On Error Resume Next
        n = UBound(res) - LBound(res) + 1
        On Error GoTo 0
        modMatrix.MxObserve "got", "array:" & n
    ElseIf IsNull(res) Then
        modMatrix.MxObserve "got", "null"
    Else
        modMatrix.MxObserve "got", "text:" & CStr(res)
    End If
    MxRecordState "state"
    MxRecordEvents
End Sub

'------------------------------------------------------------------------------
' close:<variant>
'------------------------------------------------------------------------------
Public Sub ExecClose(ByVal v As String)
    Dim ep As clsEndpoint
    Dim peer As clsEndpoint
    Dim e As Long

    Select Case v
    Case "closed"
        Set ep = modMatrixCells.MxFixture("closed")
    Case "udp"
        Set ep = modMatrixCells.MxFixture("open")
    Case "listeningPending"
        Set ep = modMatrixCells.MxEp("subject")
        ep.RefuseAccept = True
        ep.ListenOn modMatrix.MxNextPort()
        modMatrix.MxWaitState ep, sckListening, 3000
        Set peer = modMatrixCells.MxEp("pending")
        peer.ConnectTo "127.0.0.1", ep.LocalPortValue
        modMatrix.MxWaitConnect peer, 3000
        modMatrix.MxSettle 300
        Set modMatrix.gMxServer = peer
    Case "connecting"
        Set ep = modMatrixCells.MxFixture("connecting")
    Case "connected"
        Set ep = modMatrixCells.MxFixture("connected")
    Case "midSend"
        Set ep = modMatrixCells.MxFixture("connected")
        ep.ShrinkSendBuffer 2048
        SendPattern ep, 262144
    Case "midResolve"
        Set ep = modMatrixCells.MxEp("subject")
        ep.ConnectTo "localhost", modMatrix.MxNextPort()
    Case "inDataArrival"
        Set ep = modMatrixCells.MxFixture("connected")
        modMatrix.gMxConn.CloseOnData = True
    Case Else
        Set ep = modMatrixCells.MxFixture("connected")
    End Select
    Set modMatrix.gMxSubject = ep

    On Error Resume Next
    Err.Clear
    Select Case v
    Case "twice"
        ep.CloseSock
        ep.CloseSock
    Case "reuseAfterClose"
        ep.CloseSock
        ep.ConnectTo "127.0.0.1", modMatrix.MxNextPort()
    Case "inDataArrival"
        ep.SendString "x"
        modMatrix.MxSettle 500
    Case Else
        ep.CloseSock
    End Select
    e = Err.Number
    Err.Clear
    On Error GoTo 0

    modMatrix.MxObserve "err", CStr(e)
    modMatrix.MxSettle 350
    MxRecordState "state"
    modMatrix.MxObserve "closeFlag", CStr(ep.DoneClose)
    MxRecordEvents
End Sub

'------------------------------------------------------------------------------
' dns:<variant>
'------------------------------------------------------------------------------
Public Sub ExecDns(ByVal v As String)
    Dim ep As clsEndpoint
    Dim srv As clsEndpoint
    Dim e As Long
    Dim i As Long
    Dim eps() As clsEndpoint

    Select Case v
    Case "ok", "twice", "literal", "localhost"
        Set srv = modMatrixCells.MxServerOnNewPort("srv")
        Set modMatrix.gMxServer = srv
    End Select
    Set ep = modMatrixCells.MxEp("subject")
    Set modMatrix.gMxSubject = ep

    On Error Resume Next
    Err.Clear
    Select Case v
    Case "ok":           ep.ConnectTo "localhost", srv.LocalPortValue
    Case "bad":          ep.ConnectTo "no-such-host.invalid", 80
    Case "closeDuring":  ep.ConnectTo "localhost", 80: ep.CloseSock
    Case "connectDuring": ep.ConnectTo "localhost", 80: ep.ConnectTo "localhost", 81
    Case "twice":        ep.ConnectTo "localhost", srv.LocalPortValue: ep.CloseSock: ep.ConnectTo "localhost", srv.LocalPortValue
    Case "literal":      ep.ConnectTo "127.0.0.1", srv.LocalPortValue
    Case "localhost":    ep.ConnectTo "localhost", srv.LocalPortValue
    Case "empty":        ep.ConnectTo "", 80
    Case "concurrent8"
        ReDim eps(1 To 8)
        For i = 1 To 8
            Set eps(i) = modMatrixCells.MxEp("dns" & i)
            eps(i).ConnectTo "localhost", 80
        Next
        modMatrix.MxSettle 400
        For i = 1 To 8
            modMatrix.MxObserve "dns" & i, "state=" & eps(i).StateValue
        Next
    End Select
    e = Err.Number
    Err.Clear
    On Error GoTo 0

    modMatrix.MxObserve "err", CStr(e)
    modMatrix.MxSettle 400
    MxRecordState "state"
    modMatrix.MxObserve "connected", CStr(ep.DoneConnect) & "/" & ep.LastErrorNo
    MxRecordEvents
End Sub

'------------------------------------------------------------------------------
' evt:<variant>
'------------------------------------------------------------------------------
Public Sub ExecEvt(ByVal v As String)
    Dim ep As clsEndpoint
    Dim conn As clsEndpoint
    Dim srv As clsEndpoint
    Dim c1 As clsEndpoint
    Dim c2 As clsEndpoint
    Dim i As Long

    If v = "error10061" Then
        Set ep = modMatrixCells.MxEp("subject")
        Set modMatrix.gMxSubject = ep
        ep.ConnectTo "127.0.0.1", 1
        modMatrix.MxWaitError ep, 5000
        modMatrix.MxObserve "errno", CStr(ep.LastErrorNo)
        modMatrix.MxObserve "errtext", ep.LastErrorText
        MxRecordEvents
        Exit Sub
    End If

    If v = "connReqOrder" Or v = "reentrancyConnReq" Then
        Set srv = modMatrixCells.MxEp("srv")
        srv.ListenOn modMatrix.MxNextPort()
        modMatrix.MxWaitState srv, sckListening, 3000
        Set modMatrix.gMxServer = srv
        Set c1 = modMatrixCells.MxEp("c1")
        c1.ConnectTo "127.0.0.1", srv.LocalPortValue
        If v = "connReqOrder" Then
            Set c2 = modMatrixCells.MxEp("c2")
            c2.ConnectTo "127.0.0.1", srv.LocalPortValue
        End If
        modMatrix.MxSettle 500
        MxRecordState "state"
        MxRecordEvents
        Exit Sub
    End If

    Set ep = modMatrixCells.MxFixture("connected")
    Set conn = modMatrix.gMxConn
    Set modMatrix.gMxSubject = ep

    Select Case v
    Case "dataCoalescing"
        conn.Echo = False
        ep.ResetFlags
        For i = 1 To 3
            ep.SendString "chunk" & i
        Next
        modMatrix.MxSettle 500
        modMatrix.MxObserve "dataEvents", CStr(ep.DataCount)
    Case "sendProgressTcp"
        SendPattern ep, 200000
        modMatrix.MxSettle 400
    Case "sendProgressUdp"
        ExecEvtUdp
        Exit Sub
    Case "closeLocal"
        ep.CloseSock
        modMatrix.MxSettle 300
    Case "closeRemote"
        conn.CloseSock
        modMatrix.MxSettle 300
    Case "closeReset"
        conn.KillSock
        modMatrix.MxSettle 600
    Case "reentrancyData"
        conn.CloseOnData = True
        ep.SendString "close-inside"
        modMatrix.MxSettle 600
    End Select

    MxRecordState "state"
    MxRecordEvents
End Sub

Private Sub ExecEvtUdp()
    Dim r As clsEndpoint
    Dim s As clsEndpoint
    Dim b() As Byte

    Set r = modMatrixCells.MxEp("rcv")
    r.SetProtocol sckUDPProtocol
    r.AutoRead = True
    r.BindTo modMatrix.MxNextPort()
    modMatrix.MxWaitState r, sckOpen, 2000
    Set s = modMatrixCells.MxEp("snd")
    s.SetProtocol sckUDPProtocol
    s.SetRemote "127.0.0.1", r.LocalPortValue
    FillBytes b, 32
    s.SendBytes b
    modMatrix.MxSettle 300
    Set modMatrix.gMxSubject = s
    Set modMatrix.gMxServer = r
    MxRecordState "state"
    MxRecordEvents
End Sub

'------------------------------------------------------------------------------
' abort:<variant>
'------------------------------------------------------------------------------
Public Sub ExecAbort(ByVal v As String)
    Dim ep As clsEndpoint
    Dim conn As clsEndpoint
    Dim e As Long
    Dim b() As Byte

    Set ep = modMatrixCells.MxFixture("connected")
    Set conn = modMatrix.gMxConn
    Set modMatrix.gMxSubject = ep

    On Error Resume Next
    Err.Clear
    Select Case v
    Case "fin"
        conn.CloseSock
        modMatrix.MxSettle 500
    Case "rstMidSend"
        SendPattern ep, 4194304
        conn.KillSock
        modMatrix.MxSettle 700
    Case "rstMidRecv"
        conn.Echo = False
        SendPattern ep, 262144
        modMatrix.MxSettle 200
        conn.KillSock
        modMatrix.MxSettle 700
    Case "halfClose"
        conn.ShutdownSend
        modMatrix.MxSettle 300
        ep.SendString "after-half-close"
        modMatrix.MxSettle 400
    Case "partialWindow"
        conn.ShrinkReceiveBuffer 2048
        SendPattern ep, 1048576
        modMatrix.MxSettle 1500
    Case "stream16m"
        SendPattern ep, 16777216
        conn.KillSock
        modMatrix.MxSettle 900
    Case "rstThenSend"
        conn.KillSock
        modMatrix.MxSettle 300
        ep.SendString "x"
    Case "rstThenGet"
        conn.KillSock
        modMatrix.MxSettle 300
        b = ep.GetAny(vbString, Empty)
    Case "rstThenReuse"
        conn.KillSock
        modMatrix.MxSettle 300
        ep.CloseSock
        ep.ConnectTo "127.0.0.1", modMatrix.MxNextPort()
    End Select
    e = Err.Number
    Err.Clear
    On Error GoTo 0

    modMatrix.MxObserve "err", CStr(e)
    modMatrix.MxObserve "errorFlag", CStr(ep.DoneError) & "/" & ep.LastErrorNo
    modMatrix.MxObserve "closeFlag", CStr(ep.DoneClose)
    MxRecordState "state"
    MxRecordEvents
End Sub

'------------------------------------------------------------------------------
' conc:<variant>
'------------------------------------------------------------------------------
Public Sub ExecConc(ByVal v As String)
    Dim srv As clsEndpoint
    Dim i As Long
    Dim n As Long
    Dim cli() As clsEndpoint
    Dim e As Long

    On Error Resume Next
    Err.Clear
    Select Case v
    Case "clients32"
        Set srv = modMatrixCells.MxServerOnNewPort("srv")
        srv.Echo = True
        Set modMatrix.gMxServer = srv
        ReDim cli(1 To 32)
        For i = 1 To 32
            Set cli(i) = modMatrixCells.MxEp("c" & i)
            cli(i).ConnectTo "127.0.0.1", srv.LocalPortValue
            If (i Mod 8) = 0 Then modMatrix.MxWaitAccepted i, 8000
        Next
        n = 0
        For i = 1 To 32
            cli(i).ResetFlags
            cli(i).SendString "c" & i
        Next
        For i = 1 To 32
            If modMatrix.MxWaitData(cli(i), 8000) Then
                cli(i).GetStringData
                If cli(i).LastString <> "c" & i Then n = n + 1
            Else
                n = n + 1
            End If
        Next
        modMatrix.MxObserve "clients", "32 crosstalk=" & n
        modMatrix.MxObserve "accepted", CStr(modTestEngine.MxChildCount)

    Case "orderUnderLoad"
        Set srv = modMatrixCells.MxServerOnNewPort("srv")
        srv.Echo = True
        Set modMatrix.gMxServer = srv
        ReDim cli(1 To 8)
        For i = 1 To 8
            Set cli(i) = modMatrixCells.MxEp("c" & i)
            cli(i).ConnectTo "127.0.0.1", srv.LocalPortValue
        Next
        For i = 1 To 8
            modMatrix.MxWaitConnect cli(i), 8000
        Next
        modMatrix.MxWaitAccepted 8, 8000
        n = 0
        For i = 1 To 8
            cli(i).ResetFlags
            cli(i).SendString "seq" & i
        Next
        For i = 1 To 8
            If modMatrix.MxWaitData(cli(i), 8000) Then
                cli(i).GetStringData
                If cli(i).LastString <> "seq" & i Then n = n + 1
            Else
                n = n + 1
            End If
        Next
        modMatrix.MxObserve "order", "8 mismatches=" & n

    Case "listeners3"
        Dim s(1 To 3) As clsEndpoint
        n = 0
        For i = 1 To 3
            Set s(i) = modMatrixCells.MxServerOnNewPort("s" & i)
            s(i).Echo = True
            Set cli(i) = modMatrixCells.MxEp("c" & i)
            cli(i).ConnectTo "127.0.0.1", s(i).LocalPortValue
            modMatrix.MxWaitConnect cli(i), 8000
            modMatrix.MxWaitAccepted i, 8000
            cli(i).ResetFlags
            cli(i).SendString "s" & i
            If modMatrix.MxWaitData(cli(i), 8000) Then
                cli(i).GetStringData
                If cli(i).LastString <> "s" & i Then n = n + 1
            Else
                n = n + 1
            End If
        Next
        modMatrix.MxObserve "pairs", "3 mismatches=" & n

    Case "mixed"
        Set srv = modMatrixCells.MxServerOnNewPort("srv")
        srv.Echo = True
        Set modMatrix.gMxServer = srv
        Dim r As clsEndpoint
        Dim sd As clsEndpoint
        Set r = modMatrixCells.MxEp("ur")
        r.SetProtocol sckUDPProtocol
        r.AutoRead = True
        r.BindTo modMatrix.MxNextPort()
        modMatrix.MxWaitState r, sckOpen, 2000
        Set sd = modMatrixCells.MxEp("us")
        sd.SetProtocol sckUDPProtocol
        sd.SetRemote "127.0.0.1", r.LocalPortValue
        ReDim cli(1 To 1)
        Set cli(1) = modMatrixCells.MxEp("c1")
        cli(1).ConnectTo "127.0.0.1", srv.LocalPortValue
        modMatrix.MxWaitConnect cli(1), 5000
        cli(1).ResetFlags
        cli(1).SendString "tcp"
        r.ResetFlags
        sd.SendString "udp"
        modMatrix.MxSettle 700
        modMatrix.MxObserve "mixed", "tcp=" & CStr(cli(1).DoneData) & " udp=" & CStr(r.DoneData)

    Case "churn"
        Set srv = modMatrixCells.MxServerOnNewPort("srv")
        srv.Echo = True
        Set modMatrix.gMxServer = srv
        ReDim cli(1 To 1)
        For i = 1 To 100
            Set cli(1) = modMatrixCells.MxEp("churn")
            cli(1).ConnectTo "127.0.0.1", srv.LocalPortValue
            modMatrix.MxWaitConnect cli(1), 4000
            modMatrix.MxWaitAccepted i, 4000
            cli(1).CloseSock
            Set cli(1) = Nothing
        Next
        modMatrix.MxObserve "cycles", "100"

    Case "reuseSlots"
        ReDim cli(1 To 1)
        For i = 1 To 50
            Set cli(1) = modMatrixCells.MxEp("slot")
            Set cli(1) = Nothing
        Next
        modMatrix.MxObserve "slots", "50 create/release"
    End Select
    e = Err.Number
    Err.Clear
    On Error GoTo 0

    modMatrix.MxObserve "err", CStr(e)
    MxRecordEvents
End Sub
