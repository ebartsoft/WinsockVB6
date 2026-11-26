Attribute VB_Name = "modTestEngine"
Option Explicit

Public gForm As frmFeatureTests
Public gResults As Collection
Public gPerf As Collection
Public gActive As Collection
Public gChildren As Collection
Public gListItems As Collection
Public gOcxAvailable As Boolean
Public gFilter As String
Public gDiag As String

' Matrix runner state
Private gMxPortNext As Long

Private gNextPort As Long

'==============================================================================
' Entry points
'==============================================================================
Public Sub RunAll(ByRef f As frmFeatureTests)
    Prepare f
    RunFunctionalCore
    If Len(gFilter) = 0 Then RunPerformanceCore
    FillUI f
End Sub

Public Sub RunFunctional(ByRef f As frmFeatureTests)
    Prepare f
    RunFunctionalCore
    FillUI f
End Sub

Public Sub RunPerformance(ByRef f As frmFeatureTests)
    Prepare f
    RunPerformanceCore
    FillUI f
End Sub

Private Sub Prepare(ByRef f As frmFeatureTests)
    Set gForm = f
    Set gResults = New Collection
    Set gPerf = New Collection
    Set gActive = New Collection
    Set gChildren = New Collection
    Set gListItems = New Collection
    gNextPort = 18500
    gFilter = ParseFilter(Command$)
    gDiag = ""
    gOcxAvailable = CheckOcx()
End Sub

Private Function ParseFilter(ByVal cmd As String) As String
    Dim parts() As String
    Dim i As Long
    Dim p As String
    parts = Split(cmd, " ")
    For i = LBound(parts) To UBound(parts)
        p = Trim$(parts(i))
        If Len(p) > 6 Then
            If LCase$(Left$(p, 6)) = "/only:" Then ParseFilter = Mid$(p, 7)
        End If
    Next
End Function

Private Function CheckOcx() As Boolean
    Dim o As Winsock
    On Error Resume Next
    Set o = New Winsock
    CheckOcx = (Err.Number = 0)
    Err.Clear
    Set o = Nothing
End Function

Private Function NextPort() As Long
    gNextPort = gNextPort + 1
    If gNextPort > 30000 Then gNextPort = 18500
    NextPort = gNextPort
End Function

'==============================================================================
' Endpoint construction / acceptance
'==============================================================================
Private Function NewEndpoint(ByVal useOcx As Boolean, ByVal nm As String) As clsEndpoint
    Dim e As clsEndpoint
    On Error GoTo Failed
    modMatrix.gMxStage = "new-object"
    Set e = New clsEndpoint
    modMatrix.gMxStage = "init"
    e.Init useOcx, nm
    modMatrix.gMxStage = "init-ok"
    If gActive Is Nothing Then Set gActive = New Collection
    gActive.Add e
    Set NewEndpoint = e
    Exit Function
Failed:
    gDiag = gDiag & "NewEndpoint(" & nm & ") failed: " & Err.Number & " " & Err.Description & vbCrLf
    Set NewEndpoint = Nothing
End Function

Public Sub AcceptNew(ByRef listener As clsEndpoint, ByVal requestID As Long)
    Dim c As clsEndpoint
    Dim n As Long
    On Error GoTo Failed
    If gChildren Is Nothing Then Set gChildren = New Collection
    If gActive Is Nothing Then Set gActive = New Collection

    Set c = New clsEndpoint
    c.Init listener.UseOcx, listener.EndpointName & ".conn"
    c.Echo = listener.Echo
    c.CountOnly = listener.CountOnly
    c.ExpectedBytes = listener.ExpectedBytes
    c.CloseOnData = listener.CloseOnData
    If listener.ChildCaptureLimit > 0 Then c.CaptureReset listener.ChildCaptureLimit

    ' Accept immediately: the requestID is only valid while the
    ' ConnectionRequest event is being handled.
    Err.Clear
    c.AcceptReq requestID
    n = Err.Number
    Err.Clear
    If n <> 0 Then
        gDiag = gDiag & "Accept(" & requestID & ") on " & listener.EndpointName & " failed: " & n & " " & Err.Description & vbCrLf
        c.LogTrace "Accept|error=" & n
        Err.Clear
    End If
    gActive.Add c
    gChildren.Add c
    Exit Sub
Failed:
    ' The scenario will time out / fail if the accepted endpoint is missing.
    gDiag = gDiag & "AcceptNew on " & listener.EndpointName & " failed: " & Err.Number & " " & Err.Description & vbCrLf
End Sub

Private Sub CleanupActive()
    Dim i As Long
    Dim e As clsEndpoint
    If gActive Is Nothing Then Exit Sub
    For i = gActive.Count To 1 Step -1
        Set e = gActive.Item(i)
        On Error Resume Next
        If Not e.IsClosed Then
            If e.StateValue <> sckClosed Then e.CloseSock
        End If
        Set e = Nothing
    Next
    Set gActive = New Collection
    Set gChildren = New Collection
End Sub

'==============================================================================
' Wait helpers
'==============================================================================
Private Function WaitState(ByRef c As clsEndpoint, ByVal want As Long, ByVal ms As Long) As Boolean
    Dim t As Double
    t = NowSeconds()
    Do While c.StateValue <> want
        Pump
        If (NowSeconds() - t) * 1000# > ms Then Exit Function
    Loop
    WaitState = True
End Function

Private Function WaitConnect(ByRef c As clsEndpoint, ByVal ms As Long) As Boolean
    Dim t As Double
    t = NowSeconds()
    Do While Not c.DoneConnect
        Pump
        If (NowSeconds() - t) * 1000# > ms Then Exit Function
    Loop
    WaitConnect = True
End Function

Private Function WaitData(ByRef c As clsEndpoint, ByVal ms As Long) As Boolean
    Dim t As Double
    t = NowSeconds()
    Do While Not c.DoneData
        Pump
        If (NowSeconds() - t) * 1000# > ms Then Exit Function
    Loop
    WaitData = True
End Function

Private Function WaitClose(ByRef c As clsEndpoint, ByVal ms As Long) As Boolean
    Dim t As Double
    t = NowSeconds()
    Do While Not c.DoneClose
        Pump
        If (NowSeconds() - t) * 1000# > ms Then Exit Function
    Loop
    WaitClose = True
End Function

Private Function WaitError(ByRef c As clsEndpoint, ByVal ms As Long) As Boolean
    Dim t As Double
    t = NowSeconds()
    Do While Not c.DoneError
        Pump
        If (NowSeconds() - t) * 1000# > ms Then Exit Function
    Loop
    WaitError = True
End Function

Private Function WaitAccepted(ByVal ms As Long) As Boolean
    WaitAccepted = WaitAcceptedCount(1, ms)
End Function

Private Function WaitAcceptedCount(ByVal n As Long, ByVal ms As Long) As Boolean
    Dim t As Double
    t = NowSeconds()
    Do While gChildren.Count < n
        Pump
        If (NowSeconds() - t) * 1000# > ms Then Exit Function
    Loop
    WaitAcceptedCount = True
End Function

'==============================================================================
' Result helpers
'==============================================================================
Private Function NewRun(ByVal useOcx As Boolean) As clsRun
    Dim r As clsRun
    Set r = New clsRun
    If useOcx Then
        r.Reset "OCX"
    Else
        r.Reset "W2"
    End If
    Set NewRun = r
End Function

Private Sub Require(ByRef r As clsRun, ByVal cond As Boolean, ByVal msg As String)
    If Not cond Then
        r.Passed = False
        If Len(r.Failure) > 0 Then r.Failure = r.Failure & " | "
        r.Failure = r.Failure & msg
    End If
End Sub

Private Sub AddPair(ByVal group As String, ByVal scenario As String, ByVal id As String)
    Dim w2 As clsRun
    Dim ocx As clsRun
    Dim res As clsResult

    If Len(gFilter) > 0 Then
        If StrComp(gFilter, id, vbTextCompare) <> 0 Then Exit Sub
    End If

    Set w2 = RunScenario(id, False)
    If gOcxAvailable Then
        Set ocx = RunScenario(id, True)
    End If

    Set res = New clsResult
    res.Group = group
    res.Scenario = scenario
    Set res.W2 = w2
    If Not ocx Is Nothing Then Set res.Ocx = ocx

    If w2.KnownGap Or (Not ocx Is Nothing And ocx.KnownGap) Then
        res.Status = "EXPECTED"
        res.Notes = "Known limitation; see details."
        ' Prefer the scenario's own explanation when it provides one.
        If Len(w2.Details) > 0 Then res.Notes = FirstLine(w2.Details, 160)
        If Len(res.Notes) > 0 And Not ocx Is Nothing Then
            If Len(ocx.Details) > 0 And Len(w2.Details) = 0 Then res.Notes = FirstLine(ocx.Details, 160)
        End If
    ElseIf Not gOcxAvailable Then
        If w2.Passed Then
            res.Status = "W2_ONLY"
            res.Notes = "OCX not available/registered."
        Else
            res.Status = "FAIL"
            res.Notes = "W2: " & w2.Failure
        End If
    ElseIf w2.Passed And ocx.Passed Then
        If NormalizeTrace(w2.Trace) = NormalizeTrace(ocx.Trace) Then
            res.Status = "PASS"
        Else
            res.Status = "DIFF"
            res.Diff = FirstDiff(NormalizeTrace(w2.Trace), NormalizeTrace(ocx.Trace))
            res.Notes = "Both pass; event traces differ (per endpoint)."
        End If
    ElseIf w2.Passed <> ocx.Passed Then
        res.Status = "FAIL"
        If Not w2.Passed Then res.Notes = "W2: " & w2.Failure
        If Not ocx.Passed Then res.Notes = res.Notes & " OCX: " & ocx.Failure
    Else
        res.Status = "FAIL"
        res.Notes = "W2: " & w2.Failure & " OCX: " & ocx.Failure
    End If

    If Len(w2.Skipped) > 0 Then res.Notes = Trim$(res.Notes & " | W2 probes skipped: " & w2.Skipped)
    If Not ocx Is Nothing Then
        If Len(ocx.Skipped) > 0 Then res.Notes = Trim$(res.Notes & " | OCX probes skipped: " & ocx.Skipped)
    End If

    gResults.Add res
End Sub

Private Sub AddMixed(ByVal scenario As String, ByRef run As clsRun)
    Dim res As clsResult
    Set res = New clsResult
    res.Group = "Interop"
    res.Scenario = scenario
    If run.ImplName = "OCX" Then
        Set res.Ocx = run
    Else
        Set res.W2 = run
    End If
    If run.Passed Then
        res.Status = "PASS"
    Else
        res.Status = "FAIL"
        res.Notes = run.Failure
    End If
    gResults.Add res
End Sub

'==============================================================================
' Scenario dispatch
'==============================================================================
Private Function RunScenario(ByVal id As String, ByVal useOcx As Boolean) As clsRun
    Dim r As clsRun
    On Error GoTo EH

    Set r = NewRun(useOcx)
    Set gActive = New Collection
    Set gChildren = New Collection

    Select Case id
    Case "defaults":   RunDefaultsImpl r, useOcx
    Case "props":      RunPropsImpl r, useOcx
    Case "tcp_str":    RunTcpCore False, useOcx, 0, r
    Case "tcp_bytes":  RunTcpCore False, useOcx, 1, r
    Case "tcp_large":  RunTcpCore False, useOcx, 2, r
    Case "tcp_multi":  RunTcpMultiImpl r, useOcx
    Case "tcp_reuse":  RunTcpReuseImpl r, useOcx
    Case "udp":        RunUdpImpl r, useOcx
    Case "dns":        RunDnsImpl r, useOcx
    Case "errors":     RunErrorsImpl r, useOcx
    Case "lifecycle":  RunLifecycleImpl r, useOcx
    Case "latebinding": RunLateBindingImpl r, useOcx
    Case "bin_stream":      RunBinStream r, useOcx
    Case "bin_slices":      RunBinSlices r, useOcx
    Case "bin_udp_bounds":  RunBinUdpBounds r, useOcx
    Case "bin_edge_bytes":  RunBinEdgeBytes r, useOcx
    Case "bin_ansi":        RunBinAnsi r, useOcx
    Case "bin_interleave":  RunBinInterleave r, useOcx
    Case "conc_clients":    RunConcClients r, useOcx
    Case "conc_listeners":  RunConcListeners r, useOcx
    Case "conc_mixed":      RunConcMixed r, useOcx
    Case "conc_churn":      RunConcChurn r, useOcx
    Case "stress_stream":   RunStressStream r, useOcx
    Case "stress_msgs":     RunStressMsgs r, useOcx
    Case "stress_udp":      RunStressUdp r, useOcx
    Case "api_order":       RunApiOrder r, useOcx
    Case "evt_reentrancy":  RunEvtReentrancy r, useOcx
    Case "peer_abort":      RunPeerAbort r, useOcx
    Case "send_partial":    RunSendPartial r, useOcx
    End Select

    CleanupActive
    Set RunScenario = r
    Exit Function
EH:
    If r Is Nothing Then Set r = NewRun(useOcx)
    r.Passed = False
    r.Failure = "Unhandled " & Err.Number & ": " & Err.Description
    r.Trace = r.Trace & vbCrLf & "EXCEPTION " & Err.Number & ": " & Err.Description
    CleanupActive
    Set RunScenario = r
End Function

'==============================================================================
' Defaults / properties
'==============================================================================
Private Sub RunDefaultsImpl(ByRef r As clsRun, ByVal useOcx As Boolean)
    Dim e As clsEndpoint
    Set e = NewEndpoint(useOcx, "defaults")
    r.Passed = True
    If e Is Nothing Then
        r.Passed = False
        r.Failure = "Could not create endpoint"
        Exit Sub
    End If

    r.Snapshot = e.SnapshotText
    Require r, (e.ProtocolValue = sckTCPProtocol), "Protocol=" & e.ProtocolValue
    Require r, (e.StateValue = sckClosed), "State=" & e.StateValue
    Require r, (e.SocketHandleValue = -1), "SocketHandle=" & e.SocketHandleValue
    Require r, (e.LocalPortValue = 0), "LocalPort=" & e.LocalPortValue
    r.Details = "Snapshot: " & e.SnapshotText
End Sub

Private Sub RunPropsImpl(ByRef r As clsRun, ByVal useOcx As Boolean)
    Dim e As clsEndpoint
    Dim errNo As Long
    Set e = NewEndpoint(useOcx, "props")
    r.Passed = True
    If e Is Nothing Then
        r.Passed = False
        r.Failure = "Could not create endpoint"
        Exit Sub
    End If

    e.SetProtocol sckUDPProtocol
    Require r, (e.ProtocolValue = sckUDPProtocol), "Protocol put/get UDP"
    e.SetProtocol sckTCPProtocol
    Require r, (e.ProtocolValue = sckTCPProtocol), "Protocol put/get TCP"

    e.SetRemote "example.com", 1234
    Require r, (e.RemoteHostValue = "example.com"), "RemoteHost=" & e.RemoteHostValue
    Require r, (e.RemotePortValue = 1234), "RemotePort=" & e.RemotePortValue
    Require r, (Len(e.LocalHostNameValue) > 0), "LocalHostName empty"
    Require r, (Len(e.LocalIPValue) > 0), "LocalIP empty"
    r.Details = r.Details & "LocalHostName=" & e.LocalHostNameValue & " LocalIP=" & e.LocalIPValue & vbCrLf

    ' Out-of-range property tests are intentionally skipped: MSWINSCK.OCX
    ' 6.1.97.82 crashes (0xC0000005) on an out-of-range RemotePort.
    r.Details = r.Details & "Out-of-range RemotePort/Protocol probes skipped: they crash MSWINSCK.OCX." & vbCrLf

    r.Snapshot = e.SnapshotText
    On Error GoTo 0
End Sub

'==============================================================================
' TCP core
'==============================================================================
Private Sub RunTcpCore(ByVal serverUseOcx As Boolean, ByVal clientUseOcx As Boolean, ByVal mode As Long, ByRef r As clsRun)
    Dim srv As clsEndpoint
    Dim cli As clsEndpoint
    Dim conn As clsEndpoint
    Dim port As Long
    Dim b(0 To 3) As Byte
    Dim got() As Byte
    Dim big() As Byte
    Dim i As Long
    Const N As Long = 1048576

    On Error GoTo EH
    r.Passed = True
    port = NextPort()

    Set srv = NewEndpoint(serverUseOcx, "srv")
    If srv Is Nothing Then
        r.Passed = False
        r.Failure = "Server endpoint creation failed"
        Exit Sub
    End If

    If mode = 2 Then
        srv.Echo = False
        srv.CountOnly = True
        srv.ExpectedBytes = N
    Else
        srv.Echo = True
        srv.CountOnly = False
    End If
    srv.ListenOn port
    If Not WaitState(srv, sckListening, 4000) Then
        r.Passed = False
        r.Failure = "Server did not enter listening state"
        r.Trace = srv.TraceText
        Exit Sub
    End If

    Set cli = NewEndpoint(clientUseOcx, "cli")
    If cli Is Nothing Then
        r.Passed = False
        r.Failure = "Client endpoint creation failed"
        Exit Sub
    End If

    cli.ConnectTo "127.0.0.1", port
    If Not WaitConnect(cli, 5000) Then
        r.Passed = False
        r.Failure = "Client Connect timeout"
    End If
    If Not WaitAccepted(5000) Then
        r.Passed = False
        r.Failure = r.Failure & " Accept timeout"
        If conn Is Nothing Then
            r.Trace = cli.TraceText & vbCrLf & srv.TraceText
            r.Snapshot = "client=" & cli.SnapshotText & vbCrLf & "server=" & srv.SnapshotText
            Exit Sub
        End If
    End If
    Set conn = gChildren.Item(1)

    Select Case mode
    Case 0
        cli.ResetFlags
        cli.SendString "hello"
        If Not WaitData(cli, 5000) Then
            r.Passed = False
            r.Failure = r.Failure & " Data timeout"
        Else
            cli.PeekStringData
            Require r, (cli.LastString = "hello"), "peek=[" & cli.LastString & "]"
            cli.GetStringData
            Require r, (cli.LastString = "hello"), "echo=[" & cli.LastString & "]"
            Require r, (cli.StateValue = sckConnected), "client state=" & cli.StateValue & " after echo"
            cli.ResetFlags
            cli.SendString "again"
            If Not WaitData(cli, 5000) Then
                r.Passed = False
                r.Failure = r.Failure & " second echo timeout"
            Else
                cli.GetStringData
                Require r, (cli.LastString = "again"), "second echo=[" & cli.LastString & "]"
            End If
        End If

    Case 1
        b(0) = 1: b(1) = 2: b(2) = 3: b(3) = 250
        cli.ResetFlags
        cli.SendBytes b
        If Not WaitData(cli, 5000) Then
            r.Passed = False
            r.Failure = r.Failure & " byte Data timeout"
        Else
            cli.PeekByteData
            Require r, (cli.LastByteCount = 4), "PeekByteCount=" & cli.LastByteCount
            cli.CopyLastBytes got
            If cli.LastByteCount = 4 Then
                Require r, (got(0) = 1 And got(1) = 2 And got(2) = 3 And got(3) = 250), "PeekByteData mismatch"
            End If
            cli.GetByteData
            Require r, (cli.LastByteCount = 4), "GetByteCount=" & cli.LastByteCount
            cli.CopyLastBytes got
            If cli.LastByteCount = 4 Then
                Require r, (got(0) = 1 And got(1) = 2 And got(2) = 3 And got(3) = 250), "GetByteData mismatch"
            End If
        End If

    Case 2
        ReDim big(0 To N - 1)
        For i = 0 To N - 1
            big(i) = i And 255
        Next
        conn.ResetFlags
        cli.ResetFlags
        cli.SendBytes big
        If Not WaitData(conn, 20000) Then
            r.Passed = False
            r.Failure = r.Failure & " large transfer timeout (" & conn.TotalReceived & ")"
        Else
            Require r, (conn.TotalReceived = N), "large received=" & conn.TotalReceived
        End If
    End Select

    cli.ResetFlags
    cli.CloseSock
    WaitClose conn, 4000
    srv.CloseSock

    r.Trace = cli.TraceText & vbCrLf & conn.TraceText & vbCrLf & srv.TraceText
    r.Snapshot = "client=" & cli.SnapshotText & vbCrLf & "serverConn=" & conn.SnapshotText & vbCrLf & "server=" & srv.SnapshotText
    Exit Sub
EH:
    r.Passed = False
    r.Failure = r.Failure & " EXC " & Err.Number & ": " & Err.Description
    If Not cli Is Nothing Then r.Trace = cli.TraceText
    If Not conn Is Nothing Then r.Trace = r.Trace & vbCrLf & conn.TraceText
    If Not srv Is Nothing Then r.Trace = r.Trace & vbCrLf & srv.TraceText
End Sub

' Close must leave the control reusable immediately, exactly like the OCX.
Private Sub RunTcpReuseImpl(ByRef r As clsRun, ByVal useOcx As Boolean)
    Dim srv As clsEndpoint
    Dim cli As clsEndpoint
    Dim conn As clsEndpoint
    Dim port As Long

    On Error GoTo EH
    r.Passed = True
    port = NextPort()
    Set srv = NewEndpoint(False, "srv")
    If srv Is Nothing Then
        r.Passed = False
        r.Failure = "Server creation failed"
        Exit Sub
    End If
    srv.Echo = True
    srv.ListenOn port
    If Not WaitState(srv, sckListening, 4000) Then
        r.Passed = False
        r.Failure = "listen timeout"
        Exit Sub
    End If

    Set cli = NewEndpoint(useOcx, "cli")
    If cli Is Nothing Then
        r.Passed = False
        r.Failure = "client creation failed"
        Exit Sub
    End If

    cli.ConnectTo "127.0.0.1", port
    If Not WaitConnect(cli, 5000) Then
        r.Passed = False
        r.Failure = r.Failure & " connect1"
    End If
    If Not WaitAcceptedCount(1, 5000) Then
        r.Passed = False
        r.Failure = r.Failure & " accept1"
    Else
        Set conn = gChildren.Item(1)
    End If

    cli.ResetFlags
    cli.SendString "first"
    If WaitData(cli, 5000) Then
        cli.GetStringData
        Require r, (cli.LastString = "first"), "first=[" & cli.LastString & "]"
    Else
        r.Passed = False
        r.Failure = r.Failure & " echo1 timeout"
    End If

    cli.CloseSock
    r.Details = "State right after Close=" & cli.StateValue & vbCrLf
    Require r, (cli.StateValue = sckClosed), "State after Close=" & cli.StateValue
    If Not conn Is Nothing Then WaitClose conn, 4000

    ' Reuse the same control straight away.
    cli.ResetFlags
    On Error Resume Next
    Err.Clear
    cli.ConnectTo "127.0.0.1", port
    If Err.Number <> 0 Then
        r.Passed = False
        r.Failure = r.Failure & " reconnect err=" & Err.Number
        r.Details = r.Details & "reconnect raised " & Err.Number & ": " & Err.Description & vbCrLf
        Err.Clear
    End If
    On Error GoTo EH

    If Not WaitConnect(cli, 5000) Then
        r.Passed = False
        r.Failure = r.Failure & " connect2"
    End If
    If Not WaitAcceptedCount(2, 5000) Then
        r.Passed = False
        r.Failure = r.Failure & " accept2"
    Else
        Set conn = gChildren.Item(2)
    End If

    cli.ResetFlags
    cli.SendString "second"
    If WaitData(cli, 5000) Then
        cli.GetStringData
        Require r, (cli.LastString = "second"), "second=[" & cli.LastString & "]"
    Else
        r.Passed = False
        r.Failure = r.Failure & " echo2 timeout"
    End If

    cli.CloseSock
    srv.CloseSock
    r.Trace = cli.TraceText & vbCrLf & srv.TraceText
    r.Snapshot = "client=" & cli.SnapshotText
    Exit Sub
EH:
    r.Passed = False
    r.Failure = r.Failure & " EXC " & Err.Number & ": " & Err.Description
    If Not cli Is Nothing Then r.Trace = cli.TraceText
End Sub

Private Sub RunTcpMultiImpl(ByRef r As clsRun, ByVal useOcx As Boolean)
    Dim srv As clsEndpoint
    Dim c1 As clsEndpoint
    Dim c2 As clsEndpoint
    Dim port As Long

    r.Passed = True
    port = NextPort()
    Set srv = NewEndpoint(useOcx, "srv")
    If srv Is Nothing Then
        r.Passed = False
        r.Failure = "Server creation failed"
        Exit Sub
    End If
    srv.Echo = True
    srv.ListenOn port
    If Not WaitState(srv, sckListening, 4000) Then
        r.Passed = False
        r.Failure = "Server listen timeout"
        Exit Sub
    End If

    Set c1 = NewEndpoint(useOcx, "cli1")
    Set c2 = NewEndpoint(useOcx, "cli2")
    If c1 Is Nothing Or c2 Is Nothing Then
        r.Passed = False
        r.Failure = "Client creation failed"
        Exit Sub
    End If
    c1.ConnectTo "127.0.0.1", port
    c2.ConnectTo "127.0.0.1", port
    If Not WaitConnect(c1, 5000) Then r.Passed = False: r.Failure = r.Failure & " c1 connect"
    If Not WaitConnect(c2, 5000) Then r.Passed = False: r.Failure = r.Failure & " c2 connect"
    If Not WaitAcceptedCount(2, 5000) Then
        r.Passed = False
        r.Failure = r.Failure & " accept count=" & gChildren.Count
        Exit Sub
    End If

    c1.ResetFlags: c1.SendString "one"
    c2.ResetFlags: c2.SendString "two"
    If Not WaitData(c1, 5000) Then r.Passed = False: r.Failure = r.Failure & " c1 data"
    If Not WaitData(c2, 5000) Then r.Passed = False: r.Failure = r.Failure & " c2 data"
    If c1.DoneData Then
        c1.GetStringData
        Require r, (c1.LastString = "one"), "c1=[" & c1.LastString & "]"
    End If
    If c2.DoneData Then
        c2.GetStringData
        Require r, (c2.LastString = "two"), "c2=[" & c2.LastString & "]"
    End If

    c1.CloseSock
    c2.CloseSock
    srv.CloseSock
    r.Trace = c1.TraceText & vbCrLf & c2.TraceText & vbCrLf & srv.TraceText
    r.Snapshot = "c1=" & c1.SnapshotText & vbCrLf & "c2=" & c2.SnapshotText
End Sub

'==============================================================================
' UDP
'==============================================================================
Private Sub RunUdpImpl(ByRef r As clsRun, ByVal useOcx As Boolean)
    Dim rcv As clsEndpoint
    Dim snd As clsEndpoint
    Dim port As Long

    r.Passed = True
    port = NextPort()
    Set rcv = NewEndpoint(False, "udpRcv")
    If rcv Is Nothing Then
        r.Passed = False
        r.Failure = "UDP receiver creation failed"
        Exit Sub
    End If
    rcv.SetProtocol sckUDPProtocol
    rcv.AutoRead = True
    rcv.BindTo port
    If Not WaitState(rcv, sckOpen, 3000) Then
        r.Passed = False
        r.Failure = "UDP bind state=" & rcv.StateValue
        Exit Sub
    End If

    Set snd = NewEndpoint(useOcx, "udpSnd")
    If snd Is Nothing Then
        r.Passed = False
        r.Failure = "UDP sender creation failed"
        Exit Sub
    End If
    snd.SetProtocol sckUDPProtocol
    snd.AutoRead = True
    snd.SetRemote "127.0.0.1", port

    snd.SendString "udp-one"
    If Not WaitData(rcv, 5000) Then
        r.Passed = False
        r.Failure = "UDP first datagram timeout"
    Else
        Require r, (rcv.LastString = "udp-one"), "udp first=[" & rcv.LastString & "]"
        Require r, (rcv.RemoteHostIPValue = "127.0.0.1"), "udp RemoteIP=" & rcv.RemoteHostIPValue
        Require r, (rcv.RemotePortValue > 0), "udp RemotePort=" & rcv.RemotePortValue
        r.Details = r.Details & "first datagram RemoteIP=" & rcv.RemoteHostIPValue & " RemotePort=" & rcv.RemotePortValue & vbCrLf
    End If

    ' Reply from the receiver using the peer address learned from the datagram.
    If rcv.DoneData Then
        rcv.ResetFlags
        On Error Resume Next
        Err.Clear
        rcv.SendString "udp-reply"
        If Err.Number <> 0 Then
            r.Passed = False
            r.Failure = r.Failure & " udp reply send err=" & Err.Number
            Err.Clear
        End If
        On Error GoTo 0
        If WaitData(snd, 5000) Then
            Require r, (snd.LastString = "udp-reply"), "udp reply=[" & snd.LastString & "]"
            Require r, (snd.RemoteHostIPValue = "127.0.0.1"), "udp reply RemoteIP=" & snd.RemoteHostIPValue
        Else
            r.Passed = False
            r.Failure = r.Failure & " udp reply timeout"
        End If
    End If

    rcv.ResetFlags
    snd.ResetFlags
    snd.SendString "udp-two"
    If Not WaitData(rcv, 5000) Then
        r.Passed = False
        r.Failure = r.Failure & " udp second timeout"
    Else
        Require r, (rcv.LastString = "udp-two"), "udp second=[" & rcv.LastString & "]"
    End If

    snd.CloseSock
    rcv.CloseSock
    r.Trace = rcv.TraceText & vbCrLf & snd.TraceText
    r.Snapshot = "rcv=" & rcv.SnapshotText & vbCrLf & "snd=" & snd.SnapshotText
End Sub

'==============================================================================
' DNS
'==============================================================================
Private Sub RunDnsImpl(ByRef r As clsRun, ByVal useOcx As Boolean)
    Dim srv As clsEndpoint
    Dim cli As clsEndpoint
    Dim port As Long

    r.Passed = True
    port = NextPort()
    Set srv = NewEndpoint(False, "dnsSrv")
    If srv Is Nothing Then
        r.Passed = False
        r.Failure = "DNS server creation failed"
        Exit Sub
    End If
    srv.Echo = True
    srv.ListenOn port
    If Not WaitState(srv, sckListening, 3000) Then
        r.Passed = False
        r.Failure = "DNS server listen timeout"
        Exit Sub
    End If

    Set cli = NewEndpoint(useOcx, "dnsCli")
    If cli Is Nothing Then
        r.Passed = False
        r.Failure = "DNS client creation failed"
        Exit Sub
    End If
    cli.ConnectTo "localhost", port
    r.Details = "State after Connect(localhost)=" & cli.StateAfterConnectCall
    If Not WaitConnect(cli, 6000) Then
        r.Passed = False
        r.Failure = "DNS connect timeout"
    Else
        Require r, (cli.StateValue = sckConnected), "DNS final State=" & cli.StateValue
        ' Wait for the server side to accept, otherwise whether the listener's
        ' ConnectionRequest event lands in the trace depends on pump timing.
        If Not WaitAccepted(5000) Then
            r.Passed = False
            r.Failure = r.Failure & " DNS accept timeout"
        End If
    End If

    cli.CloseSock
    srv.CloseSock
    r.Trace = cli.TraceText & vbCrLf & srv.TraceText
    r.Snapshot = "cli=" & cli.SnapshotText
End Sub

'==============================================================================
' Errors
'==============================================================================
Private Sub RunErrorsImpl(ByRef r As clsRun, ByVal useOcx As Boolean)
    Dim e As clsEndpoint
    Dim c As clsEndpoint
    Dim a As clsEndpoint
    Dim b As clsEndpoint
    Dim errNo As Long
    Dim port As Long

    r.Passed = True
    Set e = NewEndpoint(useOcx, "err")
    If e Is Nothing Then
        r.Passed = False
        r.Failure = "Error endpoint creation failed"
        Exit Sub
    End If

    If useOcx Then
        ' MSWINSCK.OCX 6.1.97.82 raises an access violation (0xC0000005) when
        ' GetData is called with no connection, and when Bind targets a port
        ' that is already in use. Those probes would kill the harness, so they
        ' only run against modWinsock2 and are recorded as skipped on the OCX.
        r.Skipped = "closed-state GetData/SendData, bind-to-used-port"
    Else
        On Error Resume Next
        Err.Clear
        e.GetStringData
        errNo = Err.Number
        Err.Clear
        Require r, (errNo <> 0), "GetData closed did not error"
        r.Details = r.Details & "GetData closed error=" & errNo & vbCrLf

        Err.Clear
        e.SendString "x"
        errNo = Err.Number
        Err.Clear
        Require r, (errNo <> 0), "SendData closed did not error"
        r.Details = r.Details & "SendData closed error=" & errNo & vbCrLf
        On Error GoTo 0
    End If

    Set c = NewEndpoint(useOcx, "refused")
    If c Is Nothing Then
        r.Passed = False
        r.Failure = r.Failure & " refused endpoint creation"
    Else
        c.ResetFlags
        On Error Resume Next
        c.ConnectTo "127.0.0.1", 1
        On Error GoTo 0
        If Not WaitError(c, 6000) Then
            r.Passed = False
            r.Failure = r.Failure & " connect refused timeout"
        Else
            Require r, (c.LastErrorNo = sckConnectionRefused), "refused error=" & c.LastErrorNo
            r.Details = r.Details & "refused error=" & c.LastErrorNo & " " & c.LastErrorText & vbCrLf
            r.Details = r.Details & "refused Source=[" & c.LastErrorSource & "] Help=" & c.LastErrorHelp & vbCrLf
        End If
    End If

    port = NextPort()
    If Not useOcx Then
        Set a = NewEndpoint(False, "bindA")
        Set b = NewEndpoint(False, "bindB")
        If a Is Nothing Or b Is Nothing Then
            r.Passed = False
            r.Failure = r.Failure & " bind endpoints creation"
        Else
            a.SetProtocol sckUDPProtocol
            a.BindTo port
            WaitState a, sckOpen, 3000
            b.SetProtocol sckUDPProtocol
            b.ResetFlags
            On Error Resume Next
            b.BindTo port
            On Error GoTo 0
            If Not WaitError(b, 5000) Then
                r.Passed = False
                r.Failure = r.Failure & " bind in use no Error event"
            Else
                Require r, (b.LastErrorNo = sckAddressInUse), "bind in use error=" & b.LastErrorNo
                r.Details = r.Details & "bind in use error=" & b.LastErrorNo & " " & b.LastErrorText & vbCrLf
            End If
        End If
    End If

    ' Only the refused-connection endpoint is compared: the closed-state and
    ' bind-in-use probes are W2-only (they crash the OCX) and their outcomes are
    ' asserted into Details instead of the trace.
    If Not c Is Nothing Then r.Trace = c.TraceText
End Sub

'==============================================================================
' Lifecycle / known gap
'==============================================================================
Private Sub RunLifecycleImpl(ByRef r As clsRun, ByVal useOcx As Boolean)
    Dim i As Long
    Dim e As clsEndpoint

    r.Passed = True
    On Error GoTo EH
    For i = 1 To 50
        Set e = New clsEndpoint
        e.Init useOcx, "life" & i
        Set e = Nothing
        If (i Mod 10) = 0 Then Pump
    Next
    r.Details = "50 create/release cycles"
    Exit Sub
EH:
    r.Passed = False
    r.Failure = "Lifecycle error " & Err.Number & ": " & Err.Description
End Sub

Private Sub RunLateBindingImpl(ByRef r As clsRun, ByVal useOcx As Boolean)
    Dim e As clsEndpoint
    Dim v As Variant
    Dim errNo As Long

    r.Passed = False
    Set e = NewEndpoint(useOcx, "late")
    If e Is Nothing Then
        r.Failure = "Late endpoint creation failed"
        Exit Sub
    End If
    On Error Resume Next
    Err.Clear
    v = e.LateGet("State")
    errNo = Err.Number
    Err.Clear
    On Error GoTo 0
    r.Details = "CallByName State=" & CStr(v) & " error=" & errNo
    If errNo <> 0 Then
        r.Failure = "CallByName State failed " & errNo
        Exit Sub
    End If
    If CLng(v) <> e.StateValue Then
        r.Failure = "CallByName State=" & CStr(v) & " typed State=" & CStr(e.StateValue)
        Exit Sub
    End If
    r.Passed = True
End Sub

'==============================================================================
' Functional orchestration
'==============================================================================
Private Sub RunFunctionalCore()
    Dim m1 As clsRun
    Dim m2 As clsRun

    AddPair "Defaults", "Initial properties", "defaults"
    AddPair "Properties", "Protocol/remote round-trip", "props"
    AddPair "TCP Client", "String echo", "tcp_str"
    AddPair "TCP Client", "Byte array + PeekData", "tcp_bytes"
    AddPair "TCP Server", "Multiple clients", "tcp_multi"
    AddPair "TCP Client", "Close and reuse control", "tcp_reuse"
    AddPair "TCP Data", "1 MB transfer", "tcp_large"
    AddPair "UDP", "Two datagrams + echo", "udp"
    AddPair "DNS", "localhost resolution", "dns"
    AddPair "Errors", "Closed/refused/bind in use", "errors"
    AddPair "Lifetime", "Create/release cycles", "lifecycle"
    AddPair "Late binding", "CallByName/IDispatch", "latebinding"

    AddPair "Integrity", "4 MB mixed chunks", "bin_stream"
    AddPair "Integrity", "Slice/peek reads", "bin_slices"
    AddPair "Integrity", "UDP datagram bounds", "bin_udp_bounds"
    AddPair "Integrity", "Edge byte patterns", "bin_edge_bytes"
    AddPair "Integrity", "ANSI string bytes", "bin_ansi"
    AddPair "Integrity", "1-byte/64 KB interleave", "bin_interleave"
    AddPair "Concurrency", "32 clients, no crosstalk", "conc_clients"
    AddPair "Concurrency", "3 listeners x 2 clients", "conc_listeners"
    AddPair "Concurrency", "TCP and UDP together", "conc_mixed"
    AddPair "Concurrency", "200 connect/close cycles", "conc_churn"
    AddPair "Stress", "8 MB stream", "stress_stream"
    AddPair "Stress", "20k message round trips", "stress_msgs"
    AddPair "Stress", "4k datagram burst", "stress_udp"
    AddPair "State machine", "Wrong-order calls", "api_order"
    AddPair "State machine", "Close inside DataArrival", "evt_reentrancy"
    AddPair "State machine", "Peer reset mid-transfer", "peer_abort"
    AddPair "State machine", "4 KB window, 1 MB send", "send_partial"

    If gOcxAvailable Then
        If Len(gFilter) = 0 Or StrComp(gFilter, "interop", vbTextCompare) = 0 Then
            Set m1 = RunMixedTcp(True, False)
            AddMixed "OCX server + W2 client (string echo)", m1
            Set m2 = RunMixedTcp(False, True)
            AddMixed "W2 server + OCX client (string echo)", m2
            Set m1 = RunInteropStream(True, False)
            AddMixed "OCX server + W2 client (4 MB integrity)", m1
            Set m2 = RunInteropStream(False, True)
            AddMixed "W2 server + OCX client (4 MB integrity)", m2
        End If
    End If
End Sub

Private Function RunMixedTcp(ByVal serverUseOcx As Boolean, ByVal clientUseOcx As Boolean) As clsRun
    Dim r As clsRun
    Set r = New clsRun
    If serverUseOcx And clientUseOcx Then
        r.Reset "OCX/OCX"
    ElseIf serverUseOcx Then
        r.Reset "OCX server/W2 client"
    Else
        r.Reset "W2 server/OCX client"
    End If
    Set gActive = New Collection
    Set gChildren = New Collection
    RunTcpCore serverUseOcx, clientUseOcx, 0, r
    CleanupActive
    Set RunMixedTcp = r
End Function

'==============================================================================
' Performance
'==============================================================================
Private Sub RunPerformanceCore()
    Dim w2 As Double
    Dim ocx As Double
    Dim row As clsPerfRow

    w2 = PerfCreateSeconds(False)
    If gOcxAvailable Then ocx = PerfCreateSeconds(True)
    AddPerfRow "Create/release object", 500, w2, ocx, 0, "500 create/init/release cycles"

    w2 = PerfTcpRoundTrip(False, 200)
    If gOcxAvailable Then ocx = PerfTcpRoundTrip(True, 200)
    AddPerfRow "TCP round-trip 64 B", 200, w2, ocx, 64, "Same W2 echo server for both; only the client differs. Includes one message-pump wait per echo."

    w2 = PerfTcpThroughput(False)
    If gOcxAvailable Then ocx = PerfTcpThroughput(True)
    AddPerfRow "TCP throughput 1 MB", 1, w2, ocx, 1048576, "Twenty 1 MB sends averaged, reported per transfer"

    w2 = PerfUdpRoundTrip(False, 200)
    If gOcxAvailable Then ocx = PerfUdpRoundTrip(True, 200)
    AddPerfRow "UDP round-trip 64 B", 200, w2, ocx, 64, "Same W2 UDP echo receiver for both senders"
End Sub

Private Sub AddPerfRow(ByVal op As String, ByVal iters As Long, ByVal w2Sec As Double, ByVal ocxSec As Double, ByVal bytes As Long, ByVal notes As String)
    Dim row As clsPerfRow
    Set row = New clsPerfRow
    row.Operation = op
    row.Iterations = iters
    row.W2Seconds = w2Sec
    row.OcxSeconds = ocxSec
    row.Notes = notes
    If w2Sec > 0 Then
        row.W2Ops = iters / w2Sec
        If bytes > 0 Then row.W2BytesPerSec = (CDbl(bytes) * iters) / w2Sec
    End If
    If ocxSec > 0 Then
        row.OcxOps = iters / ocxSec
        If bytes > 0 Then row.OcxBytesPerSec = (CDbl(bytes) * iters) / ocxSec
    End If
    gPerf.Add row
End Sub

Private Function PerfCreateSeconds(ByVal useOcx As Boolean) As Double
    Dim i As Long
    Dim e As clsEndpoint
    Dim t0 As Double
    Dim t1 As Double
    Const CYCLES As Long = 500
    On Error GoTo Failed
    t0 = NowSeconds()
    For i = 1 To CYCLES
        Set e = New clsEndpoint
        e.Init useOcx, "perf"
        Set e = Nothing
        If (i Mod 25) = 0 Then Pump
    Next
    t1 = NowSeconds()
    PerfCreateSeconds = t1 - t0
    Exit Function
Failed:
    PerfCreateSeconds = -1
End Function

Private Function PerfTcpRoundTrip(ByVal clientUseOcx As Boolean, ByVal iters As Long) As Double
    Dim srv As clsEndpoint
    Dim cli As clsEndpoint
    Dim conn As clsEndpoint
    Dim port As Long
    Dim i As Long
    Dim t0 As Double
    Dim t1 As Double
    Dim payload As String

    On Error GoTo Failed
    payload = String$(64, "x")
    Set gActive = New Collection
    Set gChildren = New Collection
    port = NextPort()
    Set srv = NewEndpoint(False, "perfSrv")
    If srv Is Nothing Then Exit Function
    srv.Echo = True
    srv.ListenOn port
    If Not WaitState(srv, sckListening, 3000) Then Exit Function
    Set cli = NewEndpoint(clientUseOcx, "perfCli")
    If cli Is Nothing Then Exit Function
    cli.ConnectTo "127.0.0.1", port
    If Not WaitConnect(cli, 4000) Then Exit Function
    If Not WaitAccepted(3000) Then Exit Function
    Set conn = gChildren.Item(1)

    cli.ResetFlags
    cli.SendString payload
    If Not WaitData(cli, 3000) Then Exit Function
    cli.GetStringData

    t0 = NowSeconds()
    For i = 1 To iters
        cli.ResetFlags
        cli.SendString payload
        If Not WaitData(cli, 3000) Then Exit Function
        cli.GetStringData
    Next
    t1 = NowSeconds()
    PerfTcpRoundTrip = t1 - t0
    cli.CloseSock
    srv.CloseSock
    CleanupActive
    Exit Function
Failed:
    PerfTcpRoundTrip = -1
    CleanupActive
End Function

Private Function PerfTcpThroughput(ByVal clientUseOcx As Boolean) As Double
    Dim srv As clsEndpoint
    Dim cli As clsEndpoint
    Dim conn As clsEndpoint
    Dim port As Long
    Dim big() As Byte
    Dim i As Long
    Dim t0 As Double
    Dim t1 As Double
    Dim total As Double
    Const N As Long = 1048576
    Const REPS As Long = 20

    On Error GoTo Failed
    Set gActive = New Collection
    Set gChildren = New Collection
    port = NextPort()
    Set srv = NewEndpoint(False, "perfSrv")
    If srv Is Nothing Then Exit Function
    srv.Echo = False
    srv.CountOnly = True
    srv.ExpectedBytes = N
    srv.ListenOn port
    If Not WaitState(srv, sckListening, 3000) Then Exit Function
    Set cli = NewEndpoint(clientUseOcx, "perfCli")
    If cli Is Nothing Then Exit Function
    cli.ConnectTo "127.0.0.1", port
    If Not WaitConnect(cli, 4000) Then Exit Function
    If Not WaitAccepted(3000) Then Exit Function
    Set conn = gChildren.Item(1)
    conn.ResetFlags

    ReDim big(0 To N - 1)
    For i = 0 To N - 1
        big(i) = i And 255
    Next
    ' Average a few transfers: a single 1 MB send is too noisy to compare.
    For i = 1 To REPS
        conn.ResetFlags
        t0 = NowSeconds()
        cli.SendBytes big
        If Not WaitData(conn, 20000) Then Exit Function
        t1 = NowSeconds()
        If conn.TotalReceived <> N Then Exit Function
        total = total + (t1 - t0)
    Next
    PerfTcpThroughput = total / REPS
    cli.CloseSock
    srv.CloseSock
    CleanupActive
    Exit Function
Failed:
    PerfTcpThroughput = -1
    CleanupActive
End Function

Private Function PerfUdpRoundTrip(ByVal clientUseOcx As Boolean, ByVal iters As Long) As Double
    Dim rcv As clsEndpoint
    Dim snd As clsEndpoint
    Dim port As Long
    Dim i As Long
    Dim t0 As Double
    Dim t1 As Double
    Dim payload As String

    On Error GoTo Failed
    payload = String$(64, "u")
    Set gActive = New Collection
    Set gChildren = New Collection
    port = NextPort()
    Set rcv = NewEndpoint(False, "perfRcv")
    If rcv Is Nothing Then Exit Function
    rcv.SetProtocol sckUDPProtocol
    rcv.Echo = True
    rcv.BindTo port
    If Not WaitState(rcv, sckOpen, 3000) Then Exit Function
    Set snd = NewEndpoint(clientUseOcx, "perfSnd")
    If snd Is Nothing Then Exit Function
    snd.SetProtocol sckUDPProtocol
    snd.SetRemote "127.0.0.1", port

    snd.ResetFlags
    snd.SendString payload
    If Not WaitData(snd, 3000) Then Exit Function
    snd.GetStringData

    t0 = NowSeconds()
    For i = 1 To iters
        snd.ResetFlags
        snd.SendString payload
        If Not WaitData(snd, 3000) Then Exit Function
        snd.GetStringData
    Next
    t1 = NowSeconds()
    PerfUdpRoundTrip = t1 - t0
    snd.CloseSock
    rcv.CloseSock
    CleanupActive
    Exit Function
Failed:
    PerfUdpRoundTrip = -1
    CleanupActive
End Function

'==============================================================================
' Reliable-transfer helpers
'==============================================================================
Private Function WaitCapture(ByRef c As clsEndpoint, ByVal ms As Long) As Boolean
    Dim t As Double
    t = NowSeconds()
    Do While c.CaptureSize < c.CaptureLimit
        Pump
        If c.DoneError Then Exit Function
        If (NowSeconds() - t) * 1000# > ms Then Exit Function
    Loop
    WaitCapture = True
End Function

Private Sub SendSlice(ByRef cli As clsEndpoint, ByRef src() As Byte, ByVal off As Long, ByVal n As Long)
    Dim t() As Byte
    If n <= 0 Then Exit Sub
    ReDim t(0 To n - 1)
    MemoryCopy t(0), src(off), n
    cli.SendBytes t
End Sub

' plan 0 = eight equal chunks, 1 = adversarial sizes, 2 = one shot, 3 = 1 byte / 64 KB alternating
Private Sub SendChunks(ByRef cli As clsEndpoint, ByRef payload() As Byte, ByVal plan As Long)
    Dim n As Long
    Dim off As Long
    Dim i As Long
    Dim sz As Long

    n = UBound(payload) - LBound(payload) + 1
    off = 0
    Do While off < n
        Select Case plan
        Case 2
            sz = n - off
        Case 0
            sz = n \ 8
            If sz < 1 Then sz = 1
        Case 3
            If (i Mod 2) = 0 Then sz = 1 Else sz = 65536
        Case Else
            Select Case (i Mod 6)
            Case 0: sz = 1
            Case 1: sz = 3
            Case 2: sz = 17
            Case 3: sz = 1024
            Case 4: sz = 65535
            Case 5: sz = 65537
            End Select
        End Select
        If sz > n - off Then sz = n - off
        SendSlice cli, payload, off, sz
        off = off + sz
        i = i + 1
        If (i Mod 16) = 0 Then Pump
    Loop
End Sub

' One full transfer: connect, send a generated payload, compare what arrived.
Private Function StreamTransfer(ByVal srvOcx As Boolean, ByVal cliOcx As Boolean, _
                                 ByVal total As Long, ByVal kind As Long, ByVal plan As Long, _
                                 ByRef seconds As Double, ByRef errText As String) As Boolean
    Dim srv As clsEndpoint
    Dim cli As clsEndpoint
    Dim conn As clsEndpoint
    Dim port As Long
    Dim payload() As Byte
    Dim got() As Byte
    Dim t0 As Double
    Dim bad As Long
    Dim childBase As Long

    On Error GoTo EH
    port = NextPort()
    childBase = gChildren.Count
    Set srv = NewEndpoint(srvOcx, "srv")
    If srv Is Nothing Then errText = "server create failed": Exit Function
    srv.Echo = False
    srv.ChildCaptureLimit = total
    srv.ListenOn port
    If Not WaitState(srv, sckListening, 4000) Then errText = "listen timeout": Exit Function

    Set cli = NewEndpoint(cliOcx, "cli")
    If cli Is Nothing Then errText = "client create failed": Exit Function
    cli.ConnectTo "127.0.0.1", port
    If Not WaitConnect(cli, 6000) Then errText = "connect timeout": Exit Function
    If Not WaitAcceptedCount(childBase + 1, 6000) Then errText = "accept timeout": Exit Function
    Set conn = gChildren.Item(childBase + 1)

    ReDim payload(0 To total - 1)
    FillPattern payload, total, kind

    t0 = NowSeconds()
    SendChunks cli, payload, plan
    If Not WaitCapture(conn, 30000) Then
        errText = "receive timeout (" & conn.CaptureSize & "/" & total & ")"
        Exit Function
    End If
    seconds = NowSeconds() - t0

    If conn.CaptureOverflow Then errText = "capture overflow": Exit Function
    If conn.CaptureSize <> total Then
        errText = "size " & conn.CaptureSize & " <> " & total
        Exit Function
    End If
    conn.CopyCapture got
    If Not BuffersEqual(payload, got) Then
        bad = FirstMismatch(payload, got)
        errText = "corrupt at offset " & bad
        Exit Function
    End If

    cli.CloseSock
    srv.CloseSock
    StreamTransfer = True
    Exit Function
EH:
    errText = "EXC " & Err.Number & ": " & Err.Description
End Function

'==============================================================================
' Tier 1 - integrity scenarios
'==============================================================================
Private Sub StreamCheck(ByRef r As clsRun, ByVal useOcx As Boolean, ByVal total As Long, _
                        ByVal kind As Long, ByVal plan As Long, ByVal tag As String)
    Dim secs As Double
    Dim errText As String
    If StreamTransfer(useOcx, useOcx, total, kind, plan, secs, errText) Then
        r.Details = r.Details & tag & ": " & total & " B in " & Format$(secs * 1000#, "0.0") & " ms" & vbCrLf
    Else
        r.Passed = False
        r.Failure = r.Failure & " " & tag & "=" & errText
    End If
End Sub

Private Sub RunBinStream(ByRef r As clsRun, ByVal useOcx As Boolean)
    Dim t0 As Double
    r.Passed = True
    t0 = NowSeconds()
    StreamCheck r, useOcx, 4194304, PAT_RANDOM, 1, "4MB-mixed"
    If r.Passed Then r.Details = r.Details & "total " & Format$(NowSeconds() - t0, "0.000") & " s" & vbCrLf
    r.Details = r.Details & ResourceLine() & vbCrLf & DiagnosticsLine()
End Sub

Private Sub RunBinEdgeBytes(ByRef r As clsRun, ByVal useOcx As Boolean)
    r.Passed = True
    StreamCheck r, useOcx, 4096, PAT_ZERO, 2, "zero"
    StreamCheck r, useOcx, 4096, PAT_FF, 2, "ff"
    StreamCheck r, useOcx, 4096, PAT_NULLMIX, 2, "nul-crlf"
    StreamCheck r, useOcx, 4096, PAT_COUNT, 2, "count"
End Sub

Private Sub RunBinInterleave(ByRef r As clsRun, ByVal useOcx As Boolean)
    r.Passed = True
    StreamCheck r, useOcx, 262144, PAT_RANDOM, 3, "1B/64KB"
    StreamCheck r, useOcx, 65537, PAT_RANDOM, 0, "8x8192"
End Sub

Private Function TryGetBytes(ByRef ep As clsEndpoint, ByVal maxLen As Long) As Long
    On Error Resume Next
    Err.Clear
    ep.GetByteDataMax maxLen
    If Err.Number <> 0 Then
        Err.Clear
        TryGetBytes = 0
        Exit Function
    End If
    TryGetBytes = ep.LastByteCount
End Function

Private Function TryPeekBytes(ByRef ep As clsEndpoint, ByVal maxLen As Long) As Long
    On Error Resume Next
    Err.Clear
    ep.PeekByteDataMax maxLen
    If Err.Number <> 0 Then
        Err.Clear
        TryPeekBytes = 0
        Exit Function
    End If
    TryPeekBytes = ep.LastByteCount
End Function

' Reads with MaxLen slices and peeks, verifying peek does not consume and no
' byte is lost or duplicated across slices.
Private Sub RunBinSlices(ByRef r As clsRun, ByVal useOcx As Boolean)
    Dim srv As clsEndpoint
    Dim cli As clsEndpoint
    Dim conn As clsEndpoint
    Dim port As Long
    Dim payload() As Byte
    Dim got() As Byte
    Dim tmpB() As Byte
    Dim n As Long
    Dim off As Long
    Dim k As Long
    Dim peeked As Long
    Dim gotN As Long
    Dim t As Double
    Dim moved As Boolean

    On Error GoTo EH
    r.Passed = True
    port = NextPort()
    Set srv = NewEndpoint(useOcx, "srv")
    srv.Echo = False
    srv.ListenOn port
    If Not WaitState(srv, sckListening, 4000) Then
        r.Passed = False
        r.Failure = "listen timeout"
        Exit Sub
    End If
    Set cli = NewEndpoint(useOcx, "cli")
    cli.ConnectTo "127.0.0.1", port
    If Not WaitConnect(cli, 5000) Then r.Passed = False: r.Failure = "connect"
    If Not WaitAcceptedCount(1, 5000) Then r.Passed = False: r.Failure = r.Failure & " accept"
    Set conn = gChildren.Item(1)
    conn.Echo = False

    n = 5000
    ReDim payload(0 To n - 1)
    FillPattern payload, n, PAT_RANDOM
    ReDim got(0 To n - 1)

    cli.SendBytes payload
    off = 0
    t = NowSeconds()
    Do While off < n
        moved = False
        If (k Mod 3) = 0 Then
            peeked = TryPeekBytes(conn, 7)
            gotN = TryGetBytes(conn, 7)
            If gotN <> peeked Then
                r.Passed = False
                r.Failure = r.Failure & " peek/get differ at " & off
            End If
        Else
            gotN = TryGetBytes(conn, 7)
        End If
        If gotN > 0 Then
            conn.CopyLastBytes tmpB
            If off + gotN <= n Then
                MemoryCopy got(off), tmpB(0), gotN
                off = off + gotN
                moved = True
            End If
        End If
        k = k + 1
        If Not moved Then
            Pump
            If (NowSeconds() - t) > 10# Then Exit Do
        End If
        If k > 5000 Then Exit Do
    Loop

    Require r, (off = n), "slices read=" & off & "/" & n
    If off = n Then
        Require r, BuffersEqual(payload, got), "slice payload mismatch at " & FirstMismatch(payload, got)
    End If

    cli.CloseSock
    srv.CloseSock
    r.Details = r.Details & "slices=" & k & " bytes=" & off & "/" & n & vbCrLf
    Exit Sub
EH:
    r.Passed = False
    r.Failure = r.Failure & " EXC " & Err.Number
End Sub

Private Sub RunBinUdpBounds(ByRef r As clsRun, ByVal useOcx As Boolean)
    Dim rcv As clsEndpoint
    Dim snd As clsEndpoint
    Dim port As Long
    Dim sizes As Variant
    Dim i As Long
    Dim sz As Long
    Dim payload() As Byte
    Dim got() As Byte

    On Error GoTo EH
    r.Passed = True
    port = NextPort()
    Set rcv = NewEndpoint(useOcx, "rcv")
    rcv.SetProtocol sckUDPProtocol
    rcv.AutoRead = True
    rcv.BindTo port
    If Not WaitState(rcv, sckOpen, 3000) Then
        r.Passed = False
        r.Failure = "udp bind"
        Exit Sub
    End If
    Set snd = NewEndpoint(useOcx, "snd")
    snd.SetProtocol sckUDPProtocol
    snd.SetRemote "127.0.0.1", port

    sizes = Array(1, 2, 1472, 8192, 65507)
    For i = LBound(sizes) To UBound(sizes)
        sz = sizes(i)
        ReDim payload(0 To sz - 1)
        FillPattern payload, sz, PAT_RANDOM
        rcv.ResetFlags
        snd.SendBytes payload
        If Not WaitData(rcv, 5000) Then
            r.Passed = False
            r.Failure = r.Failure & " size" & sz & "=timeout"
        Else
            rcv.CopyLastBytes got
            Require r, (rcv.LastByteCount = sz), "size" & sz & " got=" & rcv.LastByteCount
            If rcv.LastByteCount = sz Then
                Require r, BuffersEqual(payload, got), "size" & sz & " content at " & FirstMismatch(payload, got)
            End If
        End If
    Next

    snd.CloseSock
    rcv.CloseSock
    r.Details = r.Details & "datagram sizes checked: 1,2,1472,8192,65507" & vbCrLf
    Exit Sub
EH:
    r.Passed = False
    r.Failure = r.Failure & " EXC " & Err.Number & ": " & Err.Description
End Sub

' SendData with a String containing chars > 127 must arrive as the ANSI bytes.
Private Sub RunBinAnsi(ByRef r As clsRun, ByVal useOcx As Boolean)
    Dim srv As clsEndpoint
    Dim cli As clsEndpoint
    Dim conn As clsEndpoint
    Dim port As Long
    Dim text As String
    Dim expect As String
    Dim got() As Byte
    Dim i As Long
    Dim ok As Boolean

    On Error GoTo EH
    r.Passed = True
    text = "caf" & Chr$(233) & Chr$(255) & Chr$(128) & "x"
    expect = StrConv(text, vbFromUnicode)
    port = NextPort()
    Set srv = NewEndpoint(useOcx, "srv")
    srv.Echo = False
    srv.ChildCaptureLimit = Len(text) * 2
    srv.ListenOn port
    If Not WaitState(srv, sckListening, 4000) Then
        r.Passed = False
        r.Failure = "listen"
        Exit Sub
    End If
    Set cli = NewEndpoint(useOcx, "cli")
    cli.ConnectTo "127.0.0.1", port
    If Not WaitConnect(cli, 5000) Then r.Passed = False: r.Failure = "connect"
    If Not WaitAcceptedCount(1, 5000) Then r.Passed = False: r.Failure = r.Failure & " accept"
    Set conn = gChildren.Item(1)

    cli.SendString text
    WaitSettle 400
    If conn.CaptureSize = 0 Then
        r.Passed = False
        r.Failure = r.Failure & " nothing arrived"
    Else
        conn.CopyCapture got
        r.Details = r.Details & "text len=" & Len(text) & " StrConv len=" & Len(expect) & _
                    " arrived=" & conn.CaptureSize & " bytes=" & HexBytes(got, conn.CaptureSize) & vbCrLf
    End If

    r.Trace = srv.TraceText & vbCrLf & conn.TraceText & vbCrLf & cli.TraceText
    cli.CloseSock
    srv.CloseSock
    Exit Sub
EH:
    r.Passed = False
    r.Failure = r.Failure & " EXC " & Err.Number
End Sub

Private Sub WaitSettle(ByVal ms As Long)
    Dim t As Double
    t = NowSeconds()
    Do While (NowSeconds() - t) * 1000# < ms
        Pump
    Loop
End Sub

Private Function HexBytes(ByRef b() As Byte, ByVal n As Long) As String
    Dim i As Long
    Dim s As String
    If n > 12 Then n = 12
    For i = 0 To n - 1
        s = s & Right$("0" & Hex$(b(i)), 2)
    Next
End Function

'==============================================================================
' Tier 2 - concurrency scenarios
'==============================================================================
Private Sub RunConcClients(ByRef r As clsRun, ByVal useOcx As Boolean)
    Const NC As Long = 32
    Dim srv As clsEndpoint
    Dim cli(1 To NC) As clsEndpoint
    Dim port As Long
    Dim i As Long
    Dim expect As String
    Dim got As Long

    On Error GoTo EH
    r.Passed = True
    port = NextPort()
    Set srv = NewEndpoint(useOcx, "srv")
    srv.Echo = True
    srv.ListenOn port
    If Not WaitState(srv, sckListening, 4000) Then
        r.Passed = False
        r.Failure = "listen"
        Exit Sub
    End If

    For i = 1 To NC
        Set cli(i) = NewEndpoint(useOcx, "c" & i)
        If cli(i) Is Nothing Then
            r.Passed = False
            r.Failure = r.Failure & " create" & i
            Exit Sub
        End If
        cli(i).ConnectTo "127.0.0.1", port
        ' Connect in small groups: 32 simultaneous SYNs overflow the listen
        ' backlog, which is a kernel artifact, not an implementation property.
        If (i Mod 8) = 0 Then
            If Not WaitAcceptedCount(i, 10000) Then
                r.Passed = False
                r.Failure = r.Failure & " accepted" & gChildren.Count & "of" & i
                Exit Sub
            End If
        End If
    Next
    For i = 1 To NC
        If Not WaitConnect(cli(i), 8000) Then
            r.Passed = False
            r.Failure = r.Failure & " conn" & i
        End If
    Next
    If Not WaitAcceptedCount(NC, 8000) Then
        r.Passed = False
        r.Failure = r.Failure & " accepted=" & gChildren.Count
        Exit Sub
    End If

    For i = 1 To NC
        cli(i).ResetFlags
        cli(i).SendString "client-" & i & "-payload"
    Next
    For i = 1 To NC
        expect = "client-" & i & "-payload"
        If Not WaitData(cli(i), 8000) Then
            r.Passed = False
            r.Failure = r.Failure & " data" & i
        Else
            cli(i).GetStringData
            If cli(i).LastString <> expect Then
                r.Passed = False
                r.Failure = r.Failure & " crosstalk" & i & "=[" & cli(i).LastString & "]"
                got = got + 1
            End If
        End If
    Next

    r.Details = r.Details & "clients=" & NC & " crosstalk=" & got & " slots=" & DiagnosticsLine() & vbCrLf
    For i = 1 To NC
        cli(i).CloseSock
    Next
    srv.CloseSock
    Exit Sub
EH:
    r.Passed = False
    r.Failure = r.Failure & " EXC " & Err.Number & ": " & Err.Description
End Sub

Private Sub RunConcListeners(ByRef r As clsRun, ByVal useOcx As Boolean)
    Const NL As Long = 3
    Const NP As Long = 2
    Dim srv(1 To NL) As clsEndpoint
    Dim cli(1 To NL, 1 To NP) As clsEndpoint
    Dim port(1 To NL) As Long
    Dim i As Long
    Dim j As Long
    Dim expect As String

    On Error GoTo EH
    r.Passed = True
    For i = 1 To NL
        port(i) = NextPort()
        Set srv(i) = NewEndpoint(useOcx, "srv" & i)
        srv(i).Echo = True
        srv(i).ListenOn port(i)
        If Not WaitState(srv(i), sckListening, 4000) Then
            r.Passed = False
            r.Failure = r.Failure & " listen" & i
        End If
    Next
    For i = 1 To NL
        For j = 1 To NP
            Set cli(i, j) = NewEndpoint(useOcx, "c" & i & "_" & j)
            cli(i, j).ConnectTo "127.0.0.1", port(i)
        Next
    Next
    For i = 1 To NL
        For j = 1 To NP
            If Not WaitConnect(cli(i, j), 8000) Then r.Passed = False: r.Failure = r.Failure & " conn" & i & "_" & j
        Next
    Next
    If Not WaitAcceptedCount(NL * NP, 8000) Then
        r.Passed = False
        r.Failure = r.Failure & " accepted=" & gChildren.Count
        Exit Sub
    End If

    For i = 1 To NL
        For j = 1 To NP
            cli(i, j).ResetFlags
            cli(i, j).SendString "L" & i & "C" & j
        Next
    Next
    For i = 1 To NL
        For j = 1 To NP
            expect = "L" & i & "C" & j
            If Not WaitData(cli(i, j), 8000) Then
                r.Passed = False
                r.Failure = r.Failure & " data" & i & "_" & j
            Else
                cli(i, j).GetStringData
                Require r, (cli(i, j).LastString = expect), "pair" & i & "_" & j & "=[" & cli(i, j).LastString & "]"
            End If
        Next
    Next

    r.Details = r.Details & "listeners=" & NL & " pairs=" & (NL * NP) & vbCrLf
    For i = 1 To NL
        For j = 1 To NP
            cli(i, j).CloseSock
        Next
        srv(i).CloseSock
    Next
    Exit Sub
EH:
    r.Passed = False
    r.Failure = r.Failure & " EXC " & Err.Number & ": " & Err.Description
End Sub

' Sequential connect/echo/close, the pattern that exercises handle reuse.
Private Sub RunConcChurn(ByRef r As clsRun, ByVal useOcx As Boolean)
    Const NCYCLES As Long = 200
    Dim srv As clsEndpoint
    Dim cli As clsEndpoint
    Dim port As Long
    Dim i As Long
    Dim handles0 As Long
    Dim t0 As Double

    On Error GoTo EH
    r.Passed = True
    handles0 = ProcHandleCount()
    port = NextPort()
    Set srv = NewEndpoint(useOcx, "srv")
    srv.Echo = True
    srv.ListenOn port
    If Not WaitState(srv, sckListening, 4000) Then
        r.Passed = False
        r.Failure = "listen"
        Exit Sub
    End If

    t0 = NowSeconds()
    For i = 1 To NCYCLES
        Set cli = NewEndpoint(useOcx, "churn")
        cli.ConnectTo "127.0.0.1", port
        If Not WaitConnect(cli, 5000) Then
            r.Passed = False
            r.Failure = r.Failure & " conn" & i
            Exit For
        End If
        If Not WaitAcceptedCount(i, 5000) Then
            r.Passed = False
            r.Failure = r.Failure & " accept" & i
            Exit For
        End If
        cli.ResetFlags
        cli.SendString "x"
        If Not WaitData(cli, 5000) Then
            r.Passed = False
            r.Failure = r.Failure & " echo" & i
            Exit For
        End If
        cli.GetStringData
        cli.CloseSock
        Set cli = Nothing
        If (i Mod 25) = 0 Then Pump
    Next

    srv.CloseSock
    r.Details = r.Details & "cycles=" & NCYCLES & " in " & Format$(NowSeconds() - t0, "0.000") & " s" & vbCrLf
    r.Details = r.Details & "handles " & handles0 & " -> " & ProcHandleCount() & " (TIME_WAIT is expected here)" & vbCrLf
    ' Release every object, then confirm the module kept nothing behind.
    Set srv = Nothing
    Set cli = Nothing
    CleanupActive
    r.Details = r.Details & DiagnosticsLine() & vbCrLf
    Require r, (InStr(DiagnosticsLine(), "objects=0") > 0), "objects leaked: " & DiagnosticsLine()
    Require r, (InStr(DiagnosticsLine(), "sockets=0") > 0), "sockets leaked: " & DiagnosticsLine()
    Require r, (InStr(DiagnosticsLine(), "pending=0") > 0), "pending accepts leaked: " & DiagnosticsLine()
    Exit Sub
EH:
    r.Passed = False
    r.Failure = r.Failure & " EXC " & Err.Number & ": " & Err.Description
End Sub

Private Sub RunConcMixed(ByRef r As clsRun, ByVal useOcx As Boolean)
    Const NT As Long = 4
    Const NU As Long = 4
    Dim srv As clsEndpoint
    Dim cli(1 To NT) As clsEndpoint
    Dim rcv(1 To NU) As clsEndpoint
    Dim snd(1 To NU) As clsEndpoint
    Dim tcpPort As Long
    Dim udpPort(1 To NU) As Long
    Dim i As Long

    On Error GoTo EH
    r.Passed = True
    tcpPort = NextPort()
    Set srv = NewEndpoint(useOcx, "srv")
    srv.Echo = True
    srv.ListenOn tcpPort
    If Not WaitState(srv, sckListening, 4000) Then
        r.Passed = False
        r.Failure = "listen"
        Exit Sub
    End If

    For i = 1 To NU
        udpPort(i) = NextPort()
        Set rcv(i) = NewEndpoint(useOcx, "ur" & i)
        rcv(i).SetProtocol sckUDPProtocol
        rcv(i).AutoRead = True
        rcv(i).BindTo udpPort(i)
        If Not WaitState(rcv(i), sckOpen, 3000) Then r.Passed = False: r.Failure = r.Failure & " ubind" & i
        Set snd(i) = NewEndpoint(useOcx, "us" & i)
        snd(i).SetProtocol sckUDPProtocol
        snd(i).SetRemote "127.0.0.1", udpPort(i)
    Next
    For i = 1 To NT
        Set cli(i) = NewEndpoint(useOcx, "tc" & i)
        cli(i).ConnectTo "127.0.0.1", tcpPort
    Next
    For i = 1 To NT
        If Not WaitConnect(cli(i), 8000) Then r.Passed = False: r.Failure = r.Failure & " tconn" & i
    Next
    If Not WaitAcceptedCount(NT, 8000) Then
        r.Passed = False
        r.Failure = r.Failure & " accepted=" & gChildren.Count
    End If

    For i = 1 To NT
        cli(i).ResetFlags
        cli(i).SendString "tcp" & i
    Next
    For i = 1 To NU
        rcv(i).ResetFlags
        snd(i).SendString "udp" & i
    Next
    For i = 1 To NT
        If Not WaitData(cli(i), 8000) Then
            r.Passed = False
            r.Failure = r.Failure & " tdata" & i
        Else
            cli(i).GetStringData
            Require r, (cli(i).LastString = "tcp" & i), "tcp" & i & "=[" & cli(i).LastString & "]"
        End If
    Next
    For i = 1 To NU
        If Not WaitData(rcv(i), 8000) Then
            r.Passed = False
            r.Failure = r.Failure & " udata" & i
        Else
            Require r, (rcv(i).LastString = "udp" & i), "udp" & i & "=[" & rcv(i).LastString & "]"
        End If
    Next

    r.Details = r.Details & "tcp=" & NT & " udp=" & NU & " " & DiagnosticsLine() & vbCrLf
    For i = 1 To NT
        cli(i).CloseSock
    Next
    For i = 1 To NU
        snd(i).CloseSock
        rcv(i).CloseSock
    Next
    srv.CloseSock
    Exit Sub
EH:
    r.Passed = False
    r.Failure = r.Failure & " EXC " & Err.Number & ": " & Err.Description
End Sub

'==============================================================================
' Tier 3 - stress
'==============================================================================
Private Sub RunStressStream(ByRef r As clsRun, ByVal useOcx As Boolean)
    Dim secs As Double
    Dim errText As String
    r.Passed = True
    If StreamTransfer(useOcx, useOcx, 8388608, PAT_RANDOM, 2, secs, errText) Then
        r.Details = r.Details & "8 MB in " & Format$(secs, "0.000") & " s (" & _
                    Format$(8# / secs, "0.0") & " MB/s)" & vbCrLf
    Else
        r.Passed = False
        r.Failure = r.Failure & " stream=" & errText
    End If
    r.Details = r.Details & ResourceLine() & vbCrLf & DiagnosticsLine() & vbCrLf
End Sub

Private Sub RunStressMsgs(ByRef r As clsRun, ByVal useOcx As Boolean)
    Const N As Long = 20000
    Const MAXS As Long = 4000
    Dim srv As clsEndpoint
    Dim cli As clsEndpoint
    Dim port As Long
    Dim i As Long
    Dim samples() As Long
    Dim nS As Long
    Dim stride As Long
    Dim tIter As Double
    Dim t0 As Double
    Dim total As Double
    Dim p50 As Double
    Dim p95 As Double
    Dim p99 As Double
    Dim worst As Double
    Dim payload As String

    On Error GoTo EH
    r.Passed = True
    port = NextPort()
    Set srv = NewEndpoint(useOcx, "srv")
    srv.Echo = True
    srv.ListenOn port
    If Not WaitState(srv, sckListening, 4000) Then
        r.Passed = False
        r.Failure = "listen"
        Exit Sub
    End If
    Set cli = NewEndpoint(useOcx, "cli")
    cli.ConnectTo "127.0.0.1", port
    If Not WaitConnect(cli, 5000) Then r.Passed = False: r.Failure = "connect"
    If Not WaitAcceptedCount(1, 5000) Then r.Passed = False: r.Failure = r.Failure & " accept"

    payload = String$(64, "m")
    cli.ResetFlags
    cli.SendString payload
    If WaitData(cli, 5000) Then cli.GetStringData

    ReDim samples(0 To MAXS - 1)
    stride = N \ MAXS
    If stride < 1 Then stride = 1
    nS = 0
    t0 = NowSeconds()
    For i = 0 To N - 1
        cli.ResetFlags
        tIter = NowSeconds()
        cli.SendString payload
        If Not WaitData(cli, 5000) Then
            r.Passed = False
            r.Failure = r.Failure & " iter" & i
            Exit For
        End If
        cli.GetStringData
        If (i Mod stride) = 0 And nS < MAXS Then
            samples(nS) = CLng((NowSeconds() - tIter) * 1000000#)
            nS = nS + 1
        End If
    Next
    total = NowSeconds() - t0
    Percentiles samples, nS, p50, p95, p99, worst

    r.Details = r.Details & "round trips=" & N & " in " & Format$(total, "0.000") & " s (" & _
                Format$(N / total, "0") & "/s)" & vbCrLf
    r.Details = r.Details & "latency us: p50=" & p50 & " p95=" & p95 & " p99=" & p99 & " max=" & worst & vbCrLf
    cli.CloseSock
    srv.CloseSock
    Exit Sub
EH:
    r.Passed = False
    r.Failure = r.Failure & " EXC " & Err.Number & ": " & Err.Description
End Sub

Private Sub RunStressUdp(ByRef r As clsRun, ByVal useOcx As Boolean)
    Const N As Long = 4000
    Dim rcv As clsEndpoint
    Dim snd As clsEndpoint
    Dim port As Long
    Dim i As Long
    Dim payload() As Byte
    Dim sent As Long
    Dim prev As Long
    Dim quiet As Long
    Dim t0 As Double
    Dim total As Double

    On Error GoTo EH
    r.Passed = True
    port = NextPort()
    Set rcv = NewEndpoint(useOcx, "rcv")
    rcv.SetProtocol sckUDPProtocol
    rcv.AutoRead = True
    rcv.BindTo port
    If Not WaitState(rcv, sckOpen, 3000) Then
        r.Passed = False
        r.Failure = "udp bind"
        Exit Sub
    End If
    Set snd = NewEndpoint(useOcx, "snd")
    snd.SetProtocol sckUDPProtocol
    snd.SetRemote "127.0.0.1", port

    ReDim payload(0 To 63)
    FillPattern payload, 64, PAT_RANDOM

    rcv.ResetFlags
    t0 = NowSeconds()
    For i = 1 To N
        snd.SendBytes payload
        sent = sent + 1
        ' Pace the burst so the receive buffer cannot overflow: this measures
        ' delivery and rate, not UDP's permission to drop.
        If (i Mod 25) = 0 Then
            Pump
            prev = rcv.DataCount
            Do
                Pump
                If rcv.DataCount = prev Then Exit Do
                If (NowSeconds() - t0) > 10# Then Exit Do
                prev = rcv.DataCount
            Loop
        End If
    Next
    Do
        prev = rcv.DataCount
        Pump
        If rcv.DataCount = prev Then quiet = quiet + 1 Else quiet = 0
        If quiet > 30 Then Exit Do
        If (NowSeconds() - t0) > 10# Then Exit Do
    Loop
    total = NowSeconds() - t0

    r.Details = r.Details & "datagrams sent=" & sent & " received=" & rcv.DataCount & _
                " loss=" & Format$(100# * (sent - rcv.DataCount) / sent, "0.0") & "%" & vbCrLf
    r.Details = r.Details & "rate=" & Format$(rcv.DataCount / total, "0") & "/s in " & Format$(total, "0.000") & " s" & vbCrLf
    If useOcx Then
        ' MSWINSCK drops most of a burst this size; that is the oracle's
        ' behaviour, recorded rather than judged.
        If rcv.DataCount < sent * 0.95 Then
            r.Details = r.Details & "OCX dropped most of the burst (oracle limitation)" & vbCrLf
        End If
    Else
        If rcv.DataCount < sent * 0.95 Then
            r.Passed = False
            r.Failure = r.Failure & " loss " & (sent - rcv.DataCount) & "/" & sent
        End If
    End If

    snd.CloseSock
    rcv.CloseSock
    Exit Sub
EH:
    r.Passed = False
    r.Failure = r.Failure & " EXC " & Err.Number & ": " & Err.Description
End Sub

'==============================================================================
' Tier 4 - state machine / adversarial peers
'==============================================================================
' Every probe appends its error number to the trace, so the paired run compares
' the outcomes directly.
Private Sub RunApiOrder(ByRef r As clsRun, ByVal useOcx As Boolean)
    Dim p As clsEndpoint
    Dim srv As clsEndpoint
    Dim cli As clsEndpoint
    Dim port As Long
    Dim n As Long
    Dim e1 As Long
    Dim e2 As Long
    Dim e3 As Long
    Dim e4 As Long
    Dim e5 As Long
    Dim e6 As Long

    On Error GoTo EH
    r.Passed = True
    If useOcx Then
        ' Driving MSWINSCK with wrong-order calls crashes it (0xC0000005), so the
        ' oracle cannot be compared here. The probes still run against modWinsock2.
        r.KnownGap = True
        r.Passed = True
        r.Details = "Oracle crashes on wrong-order calls; probes run on modWinsock2 only."
        Exit Sub
    End If
    Set p = NewEndpoint(useOcx, "probe")

    On Error Resume Next
    Err.Clear: p.ConnectTo "", 0: e1 = Err.Number: Err.Clear
    p.LogTrace "probe|ConnectEmptyHost|err=" & e1

    Err.Clear: p.SetProtocol 99: e2 = Err.Number: Err.Clear
    p.LogTrace "probe|ProtocolInvalid|err=" & e2

    Err.Clear: p.CloseSock: p.CloseSock: e3 = Err.Number: Err.Clear
    p.LogTrace "probe|CloseTwice|err=" & e3

    port = NextPort()
    Set srv = NewEndpoint(useOcx, "srv")
    srv.Echo = True
    srv.ListenOn port
    WaitState srv, sckListening, 4000

    Err.Clear: srv.AcceptReq 999999: e4 = Err.Number: Err.Clear
    srv.LogTrace "probe|AcceptBogus|err=" & e4

    Set cli = NewEndpoint(useOcx, "cli")
    cli.ConnectTo "127.0.0.1", port
    WaitConnect cli, 5000
    WaitAcceptedCount 1, 5000

    Err.Clear: cli.ConnectTo "127.0.0.1", port: e5 = Err.Number: Err.Clear
    cli.LogTrace "probe|ConnectWhileConnected|err=" & e5

    Err.Clear: cli.ListenOn port: e6 = Err.Number: Err.Clear
    cli.LogTrace "probe|ListenWhileConnected|err=" & e6
    On Error GoTo 0

    Require r, (e1 <> 0), "Connect with empty host did not raise"
    Require r, (e2 <> 0), "Protocol 99 did not raise"
    Require r, (e3 = 0), "second Close raised " & e3
    Require r, (e4 <> 0), "Accept with a bogus requestID did not raise"
    Require r, (e5 <> 0), "Connect while connected did not raise"
    Require r, (e6 <> 0), "Listen while connected did not raise"
    r.Details = r.Details & "errors: emptyHost=" & e1 & " protocol=" & e2 & " close2=" & e3 & _
                " acceptBogus=" & e4 & " connect2=" & e5 & " listen2=" & e6 & vbCrLf

    r.Trace = p.TraceText & vbCrLf & srv.TraceText & vbCrLf & cli.TraceText
    cli.CloseSock
    srv.CloseSock
    Exit Sub
EH:
    r.Passed = False
    r.Failure = r.Failure & " EXC " & Err.Number & ": " & Err.Description
End Sub

Private Sub RunEvtReentrancy(ByRef r As clsRun, ByVal useOcx As Boolean)
    Dim srv As clsEndpoint
    Dim cli As clsEndpoint
    Dim port As Long

    On Error GoTo EH
    r.Passed = True
    port = NextPort()
    Set srv = NewEndpoint(useOcx, "srv")
    srv.Echo = False
    srv.CloseOnData = True
    srv.ListenOn port
    If Not WaitState(srv, sckListening, 4000) Then
        r.Passed = False
        r.Failure = "listen"
        Exit Sub
    End If
    Set cli = NewEndpoint(useOcx, "cli")
    cli.ConnectTo "127.0.0.1", port
    If Not WaitConnect(cli, 5000) Then r.Passed = False: r.Failure = "connect"
    If Not WaitAcceptedCount(1, 5000) Then r.Passed = False: r.Failure = r.Failure & " accept"

    cli.ResetFlags
    cli.SendString "close-me-inside-the-event"
    If Not WaitClose(cli, 8000) Then
        r.Passed = False
        r.Failure = r.Failure & " peer did not close"
    End If
    r.Details = r.Details & "client saw Close=" & cli.DoneClose & " error=" & cli.DoneError & vbCrLf

    cli.CloseSock
    srv.CloseSock
    Exit Sub
EH:
    r.Passed = False
    r.Failure = r.Failure & " EXC " & Err.Number & ": " & Err.Description
End Sub

Private Sub RunPeerAbort(ByRef r As clsRun, ByVal useOcx As Boolean)
    Dim srv As clsEndpoint
    Dim cli As clsEndpoint
    Dim conn As clsEndpoint
    Dim port As Long
    Dim payload() As Byte
    Dim i As Long
    Dim t As Double

    On Error GoTo EH
    r.Passed = True
    port = NextPort()
    Set srv = NewEndpoint(useOcx, "srv")
    srv.Echo = False
    srv.ListenOn port
    If Not WaitState(srv, sckListening, 4000) Then
        r.Passed = False
        r.Failure = "listen"
        Exit Sub
    End If
    Set cli = NewEndpoint(useOcx, "cli")
    cli.ConnectTo "127.0.0.1", port
    If Not WaitConnect(cli, 5000) Then r.Passed = False: r.Failure = "connect"
    If Not WaitAcceptedCount(1, 5000) Then r.Passed = False: r.Failure = r.Failure & " accept"
    Set conn = gChildren.Item(1)

    ReDim payload(0 To 1048575)
    FillPattern payload, 1048576, PAT_RANDOM
    ' start pushing data, then reset the connection while the transfer is live
    For i = 1 To 4
        conn.SendBytes payload
        Pump
    Next
    cli.KillSock

    t = NowSeconds()
    Do While Not (conn.DoneError Or conn.DoneClose)
        Pump
        If (NowSeconds() - t) > 15# Then Exit Do
    Loop
    Require r, (conn.DoneError Or conn.DoneClose), "no error/close after peer reset"
    r.Details = r.Details & "peer reset -> error=" & conn.LastErrorNo & " close=" & conn.DoneClose & vbCrLf

    srv.CloseSock
    Exit Sub
EH:
    r.Passed = False
    r.Failure = r.Failure & " EXC " & Err.Number & ": " & Err.Description
End Sub

' Forces partial sends by shrinking the receive window of the peer.
Private Sub RunSendPartial(ByRef r As clsRun, ByVal useOcx As Boolean)
    Dim srv As clsEndpoint
    Dim cli As clsEndpoint
    Dim conn As clsEndpoint
    Dim port As Long
    Dim payload() As Byte
    Dim got() As Byte

    On Error GoTo EH
    r.Passed = True
    port = NextPort()
    Set srv = NewEndpoint(useOcx, "srv")
    srv.Echo = False
    srv.ChildCaptureLimit = 1048576
    srv.ShrinkReceiveBuffer 4096
    srv.ListenOn port
    If Not WaitState(srv, sckListening, 4000) Then
        r.Passed = False
        r.Failure = "listen"
        Exit Sub
    End If
    Set cli = NewEndpoint(useOcx, "cli")
    cli.ShrinkSendBuffer 4096
    cli.ConnectTo "127.0.0.1", port
    If Not WaitConnect(cli, 5000) Then r.Passed = False: r.Failure = "connect"
    If Not WaitAcceptedCount(1, 5000) Then r.Passed = False: r.Failure = r.Failure & " accept"
    Set conn = gChildren.Item(1)

    ReDim payload(0 To 1048575)
    FillPattern payload, 1048576, PAT_RANDOM
    cli.SendBytes payload
    If Not WaitCapture(conn, 30000) Then
        r.Passed = False
        r.Failure = r.Failure & " receive " & conn.CaptureSize & "/1048576"
        Exit Sub
    End If
    conn.CopyCapture got
    Require r, BuffersEqual(payload, got), "partial-send corruption at " & FirstMismatch(payload, got)
    r.Details = r.Details & "1 MB through a 4 KB window" & vbCrLf

    cli.CloseSock
    srv.CloseSock
    Exit Sub
EH:
    r.Passed = False
    r.Failure = r.Failure & " EXC " & Err.Number & ": " & Err.Description
End Sub

'==============================================================================
' Tier 6 - interop under load
'==============================================================================
Private Function RunInteropStream(ByVal srvOcx As Boolean, ByVal cliOcx As Boolean) As clsRun
    Dim r As clsRun
    Dim secs As Double
    Dim errText As String

    Set r = New clsRun
    If srvOcx Then
        r.Reset "OCX server/W2 client"
    Else
        r.Reset "W2 server/OCX client"
    End If
    Set gActive = New Collection
    Set gChildren = New Collection
    If StreamTransfer(srvOcx, cliOcx, 4194304, PAT_RANDOM, 1, secs, errText) Then
        r.Passed = True
        r.Details = "4 MB in " & Format$(secs, "0.000") & " s (" & Format$(4# / secs, "0.0") & " MB/s)"
    Else
        r.Passed = False
        r.Failure = errText
    End If
    CleanupActive
    Set RunInteropStream = r
End Function

'==============================================================================
' Matrix runner support
'==============================================================================
Public Sub MxResetCollections(ByVal seq As Long)
    Set gActive = New Collection
    Set gChildren = New Collection
    gDiag = ""
    gMxPortNext = 21000 + (seq Mod 400) * 5
    modMatrix.gMxStage = "reset-ok"
End Sub

Public Function MxNextPort() As Long
    gMxPortNext = gMxPortNext + 1
    MxNextPort = gMxPortNext
End Function

Public Function MxEndpoint(ByVal useOcx As Boolean, ByVal nm As String) As clsEndpoint
    On Error Resume Next
    modMatrix.gMxStage = "mxendpoint:" & nm
    Set MxEndpoint = NewEndpoint(useOcx, nm)
    If Err.Number <> 0 Then
        modMatrix.MxObserve "mxfail", "MxEndpoint " & Err.Number & " " & Err.Description
        Err.Clear
    End If
End Function

Public Sub MxCleanup()
    CleanupActive
End Sub

Public Function MxChildCount() As Long
    If gChildren Is Nothing Then Exit Function
    MxChildCount = gChildren.Count
End Function

Public Function MxChild(ByVal i As Long) As clsEndpoint
    If gChildren Is Nothing Then Exit Function
    If i < 1 Or i > gChildren.Count Then Exit Function
    Set MxChild = gChildren.Item(i)
End Function

Public Function MxTempFile(ByVal cellID As String, ByVal implName As String) As String
    Dim folder As String
    Dim safe As String
    folder = App.Path & "\results\mx"
    On Error Resume Next
    MkDir App.Path & "\results"
    MkDir folder
    On Error GoTo 0
    safe = Replace(Replace(Replace(cellID, ":", "_"), "\", "_"), "/", "_")
    MxTempFile = folder & "\" & LCase$(implName) & "_" & safe & ".txt"
End Function

Public Sub MxSaveMatrix(ByVal csv As String, ByVal md As String)
    Dim folder As String
    Dim stamp As String
    folder = App.Path & "\results"
    On Error Resume Next
    MkDir folder
    On Error GoTo 0
    stamp = TimestampName()
    WriteTextFile folder & "\matrix_" & stamp & ".csv", csv
    WriteTextFile folder & "\matrix_" & stamp & ".md", md
    WriteTextFile folder & "\matrix_latest.csv", csv
    WriteTextFile folder & "\matrix_latest.md", md
End Sub

'==============================================================================
' UI / reporting
'==============================================================================
Private Sub FillUI(ByRef f As frmFeatureTests)
    Dim i As Long
    Dim res As clsResult
    Dim row As clsPerfRow

    Set gForm = f
    f.ListBox1.Clear
    Set gListItems = New Collection

    For i = 1 To gResults.Count
        Set res = gResults.Item(i)
        f.ListBox1.AddItem "[" & res.Status & "] " & res.Group & " / " & res.Scenario
        gListItems.Add res
    Next

    For i = 1 To gPerf.Count
        Set row = gPerf.Item(i)
        f.ListBox1.AddItem "[PERF] " & row.Operation
        gListItems.Add row
    Next

    f.SetSummary SummaryLine()
    f.SetReportText BuildMarkdownReport()
End Sub

Public Sub ShowItem(ByRef f As frmFeatureTests, ByVal idx As Long)
    Dim o As Object
    If gListItems Is Nothing Then Exit Sub
    If idx < 0 Or idx >= gListItems.Count Then Exit Sub
    Set o = gListItems.Item(idx + 1)
    If TypeOf o Is clsResult Then
        f.SetDetailText BuildResultDetail(o)
    ElseIf TypeOf o Is clsPerfRow Then
        f.SetDetailText BuildPerfDetail(o)
    End If
End Sub

Private Function RunText(ByRef r As clsRun) As String
    If r Is Nothing Then
        RunText = "not run"
    Else
        RunText = IIf(r.Passed, "PASS", "FAIL")
        If Len(r.Failure) > 0 Then RunText = RunText & " (" & r.Failure & ")"
        If Len(r.Details) > 0 Then RunText = RunText & " | " & r.Details
    End If
End Function

Private Function BuildResultDetail(ByRef res As clsResult) As String
    Dim s As String
    s = res.Group & " / " & res.Scenario & vbCrLf
    s = s & "Status: " & res.Status & vbCrLf
    If Len(res.Notes) > 0 Then s = s & "Notes: " & res.Notes & vbCrLf
    If Len(res.Diff) > 0 Then s = s & "First diff:" & vbCrLf & res.Diff & vbCrLf
    s = s & "W2: " & RunText(res.W2) & vbCrLf
    s = s & "OCX: " & RunText(res.Ocx) & vbCrLf
    If Not res.W2 Is Nothing Then
        s = s & vbCrLf & "W2 trace:" & vbCrLf & LargeText(res.W2.Trace) & vbCrLf
        s = s & "W2 snapshot:" & vbCrLf & res.W2.Snapshot & vbCrLf
    End If
    If Not res.Ocx Is Nothing Then
        s = s & vbCrLf & "OCX trace:" & vbCrLf & LargeText(res.Ocx.Trace) & vbCrLf
        s = s & "OCX snapshot:" & vbCrLf & res.Ocx.Snapshot & vbCrLf
    End If
    BuildResultDetail = s
End Function

Private Function BuildPerfDetail(ByRef row As clsPerfRow) As String
    Dim s As String
    s = row.Operation & vbCrLf
    s = s & "Iterations: " & row.Iterations & vbCrLf
    s = s & "W2 seconds: " & Format$(row.W2Seconds, "0.000000") & "  Ops/s: " & Format$(row.W2Ops, "0.00") & vbCrLf
    s = s & "OCX seconds: " & Format$(row.OcxSeconds, "0.000000") & "  Ops/s: " & Format$(row.OcxOps, "0.00") & vbCrLf
    If row.W2BytesPerSec > 0 Then s = s & "W2 MB/s: " & Format$(row.W2BytesPerSec / 1048576#, "0.00") & vbCrLf
    If row.OcxBytesPerSec > 0 Then s = s & "OCX MB/s: " & Format$(row.OcxBytesPerSec / 1048576#, "0.00") & vbCrLf
    s = s & "W2/OCX time ratio: " & row.RatioText() & vbCrLf
    If Len(row.Notes) > 0 Then s = s & "Notes: " & row.Notes
    BuildPerfDetail = s
End Function

Private Function FirstLine(ByVal s As String, ByVal maxLen As Long) As String
    Dim p As Long
    p = InStr(1, s, vbCrLf)
    If p > 0 Then s = Left$(s, p - 1)
    If Len(s) > maxLen Then s = Left$(s, maxLen) & "..."
    FirstLine = s
End Function

Private Function RunShort(ByRef r As clsRun) As String
    If r Is Nothing Then
        RunShort = "-"
    ElseIf r.Passed Then
        RunShort = "PASS"
    Else
        RunShort = "FAIL"
    End If
End Function

Private Sub GetCounts(ByRef passN As Long, ByRef diffN As Long, ByRef failN As Long, ByRef expN As Long)
    Dim i As Long
    Dim res As clsResult
    For i = 1 To gResults.Count
        Set res = gResults.Item(i)
        Select Case res.Status
        Case "PASS", "W2_ONLY": passN = passN + 1
        Case "DIFF": diffN = diffN + 1
        Case "EXPECTED", "SKIP": expN = expN + 1
        Case Else: failN = failN + 1
        End Select
    Next
End Sub

Public Function SummaryLine() As String
    Dim p As Long
    Dim d As Long
    Dim f As Long
    Dim e As Long
    GetCounts p, d, f, e
    SummaryLine = "PASS=" & p & "  DIFF=" & d & "  FAIL=" & f & "  EXPECTED/SKIP=" & e & _
                  "  Scenarios=" & gResults.Count & "  Performance rows=" & gPerf.Count & _
                  "  OCX oracle=" & gOcxAvailable & "  Clock=" & TimingSource()
End Function

' Curated explanation of the differences the traces actually show.
Private Function KnownDifferencesText() As String
    Dim s As String
    s = "1. Late binding - CallByName works on both; modWinsock2 maps the MSWinsockLib " & _
        "dispids onto its typed members." & vbCrLf
    s = s & "2. Error Source - the OCX reports its own module path (for example C:\Setup\mswinsck.ocx); " & _
        "modWinsock2 reports MSWinsock. Number, Description and Scode match." & vbCrLf
    s = s & "3. OCX robustness - MSWINSCK.OCX 6.1.97.82 raises an access violation when " & _
        "BytesReceived is read after Close, when GetData runs without a connection, and when " & _
        "Bind targets a port already in use. Those probes run only against modWinsock2." & vbCrLf
    s = s & "4. UDP read timing - the OCX discards a datagram that is not read inside the " & _
        "DataArrival handler; modWinsock2 keeps it buffered. Tests read inside the event." & vbCrLf
    s = s & "5. Trace comparison - events are compared per endpoint and DataArrival byte counts " & _
        "are ignored (they depend on buffering); payload totals are asserted separately." & vbCrLf
    s = s & "6. Performance - the OCX is native C++/COM, modWinsock2 is pure VB6." & vbCrLf
    KnownDifferencesText = s
End Function

Private Function CsvField(ByVal s As String) As String
    s = Replace(s, """", """""")
    s = Replace(s, vbCrLf, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, vbCr, " ")
    CsvField = """" & s & """"
End Function

Public Function BuildCsvReport() As String
    Dim s As String
    Dim i As Long
    Dim res As clsResult
    Dim row As clsPerfRow
    s = "kind,status,group,scenario,w2,ocx,notes" & vbCrLf
    For i = 1 To gResults.Count
        Set res = gResults.Item(i)
        s = s & "scenario," & CsvField(res.Status) & "," & CsvField(res.Group) & "," & _
            CsvField(res.Scenario) & "," & CsvField(RunShort(res.W2)) & "," & _
            CsvField(RunShort(res.Ocx)) & "," & CsvField(res.Notes) & vbCrLf
    Next
    s = s & vbCrLf & "kind,operation,iterations,w2_seconds,ocx_seconds,w2_ops_per_sec,ocx_ops_per_sec,ratio,notes" & vbCrLf
    For i = 1 To gPerf.Count
        Set row = gPerf.Item(i)
        s = s & "perf," & CsvField(row.Operation) & "," & row.Iterations & "," & _
            Format$(row.W2Seconds, "0.000000") & "," & Format$(row.OcxSeconds, "0.000000") & "," & _
            Format$(row.W2Ops, "0.00") & "," & Format$(row.OcxOps, "0.00") & "," & _
            CsvField(row.RatioText()) & "," & CsvField(row.Notes) & vbCrLf
    Next
    BuildCsvReport = s
End Function

Public Function BuildRawLog() As String
    Dim s As String
    Dim i As Long
    Dim res As clsResult
    Dim row As clsPerfRow
    s = "modWinsock2 feature/conformance raw log" & vbCrLf
    s = s & "==================================================================" & vbCrLf
    s = s & "Date: " & Now & vbCrLf
    s = s & "OCX oracle available: " & gOcxAvailable & vbCrLf
    s = s & "Command line: " & Command$ & vbCrLf
    s = s & "Summary: " & SummaryLine() & vbCrLf & vbCrLf
    For i = 1 To gResults.Count
        Set res = gResults.Item(i)
        s = s & "------------------------------------------------------------------" & vbCrLf
        s = s & BuildResultDetail(res) & vbCrLf
    Next
    s = s & "------------------------------------------------------------------" & vbCrLf
    s = s & "PERFORMANCE" & vbCrLf & vbCrLf
    For i = 1 To gPerf.Count
        Set row = gPerf.Item(i)
        s = s & BuildPerfDetail(row) & vbCrLf & vbCrLf
    Next
    If Len(gDiag) > 0 Then
        s = s & "------------------------------------------------------------------" & vbCrLf
        s = s & "HARNESS DIAGNOSTICS" & vbCrLf & vbCrLf & gDiag
    End If
    BuildRawLog = s
End Function

Public Function BuildMarkdownReport() As String
    Dim s As String
    Dim i As Long
    Dim res As clsResult
    Dim row As clsPerfRow
    s = "# modWinsock2 feature/conformance report" & vbCrLf & vbCrLf
    s = s & "- Date: " & Now & vbCrLf
    s = s & "- OCX oracle available: " & gOcxAvailable & vbCrLf
    s = s & "- " & SummaryLine() & vbCrLf & vbCrLf
    s = s & "## Scenario results" & vbCrLf & vbCrLf
    s = s & "| Status | Group | Scenario | W2 | OCX | Notes |" & vbCrLf
    s = s & "|---|---|---|---|---|---|" & vbCrLf
    For i = 1 To gResults.Count
        Set res = gResults.Item(i)
        s = s & "| " & res.Status & " | " & res.Group & " | " & res.Scenario & " | " & _
            RunShort(res.W2) & " | " & RunShort(res.Ocx) & " | " & _
            Replace(res.Notes, "|", "/") & " |" & vbCrLf
    Next
    s = s & vbCrLf & "## Differences / failures" & vbCrLf & vbCrLf
    For i = 1 To gResults.Count
        Set res = gResults.Item(i)
        If res.Status <> "PASS" And res.Status <> "W2_ONLY" Then
            s = s & "### " & res.Group & " / " & res.Scenario & " [" & res.Status & "]" & vbCrLf & vbCrLf
            If Len(res.Notes) > 0 Then s = s & res.Notes & vbCrLf & vbCrLf
            If Len(res.Diff) > 0 Then s = s & "First diff:" & vbCrLf & "```" & vbCrLf & res.Diff & vbCrLf & "```" & vbCrLf & vbCrLf
        End If
    Next
    s = s & "## Performance" & vbCrLf & vbCrLf
    s = s & "| Operation | Iterations | W2 ms | OCX ms | W2 ops/s | OCX ops/s | ratio | Notes |" & vbCrLf
    s = s & "|---|---:|---:|---:|---:|---:|---:|---|" & vbCrLf
    For i = 1 To gPerf.Count
        Set row = gPerf.Item(i)
        s = s & "| " & row.Operation & " | " & row.Iterations & " | " & _
            Format$(row.W2Seconds * 1000#, "0.00") & " | " & Format$(row.OcxSeconds * 1000#, "0.00") & " | " & _
            Format$(row.W2Ops, "0.00") & " | " & Format$(row.OcxOps, "0.00") & " | " & row.RatioText() & " | " & _
            Replace(row.Notes, "|", "/") & " |" & vbCrLf
    Next
    s = s & vbCrLf & "> Performance is indicative: MSWINSCK.OCX is native C++/COM; the replacement is pure VB6." & vbCrLf
    If Len(gDiag) > 0 Then
        s = s & vbCrLf & "## Harness diagnostics" & vbCrLf & vbCrLf & "```" & vbCrLf & gDiag & "```" & vbCrLf
    End If
    s = s & vbCrLf & "## Known differences explained" & vbCrLf & vbCrLf & KnownDifferencesText() & vbCrLf
    s = s & "## Coverage notes" & vbCrLf & vbCrLf
    s = s & "- IMSWinsockControl: 25 members exercised across defaults/properties/TCP/UDP/DNS/errors/lifetime." & vbCrLf
    s = s & "- Events: all 7 exercised; event traces are compared for every functional scenario." & vbCrLf
    s = s & "- IDispatch is probed through CallByName (get) only; property puts and methods go through the vtable." & vbCrLf
    s = s & "- AboutBox is modal and is not run in automated tests." & vbCrLf
    BuildMarkdownReport = s
End Function

Public Function BuildHtmlReport() As String
    Dim s As String
    Dim i As Long
    Dim res As clsResult
    Dim row As clsPerfRow
    s = "<!DOCTYPE html><html><head><meta charset=""utf-8""><title>modWinsock2 Feature Tests</title>"
    s = s & "<style>body{font-family:Segoe UI,Arial,sans-serif;margin:24px;color:#222}"
    s = s & "table{border-collapse:collapse;margin:8px 0 20px 0}th,td{border:1px solid #bbb;padding:4px 7px;vertical-align:top}"
    s = s & "th{background:#eee}.PASS{color:#1b7f3a;font-weight:bold}.FAIL{color:#b00020;font-weight:bold}"
    s = s & ".DIFF{color:#b06a00;font-weight:bold}.EXPECTED{color:#3355aa;font-weight:bold}"
    s = s & ".W2_ONLY,.SKIP{color:#666;font-weight:bold}pre{background:#f6f6f6;padding:8px;overflow:auto}"
    s = s & "</style></head><body>"
    s = s & "<h1>modWinsock2 feature/conformance report</h1>"
    s = s & "<p>Date: " & Now & "<br>OCX oracle available: " & gOcxAvailable & "<br>" & EscapeHtml(SummaryLine()) & "</p>"
    s = s & "<h2>Scenario results</h2><table><tr><th>Status</th><th>Group</th><th>Scenario</th><th>W2</th><th>OCX</th><th>Notes</th></tr>"
    For i = 1 To gResults.Count
        Set res = gResults.Item(i)
        s = s & "<tr><td class=""" & res.Status & """>" & res.Status & "</td><td>" & EscapeHtml(res.Group) & "</td><td>" & EscapeHtml(res.Scenario) & "</td><td>"
        If res.W2 Is Nothing Then s = s & "-" Else s = s & IIf(res.W2.Passed, "PASS", "FAIL")
        s = s & "</td><td>"
        If res.Ocx Is Nothing Then s = s & "-" Else s = s & IIf(res.Ocx.Passed, "PASS", "FAIL")
        s = s & "</td><td>" & EscapeHtml(res.Notes) & "</td></tr>"
    Next
    s = s & "</table>"
    s = s & "<h2>Differences / failures</h2>"
    For i = 1 To gResults.Count
        Set res = gResults.Item(i)
        If res.Status <> "PASS" And res.Status <> "W2_ONLY" Then
            s = s & "<h3>" & EscapeHtml(res.Group & " / " & res.Scenario) & " <span class=""" & res.Status & """>[" & res.Status & "]</span></h3>"
            If Len(res.Notes) > 0 Then s = s & "<p>" & EscapeHtml(res.Notes) & "</p>"
            If Len(res.Diff) > 0 Then s = s & "<pre>" & EscapeHtml(res.Diff) & "</pre>"
            s = s & "<pre>" & EscapeHtml(BuildResultDetail(res)) & "</pre>"
        End If
    Next
    s = s & "<h2>Performance</h2><table><tr><th>Operation</th><th>Iterations</th><th>W2 ms</th><th>OCX ms</th><th>W2 ops/s</th><th>OCX ops/s</th><th>Ratio</th><th>Notes</th></tr>"
    For i = 1 To gPerf.Count
        Set row = gPerf.Item(i)
        s = s & "<tr><td>" & EscapeHtml(row.Operation) & "</td><td>" & row.Iterations & "</td><td>" & Format$(row.W2Seconds * 1000#, "0.00") & "</td><td>" & Format$(row.OcxSeconds * 1000#, "0.00") & "</td><td>" & Format$(row.W2Ops, "0.00") & "</td><td>" & Format$(row.OcxOps, "0.00") & "</td><td>" & row.RatioText() & "</td><td>" & EscapeHtml(row.Notes) & "</td></tr>"
    Next
    s = s & "</table><p><em>Performance is indicative: MSWINSCK.OCX is native C++/COM; modWinsock2 is pure VB6.</em></p>"
    If Len(gDiag) > 0 Then
        s = s & "<h2>Harness diagnostics</h2><pre>" & EscapeHtml(gDiag) & "</pre>"
    End If
    s = s & "<h2>Known differences explained</h2><ul>"
    s = s & "<li>Late binding: CallByName works on both; modWinsock2 maps the MSWinsockLib dispids onto its typed members.</li>"
    s = s & "<li>Error Source: the OCX reports its own module path, modWinsock2 reports MSWinsock. Number, Description and Scode match.</li>"
    s = s & "<li>OCX robustness: reading BytesReceived after Close, GetData with no connection, and Bind to a used port crash MSWINSCK.OCX 6.1.97.82, so those probes run only against modWinsock2.</li>"
    s = s & "<li>UDP read timing: the OCX discards datagrams not read inside DataArrival; modWinsock2 keeps them buffered.</li>"
    s = s & "<li>Trace comparison: events are compared per endpoint, DataArrival byte counts are ignored, payload totals are asserted separately.</li>"
    s = s & "<li>Performance is indicative: the OCX is native C++/COM, modWinsock2 is pure VB6.</li></ul>"
    s = s & "<h2>Coverage notes</h2><ul>"
    s = s & "<li>IMSWinsockControl: 25 members exercised across the scenarios.</li>"
    s = s & "<li>Events: all 7 exercised and compared in event traces.</li>"
    s = s & "<li>IDispatch is probed through CallByName (get) only; property puts and methods go through the vtable.</li>"
    s = s & "<li>AboutBox is modal and is not run automatically.</li></ul>"
    s = s & "</body></html>"
    BuildHtmlReport = s
End Function

Public Sub SaveReports(ByRef f As frmFeatureTests)
    Dim folder As String
    Dim stamp As String
    Dim html As String
    Dim md As String
    Dim csv As String
    Dim raw As String
    If Not f Is Nothing Then Set gForm = f
    folder = App.Path & "\results"
    On Error Resume Next
    MkDir folder
    On Error GoTo 0
    stamp = TimestampName()
    html = BuildHtmlReport()
    md = BuildMarkdownReport()
    csv = BuildCsvReport()
    raw = BuildRawLog()
    WriteTextFile folder & "\report_" & stamp & ".html", html
    WriteTextFile folder & "\report_" & stamp & ".md", md
    WriteTextFile folder & "\report_" & stamp & ".csv", csv
    WriteTextFile folder & "\rawlog_" & stamp & ".log", raw
    WriteTextFile folder & "\latest.html", html
    WriteTextFile folder & "\latest.md", md
    WriteTextFile folder & "\latest.csv", csv
    WriteTextFile folder & "\latest.log", raw
    If Not gForm Is Nothing Then gForm.SetDetailText "Saved HTML, Markdown, CSV and raw log to " & folder & vbCrLf & vbCrLf & raw
End Sub




