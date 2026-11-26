Attribute VB_Name = "modMatrix"
Option Explicit

'==============================================================================
' Differential conformance matrix
'
' Every cell is one observable behaviour, executed in its own process so that an
' MSWINSCK access violation or modal dialog is recorded as an observation
' instead of aborting the sweep.
'
'   parent: FeatureTests.exe /matrix[:<filter>]
'   child : FeatureTests.exe /cell:<id> /impl:<w2|ocx> /out:<file> [/seq:<n>]
'
' The child writes one key=value record; the parent compares w2 against ocx and
' emits matrix.csv plus divergences.md.
'==============================================================================

Private Const MX_TIMEOUT_MS As Long = 25000
Private Const STILL_ACTIVE   As Long = &H103&

Public gMxRows As Collection
Private gMxFindPid As Long
Private gMxFindHwnd As Long

Private Type MX_SECURITY_ATTRIBUTES
    nLength              As Long
    lpSecurityDescriptor As Long
    bInheritHandle       As Boolean
End Type

Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, _
    ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, _
    ByRef lpExitCode As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, _
    ByVal uExitCode As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, _
    ByVal lParam As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, _
    ByRef lpdwProcessId As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, _
    ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, _
    ByVal lpString As String, ByVal nMaxCount As Long) As Long

'------------------------------------------------------------------------------
' Observation record (shared with modMatrixCells)
'------------------------------------------------------------------------------
Public gMxOcx     As Boolean
Public gMxPort    As Long
Public gMxObs     As String
Public gMxSubject As clsEndpoint
Public gMxServer  As clsEndpoint
Public gMxConn    As clsEndpoint
Public gMxSeq     As Long
Public gMxStage   As String

Public Sub MxReset(ByVal seq As Long)
    gMxSeq = seq
    gMxPort = 21000 + (seq Mod 400) * 5
    gMxObs = ""
    Set gMxSubject = Nothing
    Set gMxServer = Nothing
    Set gMxConn = Nothing
    modTestEngine.MxResetCollections seq
End Sub

Public Sub MxObserve(ByVal k As String, ByVal v As String)
    v = Replace(v, vbCrLf, " | ")
    v = Replace(v, vbLf, " | ")
    v = Replace(v, vbCr, " | ")
    gMxObs = gMxObs & k & "=" & v & vbCrLf
End Sub

Public Function MxNextPort() As Long
    MxNextPort = modTestEngine.MxNextPort()
End Function

Public Function MxNew(ByVal nm As String) As clsEndpoint
    On Error Resume Next
    gMxStage = "mxnew:" & nm
    Set MxNew = modTestEngine.MxEndpoint(gMxOcx, nm)
    If Err.Number <> 0 Then
        MxObserve "mxfail", "MxNew " & Err.Number & " " & Err.Description
        Err.Clear
    End If
End Function

Public Function MxWaitState(ByRef c As clsEndpoint, ByVal want As Long, ByVal ms As Long) As Boolean
    Dim t As Double
    t = modTestUtil.NowSeconds()
    Do While c.StateValue <> want
        modTestUtil.Pump
        If (modTestUtil.NowSeconds() - t) * 1000# > ms Then Exit Function
    Loop
    MxWaitState = True
End Function

Public Function MxWaitConnect(ByRef c As clsEndpoint, ByVal ms As Long) As Boolean
    Dim t As Double
    t = modTestUtil.NowSeconds()
    Do While Not c.DoneConnect
        modTestUtil.Pump
        If (modTestUtil.NowSeconds() - t) * 1000# > ms Then Exit Function
    Loop
    MxWaitConnect = True
End Function

Public Function MxWaitData(ByRef c As clsEndpoint, ByVal ms As Long) As Boolean
    Dim t As Double
    t = modTestUtil.NowSeconds()
    Do While Not c.DoneData
        modTestUtil.Pump
        If (modTestUtil.NowSeconds() - t) * 1000# > ms Then Exit Function
    Loop
    MxWaitData = True
End Function

Public Function MxWaitError(ByRef c As clsEndpoint, ByVal ms As Long) As Boolean
    Dim t As Double
    t = modTestUtil.NowSeconds()
    Do While Not c.DoneError
        modTestUtil.Pump
        If (modTestUtil.NowSeconds() - t) * 1000# > ms Then Exit Function
    Loop
    MxWaitError = True
End Function

Public Function MxWaitClose(ByRef c As clsEndpoint, ByVal ms As Long) As Boolean
    Dim t As Double
    t = modTestUtil.NowSeconds()
    Do While Not c.DoneClose
        modTestUtil.Pump
        If (modTestUtil.NowSeconds() - t) * 1000# > ms Then Exit Function
    Loop
    MxWaitClose = True
End Function

Public Function MxWaitAccepted(ByVal n As Long, ByVal ms As Long) As Boolean
    Dim t As Double
    t = modTestUtil.NowSeconds()
    Do While modTestEngine.MxChildCount < n
        modTestUtil.Pump
        If (modTestUtil.NowSeconds() - t) * 1000# > ms Then Exit Function
    Loop
    MxWaitAccepted = True
End Function

Public Sub MxSettle(ByVal ms As Long)
    Dim t As Double
    t = modTestUtil.NowSeconds()
    Do While (modTestUtil.NowSeconds() - t) * 1000# < ms
        modTestUtil.Pump
    Loop
End Sub

' Volatile values are masked so the two implementations can be compared.
Public Function MxMask(ByVal prop As String, ByVal v As String) As String
    Select Case LCase$(prop)
    Case "sockethandle"
        If v = "-1" Then MxMask = "-1" Else MxMask = "<handle>"
    Case "localport", "remoteport"
        If v = "0" Then MxMask = "0" Else MxMask = "<port>"
    Case "remotehostip", "localip"
        MxMask = v
    Case Else
        MxMask = v
    End Select
End Function

Public Function MxTraceText(ByRef eps As Collection) As String
    Dim i As Long
    Dim s As String
    On Error Resume Next
    For i = 1 To eps.Count
        s = s & modTestUtil.NormalizeTrace(eps.Item(i).TraceText)
    Next
    MxTraceText = s
End Function

Public Sub MxFinish()
    On Error Resume Next
    modTestEngine.MxCleanup
End Sub

'------------------------------------------------------------------------------
' Cell list
'------------------------------------------------------------------------------
Public Function CellIDs() As Collection
    Dim c As Collection
    Dim st As Variant
    Dim pr As Variant
    Set c = New Collection

    ' property reads: 6 states x 10 properties
    For Each st In Array("closed", "listening", "connected", "open", "connecting", "error")
        For Each pr In Array("Protocol", "SocketHandle", "LocalPort", "LocalIP", "LocalHostName", _
                             "RemotePort", "RemoteHost", "RemoteHostIP", "BytesReceived", "State")
            c.Add "pg:" & st & ":" & pr
        Next
    Next

    ' property sets
    For Each st In Array("closed", "listening", "connected", "open")
        For Each pr In Array("Protocol=2", "Protocol=99", "LocalPort=65536", "RemotePort=65536", _
                             "RemotePort=0", "RemoteHost=", "RemoteHost=example.com", "RemotePort=1")
            c.Add "ps:" & st & ":" & pr
        Next
    Next

    Dim v As Variant
    For Each v In Array("nohost", "hostonly", "portonly", "port0", "port65536", "refused", _
                        "blackhole", "name", "ownlistener", "whileConnected", "whileListening", _
                        "afterClose", "emptyName")
        c.Add "conn:" & v
    Next
    For Each v In Array("port0", "inuse", "twice", "whileConnected", "afterBind", "closeWithPending")
        c.Add "listen:" & v
    Next
    For Each v In Array("valid", "bogus", "duplicate", "foreign", "afterClose", "notListening")
        c.Add "accept:" & v
    Next
    For Each v In Array("port0", "specific", "inuse", "twice", "afterConnect", "localIP", "onTCP")
        c.Add "bind:" & v
    Next
    For Each v In Array("emptyStr", "emptyArr", "one", "64k", "1m", "unicode", "longs", "null", _
                        "nothing", "closed", "connecting", "listening", "udpNoRemoteHost", _
                        "udpNoRemotePort", "udp65507", "udp65508", "afterReset", "whileGetPending")
        c.Add "send:" & v
    Next
    For Each v In Array("nodata", "closed", "maxlen0", "maxlen1", "maxlenHuge", "maxlenNeg", _
                        "typeString", "typeByte", "typeArray", "typeLong", "twoDatagrams", _
                        "peekThenGet", "inDataArrival")
        c.Add "get:" & v
    Next
    For Each v In Array("closed", "udp", "listeningPending", "connecting", "connected", "midSend", _
                        "midResolve", "inDataArrival", "twice", "reuseAfterClose")
        c.Add "close:" & v
    Next
    For Each v In Array("ok", "bad", "closeDuring", "connectDuring", "concurrent8", "twice", _
                        "literal", "localhost", "empty")
        c.Add "dns:" & v
    Next
    For Each v In Array("dataCoalescing", "sendProgressTcp", "sendProgressUdp", "error10061", _
                        "connReqOrder", "closeLocal", "closeRemote", "closeReset", _
                        "reentrancyData", "reentrancyConnReq")
        c.Add "evt:" & v
    Next
    For Each v In Array("fin", "rstMidSend", "rstMidRecv", "halfClose", "partialWindow", _
                        "stream16m", "rstThenSend", "rstThenGet", "rstThenReuse")
        c.Add "abort:" & v
    Next
    For Each v In Array("clients32", "listeners3", "mixed", "churn", "orderUnderLoad", "reuseSlots")
        c.Add "conc:" & v
    Next

    Set CellIDs = c
End Function

'------------------------------------------------------------------------------
' Child: execute one cell and write the record
'------------------------------------------------------------------------------
Public Function RunCell(ByVal cellID As String, ByVal implName As String, ByVal outPath As String) As Boolean
    Dim ocx As Boolean
    Dim seq As Long
    Dim p As Long
    Dim text As String

    ocx = (LCase$(implName) = "ocx")
    p = InStr(1, cellID, ":")
    seq = 0
    If p > 1 Then seq = CLng(Asc(Mid$(cellID, p + 1, 1))) * 7 + Len(cellID)

    gMxStage = "start"
    MxReset seq
    gMxOcx = ocx

    On Error GoTo Failed
    Call modMatrixCells.ExecuteCell(cellID, ocx)
    On Error GoTo 0
    MxFinish
    If Len(modTestEngine.gDiag) > 0 Then MxObserve "diag", modTestEngine.gDiag

    text = "cell=" & cellID & vbCrLf & "impl=" & LCase$(implName) & vbCrLf & gMxObs
    modTestUtil.WriteTextFile outPath, text
    RunCell = True
    Exit Function
Failed:
    MxObserve "exception", CStr(Err.Number) & " " & Err.Description
    MxObserve "stage", gMxStage
    If Len(modTestEngine.gDiag) > 0 Then MxObserve "diag", modTestEngine.gDiag
    modTestUtil.WriteTextFile outPath, "cell=" & cellID & vbCrLf & "impl=" & LCase$(implName) & vbCrLf & gMxObs
    RunCell = False
End Function

'------------------------------------------------------------------------------
' Parent: run every cell twice and compare
'------------------------------------------------------------------------------
Public Sub RunMatrix(ByRef f As frmFeatureTests, ByVal filterText As String)
    Dim ids As Collection
    Dim i As Long
    Dim cellID As String
    Dim w2 As String
    Dim ox As String
    Dim w2State As String
    Dim oxState As String
    Dim rows As Collection
    Dim t0 As Double

    Set gMxRows = New Collection
    Set ids = CellIDs()
    t0 = modTestUtil.NowSeconds()

    For i = 1 To ids.Count
        cellID = ids.Item(i)
        If Len(filterText) = 0 Or InStr(1, cellID, filterText, vbTextCompare) > 0 Then
            w2 = RunChild(cellID, "w2", w2State)
            ox = RunChild(cellID, "ocx", oxState)
            AddRow cellID, w2, w2State, ox, oxState
            If Not f Is Nothing Then
                f.SetSummary "Matrix " & i & "/" & ids.Count & "  " & cellID & _
                             "  (" & Format$(modTestUtil.NowSeconds() - t0, "0") & " s)"
            End If
        End If
    Next

    WriteMatrix f
End Sub

Private Function RunChild(ByVal cellID As String, ByVal implName As String, ByRef state As String) As String
    Dim outPath As String
    Dim exe As String
    Dim pid As Long
    Dim h As Long
    Dim code As Long
    Dim deadline As Double
    Dim dlg As String
    Dim seq As Long

    outPath = modTestEngine.MxTempFile(cellID, implName)
    On Error Resume Next
    Kill outPath
    On Error GoTo 0

    exe = App.Path & "\" & App.EXEName & ".exe"
    seq = 0
    pid = Shell("""" & exe & """ /cell:" & cellID & " /impl:" & implName & " /out:""" & outPath & """", vbHide)
    If pid = 0 Then
        state = "launch-failed"
        RunChild = ""
        Exit Function
    End If

    deadline = modTestUtil.NowSeconds() + MX_TIMEOUT_MS / 1000#
    Do
        modTestUtil.Pump
        dlg = DialogText(pid)
        If Len(dlg) > 0 Then
            state = "dialog:" & dlg
            KillPid pid
            Exit Do
        End If
        h = OpenProcess(&H1000& Or &H400&, 0&, pid)   ' QUERY_LIMITED + SYNCHRONIZE
        If h <> 0 Then
            If GetExitCodeProcess(h, code) <> 0 Then
                CloseHandle h
                If code <> STILL_ACTIVE Then
                    state = "exit:" & code
                    Exit Do
                End If
            End If
        End If
        If modTestUtil.NowSeconds() > deadline Then
            state = "timeout"
            KillPid pid
            Exit Do
        End If
    Loop

    If Not modTestUtil.FileExists(outPath) Then
        If Len(state) = 0 Then state = "no-output"
        RunChild = ""
        Exit Function
    End If
    RunChild = modTestUtil.ReadTextFile(outPath)
    If Len(state) = 0 Then state = "ok"
End Function

Private Sub KillPid(ByVal pid As Long)
    Dim h As Long
    h = OpenProcess(&H1&, 0&, pid)   ' PROCESS_TERMINATE
    If h <> 0 Then
        TerminateProcess h, 1&
        CloseHandle h
    End If
End Sub

Private Function DialogText(ByVal pid As Long) As String
    Dim txt As String * 256
    Dim n As Long
    Dim found As String

    gMxFindPid = pid
    gMxFindHwnd = 0
    EnumWindows AddressOf EnumProc, 0&
    If gMxFindHwnd = 0 Then Exit Function
    txt = Space$(255)
    n = GetWindowText(gMxFindHwnd, txt, 256)
    found = Left$(txt, n)
    If Len(found) = 0 Then found = "(dialog)"
    DialogText = found
End Function

Private Function EnumProc(ByVal hWnd As Long, ByVal lParam As Long) As Long
    Dim p As Long
    Dim cls As String * 64
    Dim n As Long

    If gMxFindHwnd <> 0 Then
        EnumProc = 0
        Exit Function
    End If
    GetWindowThreadProcessId hWnd, p
    If p = gMxFindPid Then
        cls = Space$(63)
        n = GetClassName(hWnd, cls, 64)
        If n > 0 Then
            If Left$(cls, n) = "#32770" Then
                gMxFindHwnd = hWnd
                EnumProc = 0
                Exit Function
            End If
        End If
    End If
    EnumProc = 1
End Function

'------------------------------------------------------------------------------
' Comparison / output
'------------------------------------------------------------------------------

Public Sub AddRow(ByVal cellID As String, ByVal w2 As String, ByVal w2State As String, _
                  ByVal ox As String, ByVal oxState As String)
    Dim row As clsCell
    Set row = New clsCell
    row.CellID = cellID
    row.W2State = w2State
    row.OcxState = oxState
    row.W2Text = w2
    row.OcxText = ox
    row.Verdict = ClassifyCell(w2, w2State, ox, oxState)
    row.Diff = DiffText(Comparable(w2), Comparable(ox))
    gMxRows.Add row
End Sub

Private Function ExitFailed(ByVal state As String) As Boolean
    If Left$(state, 5) <> "exit:" Then Exit Function
    ExitFailed = (Mid$(state, 6) <> "0")
End Function

' The impl name is part of the record but not of the comparison.
Private Function Comparable(ByVal s As String) As String
    Dim lines() As String
    Dim i As Long
    Dim out As String
    If Len(s) = 0 Then Exit Function
    lines = Split(s, vbCrLf)
    For i = 0 To UBound(lines)
        If Left$(lines(i), 5) <> "impl=" And Len(lines(i)) > 0 Then out = out & lines(i) & vbCrLf
    Next
    Comparable = out
End Function
Private Function ClassifyCell(ByVal w2 As String, ByVal w2State As String, _
                               ByVal ox As String, ByVal oxState As String) As String
    Dim oxCrashed As Boolean
    Dim w2Bad As Boolean

    oxCrashed = (Left$(oxState, 6) = "dialog" Or oxState = "timeout" Or ExitFailed(oxState))
    w2Bad = (Len(w2) = 0 Or Left$(w2State, 6) = "dialog" Or w2State = "timeout" Or ExitFailed(w2State))
    If w2Bad Then
        ClassifyCell = "W2_BROKEN"
    ElseIf oxCrashed Then
        ClassifyCell = "ORACLE_" & UCase$(Left$(oxState, 6))
    ElseIf Len(ox) = 0 Then
        ClassifyCell = "ORACLE_NO_OUTPUT"
    ElseIf Comparable(w2) = Comparable(ox) Then
        ClassifyCell = "MATCH"
    Else
        ClassifyCell = "DIVERGENCE"
    End If
End Function

Public Function DiffText(ByVal a As String, ByVal b As String) As String
    Dim la() As String
    Dim lb() As String
    Dim i As Long
    Dim n As Long
    Dim s As String

    la = Split(a, vbCrLf)
    lb = Split(b, vbCrLf)
    n = UBound(la)
    If UBound(lb) > n Then n = UBound(lb)
    For i = 0 To n
        If i > UBound(la) Then
            s = s & "only oracle: " & lb(i) & vbCrLf
        ElseIf i > UBound(lb) Then
            s = s & "only w2: " & la(i) & vbCrLf
        ElseIf la(i) <> lb(i) Then
            s = s & "w2   : " & la(i) & vbCrLf & "oracle: " & lb(i) & vbCrLf
        End If
    Next
    If Len(s) > 1200 Then s = Left$(s, 1200) & "..."
    DiffText = s
End Function

Public Sub WriteMatrix(ByRef f As frmFeatureTests)
    Dim s As String
    Dim csvText As String
    Dim md As String
    Dim i As Long
    Dim row As clsCell
    Dim matchN As Long
    Dim divN As Long
    Dim oracleN As Long
    Dim brokenN As Long

    For i = 1 To gMxRows.Count
        Set row = gMxRows.Item(i)
        Select Case row.Verdict
        Case "MATCH": matchN = matchN + 1
        Case "DIVERGENCE": divN = divN + 1
        Case "W2_BROKEN": brokenN = brokenN + 1
        Case Else: oracleN = oracleN + 1
        End Select
    Next

    s = "# Differential matrix vs MSWINSCK" & vbCrLf & vbCrLf
    s = s & "- Date: " & Now & vbCrLf
    s = s & "- Cells: " & gMxRows.Count & "  MATCH=" & matchN & "  DIVERGENCE=" & divN & _
            "  ORACLE-LIMITED=" & oracleN & "  W2-BROKEN=" & brokenN & vbCrLf & vbCrLf
    s = s & "| Cell | Verdict | Oracle | Diff |" & vbCrLf & "|---|---|---|---|" & vbCrLf
    For i = 1 To gMxRows.Count
        Set row = gMxRows.Item(i)
        s = s & "| " & row.CellID & " | " & row.Verdict & " | " & row.OcxState & " | " & _
            Replace(Replace(row.Diff, vbCrLf, "<br>"), "|", "/") & " |" & vbCrLf
    Next

    csvText = "cell,verdict,w2_state,ocx_state,diff" & vbCrLf
    For i = 1 To gMxRows.Count
        Set row = gMxRows.Item(i)
        csvText = csvText & Csv(row.CellID) & "," & Csv(row.Verdict) & "," & Csv(row.W2State) & "," & _
              Csv(row.OcxState) & "," & Csv(row.Diff) & vbCrLf
    Next

    md = s
    modTestEngine.MxSaveMatrix csvText, md
    If Not f Is Nothing Then
        f.SetReportText md
    End If
End Sub

Private Function Csv(ByVal s As String) As String
    s = Replace(s, """", """""")
    s = Replace(s, vbCrLf, " / ")
    s = Replace(s, vbLf, " / ")
    Csv = """" & s & """"
End Function
