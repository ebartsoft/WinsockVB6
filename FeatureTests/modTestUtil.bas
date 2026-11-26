Attribute VB_Name = "modTestUtil"
Option Explicit

Private Declare Function GetProcessHandleCount Lib "kernel32" _
    (ByVal hProcess As Long, ByRef pdwHandleCount As Long) As Long
Private Declare Function GetGuiResources Lib "user32" _
    (ByVal hProcess As Long, ByVal uiFlags As Long) As Long
Private Declare Function setsockopt Lib "ws2_32.dll" _
    (ByVal s As Long, ByVal level As Long, ByVal optname As Long, ByRef optval As Long, ByVal optlen As Long) As Long
Private Declare Function w2_shutdown Lib "ws2_32.dll" Alias "shutdown" _
    (ByVal s As Long, ByVal how As Long) As Long

Private Const W2_SOL_SOCKET   As Long = &HFFFF&
Private Const W2_SO_LINGER    As Long = &H80&
Private Const W2_SO_SNDBUF    As Long = &H1001&
Private Const W2_SO_RCVBUF    As Long = &H1002&
Private Const W2_GR_GDIOBJECTS As Long = 0
Private Const W2_GR_USEROBJECTS As Long = 1

Public Const PAT_RANDOM  As Long = 0
Public Const PAT_ZERO    As Long = 1
Public Const PAT_FF      As Long = 2
Public Const PAT_COUNT   As Long = 3
Public Const PAT_NULLMIX As Long = 4

' QPC returns a 64-bit tick count. VB6 has no Int64, and Currency only keeps
' 4 decimal places (which would quantise the counter to 1 ms), so read the raw
' halves into a UDT instead. The tlb already declares these names, hence Alias.
Private Type LARGE_INTEGER
    LowPart As Long
    HighPart As Long
End Type

Private Declare Function QpcCounter Lib "kernel32" Alias "QueryPerformanceCounter" _
    (ByRef lpPerformanceCount As LARGE_INTEGER) As Long
Private Declare Function QpcFrequency Lib "kernel32" Alias "QueryPerformanceFrequency" _
    (ByRef lpFrequency As LARGE_INTEGER) As Long

Private mFreqTicks As Double
Private mFreqReady As Boolean

Private Function TicksToDouble(ByRef li As LARGE_INTEGER) As Double
    Dim lo As Double
    lo = CDbl(li.LowPart)
    If li.LowPart < 0 Then lo = lo + 4294967296#   ' treat the low half as unsigned
    TicksToDouble = CDbl(li.HighPart) * 4294967296# + lo
End Function

'------------------------------------------------------------------------------
' High-resolution time
'------------------------------------------------------------------------------
Public Function NowSeconds() As Double
    Dim c As LARGE_INTEGER
    If Not mFreqReady Then
        mFreqTicks = 0
        If QpcFrequency(c) <> 0 Then mFreqTicks = TicksToDouble(c)
        mFreqReady = True
    End If
    If mFreqTicks > 0 Then
        If QpcCounter(c) <> 0 Then
            NowSeconds = TicksToDouble(c) / mFreqTicks
            Exit Function
        End If
    End If
    NowSeconds = Timer
End Function

Public Function ElapsedMs(ByVal startSec As Double, ByVal endSec As Double) As Double
    ElapsedMs = (endSec - startSec) * 1000#
End Function

' Describes the clock actually backing NowSeconds (QPC when available).
Public Function TimingSource() As String
    Dim c As LARGE_INTEGER
    If mFreqReady Then
        If mFreqTicks > 0 Then
            TimingSource = "QueryPerformanceCounter (" & Format$(mFreqTicks, "0") & " Hz)"
        Else
            TimingSource = "VB Timer (millisecond resolution)"
        End If
        Exit Function
    End If
    Dim f As LARGE_INTEGER
    If QpcFrequency(f) <> 0 And TicksToDouble(f) > 0 Then
        TimingSource = "QueryPerformanceCounter (" & Format$(TicksToDouble(f), "0") & " Hz)"
    Else
        TimingSource = "VB Timer (millisecond resolution)"
    End If
End Function

'------------------------------------------------------------------------------
' Message pump (no DoEvents inside event loggers)
'------------------------------------------------------------------------------
Public Sub Pump()
    DoEvents
End Sub

'------------------------------------------------------------------------------
' Text helpers
'------------------------------------------------------------------------------
Public Function FirstDiff(ByVal a As String, ByVal b As String) As String
    Dim la() As String
    Dim lb() As String
    Dim n As Long
    Dim i As Long
    la = Split(a, vbCrLf)
    lb = Split(b, vbCrLf)
    n = UBound(la)
    If UBound(lb) > n Then n = UBound(lb)
    For i = 0 To n
        If i > UBound(la) Then
            FirstDiff = "W2 ends at line " & (i + 1) & "; OCX: " & lb(i)
            Exit Function
        End If
        If i > UBound(lb) Then
            FirstDiff = "OCX ends at line " & (i + 1) & "; W2: " & la(i)
            Exit Function
        End If
        If la(i) <> lb(i) Then
            FirstDiff = "Line " & (i + 1) & vbCrLf & _
                        "W2 : " & la(i) & vbCrLf & _
                        "OCX: " & lb(i)
            Exit Function
        End If
    Next
    FirstDiff = ""
End Function

Public Function EscapeHtml(ByVal s As String) As String
    s = Replace(s, "&", "&amp;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    s = Replace(s, """", "&quot;")
    EscapeHtml = s
End Function

Public Function BytesToText(ByRef b() As Byte) As String
    On Error Resume Next
    BytesToText = StrConv(b, vbUnicode)
    If Err.Number <> 0 Then BytesToText = ""
    Err.Clear
End Function

' Cross-endpoint interleaving depends on message-pump timing, not on semantics,
' so traces are compared per endpoint: event order within an endpoint is kept,
' and endpoints are emitted in a canonical (sorted) order.
Public Function NormalizeTrace(ByVal s As String) As String
    Dim lines() As String
    Dim keys() As String
    Dim nKeys As Long
    Dim i As Long
    Dim k As Long
    Dim key As String
    Dim p As Long
    Dim out As String
    Dim lastLine As String
    Dim canon As String

    If Len(s) = 0 Then Exit Function
    lines = Split(s, vbCrLf)
    ReDim keys(0 To UBound(lines))
    For i = 0 To UBound(lines)
        If Len(lines(i)) > 0 Then
            key = TraceKey(lines(i))
            If Not HasKey(keys, nKeys, key) Then
                keys(nKeys) = key
                nKeys = nKeys + 1
            End If
        End If
    Next
    SortKeys keys, nKeys
    For k = 0 To nKeys - 1
        lastLine = ""
        For i = 0 To UBound(lines)
            If Len(lines(i)) > 0 Then
                If TraceKey(lines(i)) = keys(k) Then
                    canon = CanonicalLine(lines(i))
                    If canon <> lastLine Then
                        out = out & canon & vbCrLf
                        lastLine = canon
                    End If
                End If
            End If
        Next
    Next
    NormalizeTrace = out
End Function

' DataArrival's bytesTotal/BytesReceived depend on how much data happened to be
' buffered when the event fired, and the same buffered data may be reported in
' one or several events. Payload totals are asserted separately, so the trace
' keeps the event and drops the volatile numbers and repeats.
Private Function CanonicalLine(ByVal line As String) As String
    Dim p As Long
    p = InStr(1, line, "|DataArrival|", vbTextCompare)
    If p > 0 Then
        CanonicalLine = Left$(line, p - 1) & "|DataArrival"
    Else
        CanonicalLine = line
    End If
End Function

Private Function TraceKey(ByVal line As String) As String
    Dim p As Long
    p = InStr(1, line, "|")
    If p > 1 Then
        TraceKey = Left$(line, p - 1)
    Else
        TraceKey = line
    End If
End Function

Private Function HasKey(ByRef keys() As String, ByVal n As Long, ByVal key As String) As Boolean
    Dim i As Long
    For i = 0 To n - 1
        If keys(i) = key Then HasKey = True: Exit Function
    Next
End Function

Private Sub SortKeys(ByRef keys() As String, ByVal n As Long)
    Dim i As Long
    Dim j As Long
    Dim t As String
    For i = 1 To n - 1
        t = keys(i)
        j = i - 1
        Do While j >= 0
            If StrComp(keys(j), t, vbTextCompare) <= 0 Then Exit Do
            keys(j + 1) = keys(j)
            j = j - 1
        Loop
        keys(j + 1) = t
    Next
End Sub

Public Function LargeText(ByVal s As String) As String
    If Len(s) > 8000 Then
        LargeText = Left$(s, 8000) & vbCrLf & "... [truncated]"
    Else
        LargeText = s
    End If
End Function

Public Function StatusColor(ByVal status As String) As String
    Select Case status
    Case "PASS": StatusColor = "#1b7f3a"
    Case "DIFF": StatusColor = "#b06a00"
    Case "EXPECTED": StatusColor = "#3355aa"
    Case "W2_ONLY": StatusColor = "#666666"
    Case "SKIP": StatusColor = "#666666"
    Case Else: StatusColor = "#b00020"
    End Select
End Function

'------------------------------------------------------------------------------
' File output
'------------------------------------------------------------------------------
Public Sub WriteTextFile(ByVal path As String, ByVal text As String)
    Dim f As Integer
    On Error GoTo EH
    f = FreeFile
    Open path For Output As #f
    Print #f, text;
    Close #f
    Exit Sub
EH:
    On Error Resume Next
    If f <> 0 Then Close #f
End Sub

Public Function TimestampName() As String
    TimestampName = Format$(Now, "yyyymmdd_hhnnss")
End Function

'==============================================================================
' Reliable-payload helpers
'==============================================================================

' Deterministic, position-dependent pattern (period 65536): cheap to build and
' it exposes any offset, duplication or reordering of a stream.
Public Sub FillPattern(ByRef b() As Byte, ByVal n As Long, ByVal kind As Long)
    Dim i As Long
    Select Case kind
    Case PAT_ZERO
        For i = 0 To n - 1: b(i) = 0: Next
    Case PAT_FF
        For i = 0 To n - 1: b(i) = 255: Next
    Case PAT_COUNT
        For i = 0 To n - 1: b(i) = (i And 255): Next
    Case PAT_NULLMIX
        For i = 0 To n - 1
            Select Case (i Mod 4)
            Case 0: b(i) = 0
            Case 1: b(i) = 13
            Case 2: b(i) = 10
            Case 3: b(i) = 255
            End Select
        Next
    Case Else
        For i = 0 To n - 1
            b(i) = ((i * 31& + 7&) And &HFF) Xor ((i \ 256&) And &HFF)
        Next
    End Select
End Sub

' Native-speed byte comparison (StrConv + StrComp), no per-byte loop.
Public Function BuffersEqual(ByRef a() As Byte, ByRef b() As Byte) As Boolean
    Dim sa As String
    Dim sb As String
    Dim la As Long
    Dim lb As Long
    On Error GoTo Failed
    la = UBound(a) - LBound(a) + 1
    lb = UBound(b) - LBound(b) + 1
    If la <> lb Then Exit Function
    sa = StrConv(a, vbUnicode)
    sb = StrConv(b, vbUnicode)
    BuffersEqual = (StrComp(sa, sb, vbBinaryCompare) = 0)
    Exit Function
Failed:
    BuffersEqual = False
End Function

Public Function FirstMismatch(ByRef a() As Byte, ByRef b() As Byte) As Long
    Dim la As Long
    Dim lb As Long
    Dim n As Long
    Dim i As Long
    la = UBound(a) - LBound(a) + 1
    lb = UBound(b) - LBound(b) + 1
    n = la
    If lb < n Then n = lb
    For i = 0 To n - 1
        If a(LBound(a) + i) <> b(LBound(b) + i) Then
            FirstMismatch = i
            Exit Function
        End If
    Next
    If la <> lb Then FirstMismatch = n Else FirstMismatch = -1
End Function

'==============================================================================
' Percentiles
'==============================================================================
Public Sub Percentiles(ByRef samples() As Long, ByVal n As Long, ByRef p50 As Double, ByRef p95 As Double, ByRef p99 As Double, ByRef worst As Double)
    Dim i As Long
    Dim j As Long
    Dim t As Long
    If n <= 0 Then Exit Sub
    ' insertion sort: sample counts here are in the hundreds/thousands
    For i = 1 To n - 1
        t = samples(i)
        j = i - 1
        Do While j >= 0
            If samples(j) <= t Then Exit Do
            samples(j + 1) = samples(j)
            j = j - 1
        Loop
        samples(j + 1) = t
    Next
    p50 = samples(Int(n * 0.5))
    p95 = samples(Int(n * 0.95))
    p99 = samples(Int(n * 0.99))
    worst = samples(n - 1)
End Sub

'==============================================================================
' Resource counters (process pseudo-handle -1)
'==============================================================================
Public Function ProcHandleCount() As Long
    Dim n As Long
    If GetProcessHandleCount(-1&, n) <> 0 Then ProcHandleCount = n
End Function

Public Function ProcGdiCount() As Long
    ProcGdiCount = GetGuiResources(-1&, W2_GR_GDIOBJECTS)
End Function

Public Function ProcUserCount() As Long
    ProcUserCount = GetGuiResources(-1&, W2_GR_USEROBJECTS)
End Function

Public Function ResourceLine() As String
    ResourceLine = "handles=" & ProcHandleCount() & " gdi=" & ProcGdiCount() & " user=" & ProcUserCount()
End Function

' modWinsock2 keeps read-only counters for exactly this purpose.
Public Function DiagnosticsLine() As String
    On Error Resume Next
    DiagnosticsLine = "diag{" & W2_Diagnostics() & "}"
    If Err.Number <> 0 Then
        DiagnosticsLine = "diag{unavailable}"
        Err.Clear
    End If
End Function

' Forces a reset (RST) instead of a graceful FIN on the next close.
Public Sub SetLingerZero(ByVal s As Long)
    Dim v As Long
    If s = -1 Or s = 0 Then Exit Sub
    v = 1&   ' l_onoff = 1, l_linger = 0
    setsockopt s, W2_SOL_SOCKET, W2_SO_LINGER, v, 4&
End Sub

Public Sub SetSockBuf(ByVal s As Long, ByVal opt As Long, ByVal bytes As Long)
    If s = -1 Or s = 0 Then Exit Sub
    setsockopt s, W2_SOL_SOCKET, opt, bytes, 4&
End Sub

Public Sub SetRecvBuf(ByVal s As Long, ByVal bytes As Long)
    SetSockBuf s, W2_SO_RCVBUF, bytes
End Sub

Public Sub SetSendBuf(ByVal s As Long, ByVal bytes As Long)
    SetSockBuf s, W2_SO_SNDBUF, bytes
End Sub

Public Sub ShutdownSend(ByVal s As Long)
    If s = -1 Or s = 0 Then Exit Sub
    w2_shutdown s, 1&
End Sub

Public Function FileExists(ByVal path As String) As Boolean
    On Error Resume Next
    FileExists = (Len(Dir$(path)) > 0)
End Function

Public Function ReadTextFile(ByVal path As String) As String
    Dim f As Integer
    Dim s As String
    On Error GoTo Failed
    f = FreeFile
    Open path For Binary As #f
    s = Space$(LOF(f))
    Get #f, , s
    Close #f
    ReadTextFile = s
    Exit Function
Failed:
    On Error Resume Next
    If f <> 0 Then Close #f
End Function

Public Sub AppendTextFile(ByVal path As String, ByVal text As String)
    Dim f As Integer
    On Error Resume Next
    f = FreeFile
    Open path For Append As #f
    Print #f, text
    Close #f
End Sub
