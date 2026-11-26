Attribute VB_Name = "modWinsock2"
Option Explicit

'==============================================================================
'        modWinsock2 - Pure VB6 replacement for MSWinsockLib.Winsock
'==============================================================================
'
' O    O    O
'  \__/ \__/
'  /=||=||=\   oouuuunnnnnnnnmmmmmmmmmmmmmm\
' // ||_||                WinsockVB6        \
' \\ /\ #\     oouuuunnnnnnnnmmmmmmmmmmmmmmmm\
' /=(  \  )==>       by EBArtSoft@ + AI       \
'//  \O_\/         Copyright © 2006 - 2026     \
'\\  || ||        https://www.ebartsoft.com     \
' \==||=||==/  oouuuunnnnnnnnmmmmmmmmmmmmmmmmmmmm\
' ===========
'==== E.B ====
'
'=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!
'
'                            AI ASSISTANCE
'
'                  OpenCode - DeepSeek V4.1 Flash 1017
'
'=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!
'
' MIT License
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in
' all copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
' THE SOFTWARE.
'

'================================================
' Module Usage
'
' Add "modWinsock2.bas" to your project
'
' Then add reference to "Win32API.tlb"
'
' then create a new socket object using
'   Set sock = New_WinsockVB6()
'
' If you want to use WinsockVB6.dll
' add the line below in your code
'   Public Declare Function New_WinsockVB6 _
'       Lib "WinsockVB6.dll" () As Winsock
' and remove "modWinsock2.bas"
'
'================================================

'------------------------------------------------------------------------------
' Const
'------------------------------------------------------------------------------
Private Const W2_VERSION                As String = "1.0"

Private Const W2_AF_INET                As Long = 2
Private Const W2_SOCK_STREAM            As Long = 1
Private Const W2_SOCK_DGRAM             As Long = 2
Private Const W2_IPPROTO_TCP            As Long = 6
Private Const W2_IPPROTO_UDP            As Long = 17
Private Const W2_INVALID_SOCKET         As Long = -1
Private Const W2_SOCKET_ERROR           As Long = -1
Private Const W2_INADDR_NONE            As Long = -1
Private Const W2_INADDR_ANY             As Long = 0

Private Const W2_FD_READ                As Long = &H1&
Private Const W2_FD_WRITE               As Long = &H2&
Private Const W2_FD_OOB                 As Long = &H4&
Private Const W2_FD_ACCEPT              As Long = &H8&
Private Const W2_FD_CONNECT             As Long = &H10&
Private Const W2_FD_CLOSE               As Long = &H20&

Private Const W2_SD_SEND                As Long = 1

Private Const W2_GMEM_FIXED             As Long = 0
Private Const W2_MAXGETHOSTSTRUCT       As Long = 1024
Private Const W2_HEAP_ZERO_MEMORY       As Long = &H8
Private Const W2_GWL_WNDPROC            As Long = -4
Private Const W2_HWND_MESSAGE           As Long = -3
Private Const W2_PM_REMOVE              As Long = &H1&

Private Const W2_WSAEWOULDBLOCK         As Long = 10035
Private Const W2_WSAEINPROGRESS         As Long = 10036
Private Const W2_WSAEALREADY            As Long = 10037
Private Const W2_WSAENOTSOCK            As Long = 10038
Private Const W2_WSAEDESTADDRREQ        As Long = 10039
Private Const W2_WSAEMSGSIZE            As Long = 10040
Private Const W2_WSAEADDRINUSE          As Long = 10048
Private Const W2_WSAEADDRNOTAVAIL       As Long = 10049
Private Const W2_WSAENETDOWN            As Long = 10050
Private Const W2_WSAENETUNREACH         As Long = 10051
Private Const W2_WSAENETRESET           As Long = 10052
Private Const W2_WSAECONNABORTED        As Long = 10053
Private Const W2_WSAECONNRESET          As Long = 10054
Private Const W2_WSAENOBUFS             As Long = 10055
Private Const W2_WSAEISCONN             As Long = 10056
Private Const W2_WSAENOTCONN            As Long = 10057
Private Const W2_WSAESHUTDOWN           As Long = 10058
Private Const W2_WSAETIMEDOUT           As Long = 10060
Private Const W2_WSAECONNREFUSED        As Long = 10061
Private Const W2_WSAEHOSTUNREACH        As Long = 10065
Private Const W2_WSANOTINITIALISED      As Long = 10093
Private Const W2_WSAHOST_NOT_FOUND      As Long = 11001
Private Const W2_WSATRY_AGAIN           As Long = 11002
Private Const W2_WSANO_RECOVERY         As Long = 11003
Private Const W2_WSANO_DATA             As Long = 11004

Private Const W2_E_NOTIMPL              As Long = &H80004001
Private Const W2_E_NOINTERFACE          As Long = &H80004002
Private Const W2_E_POINTER              As Long = &H80004003
Private Const W2_E_FAIL                 As Long = &H80004005
Private Const W2_DISP_E_BADINDEX        As Long = &H8002000B
Private Const W2_DISP_E_UNKNOWNNAME     As Long = &H80020006
Private Const W2_DISP_E_MEMBERNOTFOUND  As Long = &H80020003
Private Const W2_DISP_E_EXCEPTION       As Long = &H80020009
Private Const W2_ARG_NOT_FOUND          As Long = &H7FFFFFFF
Private Const W2_CONNECT_E_NOCONNECTION As Long = &H80040202

Private Const W2_EV_CONNECT             As Long = 1
Private Const W2_EV_DATA                As Long = 2
Private Const W2_EV_CONNREQ             As Long = 3
Private Const W2_EV_SENDPROGRESS        As Long = 4
Private Const W2_EV_SENDCOMPLETE        As Long = 5
Private Const W2_EV_CLOSE               As Long = 6

'------------------------------------------------------------------------------
' Types
'------------------------------------------------------------------------------
Private Type W2_MSG
    hWnd    As Long
    message As Long
    wParam  As Long
    lParam  As Long
    time    As Long
    ptX     As Long
    ptY     As Long
End Type

Private Type W2_IFACE
    VTable      As Long
    StateIndex  As Long
End Type

Private Type W2_DISPPARAMS
    rgvarg              As Long
    rgdispidNamedArgs   As Long
    cArgs               As Long
    cNamedArgs          As Long
End Type

Private Type W2_EXCEPINFO
    wCode               As Integer
    wReserved           As Integer
    bstrSource          As Long
    bstrDescription     As Long
    bstrHelpFile        As Long
    dwHelpContext       As Long
    pvReserved          As Long
    scode               As Long
    pfnDeferredFillIn   As Long
End Type

Private Type W2_STATE
    InUse               As Boolean
    NextFree            As Long
    RefCount            As Long
    HeaderPtr           As Long

    Protocol            As Long
    State               As Long
    SocketHandle        As Long
    SocketRegistered    As Boolean
    Bound               As Boolean
    UdpConnected        As Boolean

    LocalPort           As Long
    LocalIP             As String
    LocalHostName       As String
    RemotePort          As Long
    RemoteHost          As String
    RemoteHostIP        As String
    BytesReceived       As Long

    ReadData()          As Byte
    ReadCap             As Long
    ReadSize            As Long
    ReadOffset          As Long

    WriteData()         As Byte
    WriteCap            As Long
    WriteSize           As Long
    WriteOffset         As Long
    SendTotal           As Long
    SendPending         As Boolean

    UdpData             As Object      ' Collection of Byte arrays
    Sinks               As Object      ' Collection: cookie -> event sink
    NextCookie          As Long

    DnsReq              As Long
    DnsHandle           As Long
    DnsPending          As Boolean

    CloseRequested      As Boolean
    Pending             As Object      ' Collection of pending requestID strings
End Type

'------------------------------------------------------------------------------
' Module state
'------------------------------------------------------------------------------
Private gStates()           As W2_STATE
Private gAllocated          As Long
Private gHighWater          As Long
Private gFreeHead           As Long
Private gObjCount           As Long
Private gV1()               As Long
Private gV2()               As Long
Private gV3()               As Long
Private gV4()               As Long
Private gReady              As Boolean

' Shared scratch buffer for UDP receives: a 64 KB ReDim per datagram shows up
' clearly in the UDP round-trip benchmark.
Private gUdpBuf()           As Byte
Private gUdpCap             As Long

Private gIID_IUnknown               As CLSID
Private gIID_IDispatch              As CLSID
Private gIID_IMSWinsockControl      As CLSID
Private gIID_IConnectionPoint       As CLSID
Private gIID_IConnectionPointCont   As CLSID
Private gIID_IEnumConnections       As CLSID
Private gIID_Events                 As CLSID

Private gWin                As Long
Private gOldProc            As Long
Private gMsgSock            As Long
Private gMsgDns             As Long

Private gSockets            As Collection
Private gDnsMap             As Collection
Private gPendingSocks       As Collection
Private gPendingOwners      As Collection
Private gNextRequestID      As Long

Private gWSAStarted         As Boolean
Private gLocalHostName      As String
Private gLocalIP            As String
Private gInErrorBox         As Boolean

'==============================================================================
' Public factory
'==============================================================================

'================================================
' New_Winsock2
'================================================
Public Function New_Winsock2() As Winsock
    Dim idx As Long
    Dim p   As Long

    W2_EnsureModule

    idx = W2_AllocState()
    If idx < 0 Then
        Err.Raise sckOutOfMemory, "New_Winsock2", "Unable to allocate Winsock object."
    End If

    p = W2_AllocHeaders(idx)
    If p = 0 Then
        W2_FreeState idx
        Err.Raise sckOutOfMemory, "New_Winsock2", "Unable to allocate COM interface headers."
    End If

    gStates(idx).HeaderPtr = p
    MemoryCopy New_Winsock2, p, 4
End Function

'==============================================================================
' Module initialization / shutdown
'==============================================================================

'================================================
' W2_EnsureModule
'================================================
Private Sub W2_EnsureModule()
    Dim wsa As WSADATA

    If Not gReady Then
        W2_InitIIDs
        W2_BuildVtables
        Set gSockets = New Collection
        Set gDnsMap = New Collection
        Set gPendingSocks = New Collection
        Set gPendingOwners = New Collection
        gNextRequestID = 0
        gFreeHead = -1
        gAllocated = 0
        gHighWater = 0
        gObjCount = 0
        gReady = True
    End If

    If Not gWSAStarted Then
        If WSAStartup(&H101, wsa) <> 0 Then
            Err.Raise sckNotInitialized, "New_Winsock2", "Unable to initialize Winsock."
        End If
        gWSAStarted = True
    End If

    If gLocalHostName = "" And gLocalIP = "" Then
        W2_LoadLocalInfo
    End If

    If gWin = 0 Then
        W2_CreateMessageWindow
    End If
End Sub

'================================================
' W2_ShutdownModule
'================================================
Private Sub W2_ShutdownModule()
    If gWin <> 0 Then
        W2_DestroyMessageWindow
    End If

    If gWSAStarted Then
        WSACleanup
        gWSAStarted = False
    End If

    Set gSockets = Nothing
    Set gDnsMap = Nothing
    Set gPendingSocks = Nothing
    Set gPendingOwners = Nothing

    Erase gStates
    Erase gUdpBuf
    gUdpCap = 0
    gAllocated = 0
    gHighWater = 0
    gFreeHead = -1
    gReady = False
End Sub

'================================================
' W2_InitIIDs
'================================================
Private Sub W2_InitIIDs()
    CLSIDFromString "{00000000-0000-0000-C000-000000000046}", gIID_IUnknown
    CLSIDFromString "{00020400-0000-0000-C000-000000000046}", gIID_IDispatch
    CLSIDFromString "{248DD892-BB45-11CF-9ABC-0080C7E7B78D}", gIID_IMSWinsockControl
    CLSIDFromString "{B196B284-BAB4-101A-B69C-00AA00341D07}", gIID_IConnectionPointCont
    CLSIDFromString "{B196B286-BAB4-101A-B69C-00AA00341D07}", gIID_IConnectionPoint
    CLSIDFromString "{B196B287-BAB4-101A-B69C-00AA00341D07}", gIID_IEnumConnections
    CLSIDFromString "{248DD893-BB45-11CF-9ABC-0080C7E7B78D}", gIID_Events
End Sub

'================================================
' W2_BuildVtables
'================================================
Private Sub W2_BuildVtables()
    ' IMSWinsockControl: IUnknown (3) + IDispatch (4) + interface (25) = 32 slots
    ReDim gV1(0 To 31)
    gV1(0) = W2_AdrsOf(AddressOf W2_QueryInterface)
    gV1(1) = W2_AdrsOf(AddressOf W2_AddRef)
    gV1(2) = W2_AdrsOf(AddressOf W2_Release)
    gV1(3) = W2_AdrsOf(AddressOf W2_GetTypeInfoCount)
    gV1(4) = W2_AdrsOf(AddressOf W2_GetTypeInfo)
    gV1(5) = W2_AdrsOf(AddressOf W2_GetIDsOfNames)
    gV1(6) = W2_AdrsOf(AddressOf W2_Invoke)
    gV1(7) = W2_AdrsOf(AddressOf W2_AboutBox)
    gV1(8) = W2_AdrsOf(AddressOf W2_GetProtocol)
    gV1(9) = W2_AdrsOf(AddressOf W2_LetProtocol)
    gV1(10) = W2_AdrsOf(AddressOf W2_GetRemoteHostIP)
    gV1(11) = W2_AdrsOf(AddressOf W2_GetLocalHostName)
    gV1(12) = W2_AdrsOf(AddressOf W2_GetLocalIP)
    gV1(13) = W2_AdrsOf(AddressOf W2_GetSocketHandle)
    gV1(14) = W2_AdrsOf(AddressOf W2_Get_RemoteHost)
    gV1(15) = W2_AdrsOf(AddressOf W2_Let_RemoteHost)
    gV1(16) = W2_AdrsOf(AddressOf W2_GetRemotePort)
    gV1(17) = W2_AdrsOf(AddressOf W2_LetRemotePort)
    gV1(18) = W2_AdrsOf(AddressOf W2_GetLocalPort)
    gV1(19) = W2_AdrsOf(AddressOf W2_LetLocalPort)
    gV1(20) = W2_AdrsOf(AddressOf W2_GetState)
    gV1(21) = W2_AdrsOf(AddressOf W2_GetBytesReceived)
    gV1(22) = W2_AdrsOf(AddressOf W2_Connect)
    gV1(23) = W2_AdrsOf(AddressOf W2_Listen)
    gV1(24) = W2_AdrsOf(AddressOf W2_Accept)
    gV1(25) = W2_AdrsOf(AddressOf W2_SendData)
    gV1(26) = W2_AdrsOf(AddressOf W2_GetData)
    gV1(27) = W2_AdrsOf(AddressOf W2_PeekData)
    gV1(28) = W2_AdrsOf(AddressOf W2_Close)
    gV1(29) = W2_AdrsOf(AddressOf W2_Bind)
    gV1(30) = W2_AdrsOf(AddressOf W2_GetRemoteHost)
    gV1(31) = W2_AdrsOf(AddressOf W2_LetRemoteHost)

    ' IConnectionPointContainer: IUnknown + 2
    ReDim gV2(0 To 4)
    gV2(0) = W2_AdrsOf(AddressOf W2_QueryInterface)
    gV2(1) = W2_AdrsOf(AddressOf W2_AddRef)
    gV2(2) = W2_AdrsOf(AddressOf W2_Release)
    gV2(3) = W2_AdrsOf(AddressOf W2_EnumConnectionPoints)
    gV2(4) = W2_AdrsOf(AddressOf W2_FindConnectionPoint)

    ' IConnectionPoint: IUnknown + 5
    ReDim gV3(0 To 7)
    gV3(0) = W2_AdrsOf(AddressOf W2_QueryInterface)
    gV3(1) = W2_AdrsOf(AddressOf W2_AddRef)
    gV3(2) = W2_AdrsOf(AddressOf W2_Release)
    gV3(3) = W2_AdrsOf(AddressOf W2_GetConnectionInterface)
    gV3(4) = W2_AdrsOf(AddressOf W2_GetConnectionPointContainer)
    gV3(5) = W2_AdrsOf(AddressOf W2_Advise)
    gV3(6) = W2_AdrsOf(AddressOf W2_Unadvise)
    gV3(7) = W2_AdrsOf(AddressOf W2_EnumConnections)

    ' IEnumConnections: IUnknown + 4 (not exposed in v1)
    ReDim gV4(0 To 6)
    gV4(0) = W2_AdrsOf(AddressOf W2_QueryInterface)
    gV4(1) = W2_AdrsOf(AddressOf W2_AddRef)
    gV4(2) = W2_AdrsOf(AddressOf W2_Release)
    gV4(3) = W2_AdrsOf(AddressOf W2_EnumNext)
    gV4(4) = W2_AdrsOf(AddressOf W2_EnumSkip)
    gV4(5) = W2_AdrsOf(AddressOf W2_EnumReset)
    gV4(6) = W2_AdrsOf(AddressOf W2_EnumClone)
End Sub

'================================================
' W2_AdrsOf
'================================================
Private Function W2_AdrsOf(ByVal p As Long) As Long
    W2_AdrsOf = p
End Function

'==============================================================================
' State allocator (dynamic pool + embedded free list)
'==============================================================================

'================================================
' W2_AllocState
'================================================
Private Function W2_AllocState() As Long
    Dim idx As Long

    On Error GoTo OOM

    If gFreeHead >= 0 Then
        idx = gFreeHead
        gFreeHead = gStates(idx).NextFree
    Else
        idx = gHighWater
        If idx >= gAllocated Then
            If gAllocated = 0 Then
                ReDim gStates(0 To 31)
                gAllocated = 32
            Else
                ReDim Preserve gStates(0 To gAllocated + 31)
                gAllocated = gAllocated + 32
            End If
        End If
        gHighWater = idx + 1
    End If

    W2_ClearState idx
    gStates(idx).InUse = True
    gStates(idx).RefCount = 1
    gObjCount = gObjCount + 1
    W2_AllocState = idx
    Exit Function
OOM:
    W2_AllocState = -1
End Function

'================================================
' W2_ClearState
'================================================
Private Sub W2_ClearState(ByVal idx As Long)
    With gStates(idx)
        .InUse = False
        .NextFree = -1
        .RefCount = 0
        .HeaderPtr = 0

        .Protocol = sckTCPProtocol
        .State = sckClosed
        .SocketHandle = W2_INVALID_SOCKET
        .SocketRegistered = False
        .Bound = False
        .UdpConnected = False

        .LocalPort = 0
        .LocalIP = ""
        .LocalHostName = gLocalHostName
        .RemotePort = 0
        .RemoteHost = ""
        .RemoteHostIP = ""
        .BytesReceived = 0

        Erase .ReadData
        .ReadCap = 0
        .ReadSize = 0
        .ReadOffset = 0

        Erase .WriteData
        .WriteCap = 0
        .WriteSize = 0
        .WriteOffset = 0
        .SendTotal = 0
        .SendPending = False

        Set .UdpData = Nothing
        Set .UdpData = New Collection
        Set .Sinks = Nothing
        Set .Sinks = New Collection
        .NextCookie = 0

        .DnsReq = 0
        .DnsHandle = 0
        .DnsPending = False
        .CloseRequested = False

        Set .Pending = Nothing
        Set .Pending = New Collection
    End With
End Sub

'================================================
' W2_FreeState
'================================================
Private Sub W2_FreeState(ByVal idx As Long)
    If idx < 0 Or idx >= gHighWater Then Exit Sub
    If Not gStates(idx).InUse Then Exit Sub

    ' Close connections accepted on this object that nobody ever took.
    W2_ClosePending idx

    With gStates(idx)
        Set .UdpData = Nothing
        Set .Sinks = Nothing
        Set .Pending = Nothing
        .RemoteHost = ""
        .RemoteHostIP = ""
        .LocalIP = ""
        .LocalHostName = ""
        Erase .ReadData
        Erase .WriteData
        .ReadCap = 0
        .ReadSize = 0
        .ReadOffset = 0
        .WriteCap = 0
        .WriteSize = 0
        .WriteOffset = 0
        .InUse = False
        .NextFree = gFreeHead
    End With

    gFreeHead = idx
    gObjCount = gObjCount - 1
    If gObjCount < 0 Then gObjCount = 0
    ' See W2_DestroyObject: the module state outlives the individual objects.
End Sub

'==============================================================================
' COM headers
'==============================================================================

'================================================
' W2_AllocHeaders
'================================================
Private Function W2_AllocHeaders(ByVal idx As Long) As Long
    Dim p As Long

    p = HeapAlloc(GetProcessHeap(), W2_HEAP_ZERO_MEMORY, 32&)
    If p = 0 Then Exit Function

    W2_Poke4 p + 0, VarPtr(gV1(0))
    W2_Poke4 p + 4, idx
    W2_Poke4 p + 8, VarPtr(gV2(0))
    W2_Poke4 p + 12, idx
    W2_Poke4 p + 16, VarPtr(gV3(0))
    W2_Poke4 p + 20, idx
    W2_Poke4 p + 24, VarPtr(gV4(0))
    W2_Poke4 p + 28, idx

    W2_AllocHeaders = p
End Function

'================================================
' W2_Poke4
'================================================
Private Sub W2_Poke4(ByVal p As Long, ByVal v As Long)
    MemoryCopy ByVal p, v, 4
End Sub

'==============================================================================
' Small utilities
'==============================================================================

'================================================
' W2_ValidState
'================================================
Private Function W2_ValidState(ByVal idx As Long) As Boolean
    If idx < 0 Then Exit Function
    If idx >= gHighWater Then Exit Function
    If gHighWater = 0 Then Exit Function
    W2_ValidState = gStates(idx).InUse
End Function

'================================================
' W2_IsEqualIID
'================================================
Private Function W2_IsEqualIID(ByRef a As CLSID, ByRef b As CLSID) As Boolean
    If a.Data1 <> b.Data1 Then Exit Function
    If a.Data2 <> b.Data2 Then Exit Function
    If a.Data3 <> b.Data3 Then Exit Function
    If a.Data4(0) <> b.Data4(0) Then Exit Function
    If a.Data4(1) <> b.Data4(1) Then Exit Function
    If a.Data4(2) <> b.Data4(2) Then Exit Function
    If a.Data4(3) <> b.Data4(3) Then Exit Function
    If a.Data4(4) <> b.Data4(4) Then Exit Function
    If a.Data4(5) <> b.Data4(5) Then Exit Function
    If a.Data4(6) <> b.Data4(6) Then Exit Function
    If a.Data4(7) <> b.Data4(7) Then Exit Function
    W2_IsEqualIID = True
End Function

'================================================
' W2_ToUInt
'================================================
Private Function W2_ToUInt(ByVal v As Integer) As Long
    MemoryCopy W2_ToUInt, v, 2
End Function

'================================================
' W2_ToInt
'================================================
Private Function W2_ToInt(ByVal v As Long) As Integer
    MemoryCopy W2_ToInt, v, 2
End Function

'================================================
' W2_LongToIP
'================================================
Private Function W2_LongToIP(ByVal v As Long) As String
    Dim b(0 To 3) As Byte
    MemoryCopy b(0), v, 4
    W2_LongToIP = b(0) & "." & b(1) & "." & b(2) & "." & b(3)
End Function

'================================================
' W2_StringFromPointer
'================================================
Private Function W2_StringFromPointer(ByVal p As Long) As String
    Dim s() As Byte
    Dim n   As Long
    If p = 0 Then Exit Function
    n = lstrlen(ByVal p)
    If n <= 0 Then Exit Function
    ReDim s(0 To n - 1)
    MemoryCopy s(0), ByVal p, n
    W2_StringFromPointer = StrConv(s, vbUnicode)
End Function

'================================================
' W2_IsIPLiteral
'================================================
Private Function W2_IsIPLiteral(ByVal s As String) As Boolean
    If s = "255.255.255.255" Then
        W2_IsIPLiteral = True
        Exit Function
    End If
    W2_IsIPLiteral = (inet_addr(s) <> W2_INADDR_NONE)
End Function

'================================================
' W2_GrowCap
'================================================
Private Function W2_GrowCap(ByVal need As Long) As Long
    Dim cap As Long
    cap = 256
    Do While cap < need
        If cap > &H3FFFFFFF Then
            cap = need
            Exit Do
        End If
        cap = cap * 2
    Loop
    W2_GrowCap = cap
End Function

' Length of a Byte() that may not be dimensioned. MSWINSCK treats an empty
' array as a no-op; LBound on an empty array raises error 9 instead.
'================================================
' W2_ArrayLen
'================================================
Private Function W2_ArrayLen(ByRef b() As Byte) As Long
    Dim lb As Long
    Dim ub As Long
    On Error GoTo LenFailed
    lb = LBound(b)
    ub = UBound(b)
    W2_ArrayLen = ub - lb + 1
    Exit Function
LenFailed:
    W2_ArrayLen = 0
End Function

'================================================
' W2_TypeSize
'================================================
Private Function W2_TypeSize(ByVal vt As Long) As Long
    Select Case vt
    Case vbByte:     W2_TypeSize = 1
    Case vbBoolean:  W2_TypeSize = 2
    Case vbInteger:  W2_TypeSize = 2
    Case vbLong:     W2_TypeSize = 4
    Case vbSingle:   W2_TypeSize = 4
    Case vbDouble:   W2_TypeSize = 8
    Case vbCurrency: W2_TypeSize = 8
    Case vbDate:     W2_TypeSize = 8
    Case Else:       W2_TypeSize = 0
    End Select
End Function

'==============================================================================
' IUnknown
'==============================================================================

'================================================
' W2_QueryInterface
'================================================
Private Function W2_QueryInterface(ByRef pThis As W2_IFACE, ByRef riid As CLSID, ByRef ppv As Long) As Long
    Dim idx As Long

    idx = pThis.StateIndex
    If Not W2_ValidState(idx) Then
        W2_QueryInterface = W2_E_FAIL
        Exit Function
    End If

    If W2_IsEqualIID(riid, gIID_IUnknown) Or W2_IsEqualIID(riid, gIID_IDispatch) Or W2_IsEqualIID(riid, gIID_IMSWinsockControl) Then
        ppv = gStates(idx).HeaderPtr
        W2_AddRefIndex idx
        W2_QueryInterface = 0

    ElseIf W2_IsEqualIID(riid, gIID_IConnectionPointCont) Then
        ppv = gStates(idx).HeaderPtr + 8
        W2_AddRefIndex idx
        W2_QueryInterface = 0

    ElseIf W2_IsEqualIID(riid, gIID_IConnectionPoint) Then
        ppv = gStates(idx).HeaderPtr + 16
        W2_AddRefIndex idx
        W2_QueryInterface = 0

    Else
        ppv = 0
        W2_QueryInterface = W2_E_NOINTERFACE
    End If
End Function

'================================================
' W2_AddRef
'================================================
Private Function W2_AddRef(ByRef pThis As W2_IFACE) As Long
    Dim idx As Long
    idx = pThis.StateIndex
    If Not W2_ValidState(idx) Then Exit Function
    gStates(idx).RefCount = gStates(idx).RefCount + 1
    W2_AddRef = gStates(idx).RefCount
End Function

'================================================
' W2_Release
'================================================
Private Function W2_Release(ByRef pThis As W2_IFACE) As Long
    Dim idx As Long
    Dim rc  As Long

    idx = pThis.StateIndex
    If Not W2_ValidState(idx) Then Exit Function

    gStates(idx).RefCount = gStates(idx).RefCount - 1
    rc = gStates(idx).RefCount
    If rc <= 0 Then W2_DestroyObject idx
    W2_Release = rc
End Function

'================================================
' W2_AddRefIndex
'================================================
Private Sub W2_AddRefIndex(ByVal idx As Long)
    If W2_ValidState(idx) Then gStates(idx).RefCount = gStates(idx).RefCount + 1
End Sub

'================================================
' W2_ReleaseIndex
'================================================
Private Sub W2_ReleaseIndex(ByVal idx As Long)
    Dim rc As Long
    If Not W2_ValidState(idx) Then Exit Sub
    gStates(idx).RefCount = gStates(idx).RefCount - 1
    rc = gStates(idx).RefCount
    If rc <= 0 Then W2_DestroyObject idx
End Sub

'==============================================================================
' IDispatch
'==============================================================================

'================================================
' W2_GetTypeInfoCount
'================================================
Private Function W2_GetTypeInfoCount(ByRef pThis As W2_IFACE, ByRef ppctinfo As Long) As Long
    ppctinfo = 0
End Function

'================================================
' W2_GetTypeInfo
'================================================
Private Function W2_GetTypeInfo(ByRef pThis As W2_IFACE, ByVal itinfo As Long, ByVal lcid As Long, ByRef ppTInfo As ITypeInfo) As Long
    W2_GetTypeInfo = W2_DISP_E_BADINDEX
End Function

'==============================================================================
' IDispatch automation layer
' A VB6 form host reaches a control only through IDispatch::Invoke, using the
' dispids of MSWinsockLib (see TypeLib\MSWinsockLib.idl). The dispatch has to
' live in VB and not in the C shim: when a method raises, the VB runtime only
' hands the caller a proper HRESULT if the call was made from VB. Through a C
' frame the raise escapes as RPC_E_DISCONNECTED (0x80010108) instead.
'==============================================================================

'================================================
' W2_Peek2
'================================================
Private Function W2_Peek2(ByVal p As Long) As Integer
    If p = 0 Then Exit Function
    MemoryCopy W2_Peek2, ByVal p, 2
End Function

'================================================
' W2_Peek4
'================================================
Private Function W2_Peek4(ByVal p As Long) As Long
    If p = 0 Then Exit Function
    MemoryCopy W2_Peek4, ByVal p, 4
End Function

'================================================
' W2_ArgCount
'================================================
Private Function W2_ArgCount(ByVal pDispParams As Long) As Long
    Dim dp As W2_DISPPARAMS
    If pDispParams = 0 Then Exit Function
    MemoryCopy dp, ByVal pDispParams, LenB(dp)
    W2_ArgCount = dp.cArgs
End Function

' DISPPARAMS stores the arguments in reverse order: index 0 is the rightmost.
'================================================
' W2_ArgV
'================================================
Private Function W2_ArgV(ByVal pDispParams As Long, ByVal index As Long) As Variant
    Dim dp As W2_DISPPARAMS
    Dim v  As Variant
    If pDispParams = 0 Then Exit Function
    MemoryCopy dp, ByVal pDispParams, LenB(dp)
    If index >= dp.cArgs Then Exit Function
    MemoryCopy v, ByVal dp.rgvarg + (dp.cArgs - 1 - index) * 16, 16
    W2_ArgV = v
End Function

'================================================
' W2_ArgMissing
'================================================
Private Function W2_ArgMissing(ByVal pDispParams As Long, ByVal index As Long) As Boolean
    Dim dp    As W2_DISPPARAMS
    Dim v     As Variant
    Dim scode As Long
    If pDispParams = 0 Then W2_ArgMissing = True: Exit Function
    MemoryCopy dp, ByVal pDispParams, LenB(dp)
    If index >= dp.cArgs Then W2_ArgMissing = True: Exit Function
    MemoryCopy v, ByVal dp.rgvarg + (dp.cArgs - 1 - index) * 16, 16
    If VarType(v) = vbError Then
        scode = W2_Peek4(VarPtr(v) + 8)
        W2_ArgMissing = (scode = &H80020004)
    End If
End Function

'================================================
' W2_ArgLong
'================================================
Private Function W2_ArgLong(ByVal pDispParams As Long, ByVal index As Long) As Long
    Dim v As Variant
    v = W2_ArgV(pDispParams, index)
    If IsNumeric(v) Then W2_ArgLong = CLng(v)
End Function

'================================================
' W2_ArgStr
'================================================
Private Function W2_ArgStr(ByVal pDispParams As Long, ByVal index As Long) As String
    Dim v As Variant
    v = W2_ArgV(pDispParams, index)
    If VarType(v) = vbString Then W2_ArgStr = CStr(v)
End Function

' Moves the value into the caller's VARIANT and gives up ownership of it.
'================================================
' W2_PutResult
'================================================
Private Sub W2_PutResult(ByVal pVarResult As Long, ByRef v As Variant)
    Dim zero As Long
    If pVarResult = 0 Then Exit Sub
    MemoryCopy ByVal pVarResult, v, 16
    MemoryCopy v, zero, 2
End Sub

'================================================
' W2_ClearResult
'================================================
Private Sub W2_ClearResult(ByVal pVarResult As Long)
    Dim v As Variant
    If pVarResult = 0 Then Exit Sub
    MemoryCopy ByVal pVarResult, v, 16
End Sub

'================================================
' W2_NameFromPtr
'================================================
Private Function W2_NameFromPtr(ByVal p As Long) As String
    Dim ch As Integer
    Dim s  As String
    Do While p <> 0
        ch = W2_Peek2(p)
        If ch = 0 Then Exit Do
        s = s & ChrW$(ch)
        p = p + 2
    Loop
    W2_NameFromPtr = s
End Function

'================================================
' W2_NameToId
'================================================
Private Function W2_NameToId(ByVal nm As String) As Long
    Select Case LCase$(nm)
    Case "_remotehost":  W2_NameToId = 0
    Case "remoteport":   W2_NameToId = 1
    Case "localport":    W2_NameToId = 2
    Case "protocol":     W2_NameToId = 3
    Case "remotehostip": W2_NameToId = 4
    Case "localhostname": W2_NameToId = 5
    Case "localip":      W2_NameToId = 6
    Case "sockethandle": W2_NameToId = 7
    Case "state":        W2_NameToId = 8
    Case "bytesreceived": W2_NameToId = 9
    Case "remotehost":   W2_NameToId = 10
    Case "connect":      W2_NameToId = 64
    Case "listen":       W2_NameToId = 65
    Case "accept":       W2_NameToId = 66
    Case "senddata":     W2_NameToId = 67
    Case "getdata":      W2_NameToId = 68
    Case "peekdata":     W2_NameToId = 69
    Case "close":        W2_NameToId = 70
    Case "bind":         W2_NameToId = 71
    Case "aboutbox":     W2_NameToId = -552
    Case Else:           W2_NameToId = W2_ARG_NOT_FOUND
    End Select
End Function

'================================================
' W2_GetIDsOfNames
'================================================
Private Function W2_GetIDsOfNames(ByRef pThis As W2_IFACE, ByVal riid As Long, ByVal rgszNames As Long, ByVal cNames As Long, ByVal lcid As Long, ByVal rgdispid As Long) As Long
    Dim i  As Long
    Dim pp As Long
    Dim id As Long

    If rgszNames = 0 Or cNames = 0 Or rgdispid = 0 Then
        W2_GetIDsOfNames = W2_E_POINTER
        Exit Function
    End If

    pp = rgdispid
    For i = 0 To cNames - 1
        id = W2_NameToId(W2_NameFromPtr(W2_Peek4(rgszNames + i * 4)))
        If id = W2_ARG_NOT_FOUND Then
            W2_Poke4 pp, -1
            W2_GetIDsOfNames = W2_DISP_E_UNKNOWNNAME
        Else
            W2_Poke4 pp, id
        End If
        pp = pp + 4
    Next i
End Function

'================================================
' W2_Invoke
'================================================
Private Function W2_Invoke(ByRef pThis As W2_IFACE, ByVal dispidMember As Long, ByVal riid As Long, ByVal lcid As Long, ByVal wFlags As Long, ByVal pDispParams As Long, ByVal pVarResult As Long, ByVal pExcepInfo As Long, ByVal puArgErr As Long) As Long
    Dim v      As Variant
    Dim arg0   As Variant
    Dim arg1   As Variant
    Dim arg2   As Variant
    Dim dataV  As Variant
    Dim refPtr As Long
    Dim s      As String
    Dim l      As Long
    Dim sh     As Integer
    Dim zero   As Long
    Dim ei     As W2_EXCEPINFO

    On Error GoTo Failed

    W2_ClearResult pVarResult



    ' ---- propput -------------------------------------------------------------
    If (wFlags And 4) <> 0 Then
        Select Case dispidMember
        Case 3:  W2_LetProtocol pThis, W2_ArgLong(pDispParams, 0)
        Case 1:  W2_LetRemotePort pThis, W2_ArgLong(pDispParams, 0)
        Case 2:  W2_LetLocalPort pThis, W2_ArgLong(pDispParams, 0)
        Case 0:  W2_Let_RemoteHost pThis, W2_ArgStr(pDispParams, 0)
        Case 10: W2_LetRemoteHost pThis, W2_ArgStr(pDispParams, 0)
        Case Else: W2_Invoke = W2_DISP_E_MEMBERNOTFOUND
        End Select
        Exit Function
    End If

    ' ---- propget -------------------------------------------------------------
    If (wFlags And 2) <> 0 Then
        Select Case dispidMember
        Case 0:
            s = "": W2_Get_RemoteHost pThis, s: v = s: W2_PutResult pVarResult, v
        Case 1:
            l = 0: W2_GetRemotePort pThis, l: v = l: W2_PutResult pVarResult, v
        Case 2:
            l = 0: W2_GetLocalPort pThis, l: v = l: W2_PutResult pVarResult, v
        Case 3:
            l = 0: W2_GetProtocol pThis, l: v = l: W2_PutResult pVarResult, v
        Case 4:
            s = "": W2_GetRemoteHostIP pThis, s: v = s: W2_PutResult pVarResult, v
        Case 5:
            s = "": W2_GetLocalHostName pThis, s: v = s: W2_PutResult pVarResult, v
        Case 6:
            s = "": W2_GetLocalIP pThis, s: v = s: W2_PutResult pVarResult, v
        Case 7:
            l = 0: W2_GetSocketHandle pThis, l: v = l: W2_PutResult pVarResult, v
        Case 8:
            sh = 0: W2_GetState pThis, sh: v = CInt(sh): W2_PutResult pVarResult, v
        Case 9:
            l = 0: W2_GetBytesReceived pThis, l: v = l: W2_PutResult pVarResult, v
        Case 10:
            s = "": W2_GetRemoteHost pThis, s: v = s: W2_PutResult pVarResult, v
        Case Else: W2_Invoke = W2_DISP_E_MEMBERNOTFOUND
        End Select
        Exit Function
    End If

    ' ---- methods -------------------------------------------------------------
    Select Case dispidMember
    Case -552:
        W2_AboutBox pThis

    Case 64:
        If W2_ArgMissing(pDispParams, 0) Then
            If W2_ArgMissing(pDispParams, 1) Then
                W2_Connect pThis
            Else
                W2_Connect pThis, , W2_ArgV(pDispParams, 1)
            End If
        Else
            If W2_ArgMissing(pDispParams, 1) Then
                W2_Connect pThis, W2_ArgV(pDispParams, 0)
            Else
                W2_Connect pThis, W2_ArgV(pDispParams, 0), W2_ArgV(pDispParams, 1)
            End If
        End If

    Case 65:
        W2_Listen pThis

    Case 66:
        W2_Accept pThis, W2_ArgLong(pDispParams, 0)

    Case 67:
        arg0 = W2_ArgV(pDispParams, 0)
        W2_SendData pThis, arg0

    Case 68, 69:
        ' GetData/PeekData take the data buffer ByRef: the caller passes
        ' VT_BYREF|VT_VARIANT pointing at its own Variant.
        refPtr = 0
        arg0 = W2_ArgV(pDispParams, 0)
        If (VarType(arg0) And &H4000) <> 0 Then
            refPtr = W2_Peek4(VarPtr(arg0) + 8)
            If refPtr <> 0 Then MemoryCopy dataV, ByVal refPtr, 16
        Else
            dataV = arg0
        End If
        If W2_ArgMissing(pDispParams, 1) Then arg1 = Empty Else arg1 = W2_ArgV(pDispParams, 1)
        If W2_ArgMissing(pDispParams, 2) Then arg2 = Empty Else arg2 = W2_ArgV(pDispParams, 2)
        If dispidMember = 68 Then
            W2_GetData pThis, dataV, arg1, arg2
        Else
            W2_PeekData pThis, dataV, arg1, arg2
        End If
        If refPtr <> 0 Then
            MemoryCopy ByVal refPtr, dataV, 16
            MemoryCopy dataV, zero, 2
        End If

    Case 70:
        W2_Close pThis

    Case 71:
        If W2_ArgMissing(pDispParams, 0) Then
            If W2_ArgMissing(pDispParams, 1) Then
                W2_Bind pThis
            Else
                W2_Bind pThis, , W2_ArgV(pDispParams, 1)
            End If
        Else
            If W2_ArgMissing(pDispParams, 1) Then
                W2_Bind pThis, W2_ArgV(pDispParams, 0)
            Else
                W2_Bind pThis, W2_ArgV(pDispParams, 0), W2_ArgV(pDispParams, 1)
            End If
        End If

    Case Else:
        W2_Invoke = W2_DISP_E_MEMBERNOTFOUND
    End Select
    Exit Function

Failed:
    ' A VB method called through a hand-built vtable gets no framework help in
    ' filling EXCEPINFO. Without it the client turns the raise into
    ' RPC_E_DISCONNECTED (0x80010108) instead of the real error number.
    ei.wCode = 0
    ei.wReserved = 0
    ei.bstrSource = SysAllocString(StrPtr(Err.Source))
    ei.bstrDescription = SysAllocString(StrPtr(Err.Description))
    ei.bstrHelpFile = 0
    ei.dwHelpContext = 0
    ei.pvReserved = 0
    ei.scode = &H800A0000 Or (Err.number And &HFFFF&)
    ei.pfnDeferredFillIn = 0
    If pExcepInfo <> 0 Then MemoryCopy ByVal pExcepInfo, ei, LenB(ei)
    ' Return the VB error HRESULT itself: a VB6 client maps 0x800Axxxx back to
    ' the original error number even without IErrorInfo.
    W2_Invoke = &H800A0000 Or (Err.number And &HFFFF&)
    Err.Clear
End Function

'==============================================================================
' IMSWinsockControl properties
'==============================================================================

'================================================
' W2_AboutBox
'================================================
Private Function W2_AboutBox(ByRef pThis As W2_IFACE) As Long
    MsgBox "VB6 Winsock replacement (modWinsock2) v" & W2_VERSION & vbCrLf & _
           "Pure VB6 + Win32 API. No OCX, no registration.", _
           vbInformation Or vbOKOnly, "About"
End Function

'================================================
' W2_GetProtocol
'================================================
Private Function W2_GetProtocol(ByRef pThis As W2_IFACE, ByRef pRetVal As Long) As Long
    Dim idx As Long
    idx = pThis.StateIndex
    If W2_ValidState(idx) Then pRetVal = gStates(idx).Protocol
End Function

'================================================
' W2_LetProtocol
'================================================
Private Function W2_LetProtocol(ByRef pThis As W2_IFACE, ByVal pValue As Long) As Long
    Dim idx As Long
    idx = pThis.StateIndex
    If Not W2_ValidState(idx) Then Exit Function
    If pValue <> sckTCPProtocol And pValue <> sckUDPProtocol Then
        Err.Raise sckInvalidArg, "Winsock.Protocol", "Invalid protocol value."
    End If
    If gStates(idx).State <> sckClosed Then
        Err.Raise sckInvalidOp, "Winsock.Protocol", "The protocol cannot be changed while the socket is open."
    End If
    gStates(idx).Protocol = pValue
End Function

'================================================
' W2_GetRemoteHostIP
'================================================
Private Function W2_GetRemoteHostIP(ByRef pThis As W2_IFACE, ByRef pRetVal As String) As Long
    Dim idx As Long
    idx = pThis.StateIndex
    If W2_ValidState(idx) Then pRetVal = gStates(idx).RemoteHostIP
End Function

'================================================
' W2_GetLocalHostName
'================================================
Private Function W2_GetLocalHostName(ByRef pThis As W2_IFACE, ByRef pRetVal As String) As Long
    Dim idx As Long
    idx = pThis.StateIndex
    If W2_ValidState(idx) Then pRetVal = gStates(idx).LocalHostName
End Function

'================================================
' W2_GetLocalIP
'================================================
Private Function W2_GetLocalIP(ByRef pThis As W2_IFACE, ByRef pRetVal As String) As Long
    Dim idx As Long
    idx = pThis.StateIndex
    If W2_ValidState(idx) Then
        If Len(gStates(idx).LocalIP) > 0 Then
            pRetVal = gStates(idx).LocalIP
        Else
            pRetVal = gLocalIP
        End If
    End If
End Function

'================================================
' W2_GetSocketHandle
'================================================
Private Function W2_GetSocketHandle(ByRef pThis As W2_IFACE, ByRef pRetVal As Long) As Long
    Dim idx As Long
    idx = pThis.StateIndex
    If W2_ValidState(idx) Then pRetVal = gStates(idx).SocketHandle
End Function

'================================================
' W2_Get_RemoteHost
'================================================
Private Function W2_Get_RemoteHost(ByRef pThis As W2_IFACE, ByRef pRetVal As String) As Long
    Dim idx As Long
    idx = pThis.StateIndex
    If W2_ValidState(idx) Then pRetVal = gStates(idx).RemoteHost
End Function

'================================================
' W2_Let_RemoteHost
'================================================
Private Function W2_Let_RemoteHost(ByRef pThis As W2_IFACE, ByVal pValue As String) As Long
    Dim idx As Long
    idx = pThis.StateIndex
    If W2_ValidState(idx) Then gStates(idx).RemoteHost = pValue
End Function

'================================================
' W2_GetRemotePort
'================================================
Private Function W2_GetRemotePort(ByRef pThis As W2_IFACE, ByRef pRetVal As Long) As Long
    Dim idx As Long
    idx = pThis.StateIndex
    If W2_ValidState(idx) Then pRetVal = gStates(idx).RemotePort
End Function

'================================================
' W2_LetRemotePort
'================================================
Private Function W2_LetRemotePort(ByRef pThis As W2_IFACE, ByVal pValue As Long) As Long
    Dim idx As Long
    idx = pThis.StateIndex
    If Not W2_ValidState(idx) Then Exit Function
    If pValue < 0 Or pValue > 65535 Then
        Err.Raise sckOutOfRange, "Winsock.RemotePort", "Port must be between 0 and 65535."
    End If
    gStates(idx).RemotePort = pValue
End Function

'================================================
' W2_GetLocalPort
'================================================
Private Function W2_GetLocalPort(ByRef pThis As W2_IFACE, ByRef pRetVal As Long) As Long
    Dim idx As Long
    idx = pThis.StateIndex
    If W2_ValidState(idx) Then pRetVal = gStates(idx).LocalPort
End Function

'================================================
' W2_LetLocalPort
'================================================
Private Function W2_LetLocalPort(ByRef pThis As W2_IFACE, ByVal pValue As Long) As Long
    Dim idx As Long
    idx = pThis.StateIndex
    If Not W2_ValidState(idx) Then Exit Function
    If pValue < 0 Or pValue > 65535 Then
        Err.Raise sckOutOfRange, "Winsock.LocalPort", "Port must be between 0 and 65535."
    End If
    If gStates(idx).Bound Or gStates(idx).SocketHandle <> W2_INVALID_SOCKET Then
        Err.Raise sckInvalidOp, "Winsock.LocalPort", "LocalPort cannot be changed after the socket is bound."
    End If
    gStates(idx).LocalPort = pValue
End Function

'================================================
' W2_GetState
'================================================
Private Function W2_GetState(ByRef pThis As W2_IFACE, ByRef pRetVal As Integer) As Long
    Dim idx As Long
    idx = pThis.StateIndex
    If W2_ValidState(idx) Then pRetVal = CInt(gStates(idx).State)
End Function

'================================================
' W2_GetBytesReceived
'================================================
Private Function W2_GetBytesReceived(ByRef pThis As W2_IFACE, ByRef pRetVal As Long) As Long
    Dim idx As Long
    idx = pThis.StateIndex
    If W2_ValidState(idx) Then pRetVal = gStates(idx).BytesReceived
End Function

'==============================================================================
' IMSWinsockControl methods
'==============================================================================

'================================================
' W2_Connect
'================================================
Private Function W2_Connect(ByRef pThis As W2_IFACE, Optional ByVal pRemoteHost As Variant, Optional ByVal pRemotePort As Variant) As Long
    Dim idx As Long
    Dim v   As Variant

    idx = pThis.StateIndex
    If Not W2_ValidState(idx) Then Exit Function

    If gStates(idx).State <> sckClosed Then
        Err.Raise sckBadState, "Winsock.Connect", "Wrong protocol or connection state for the requested transaction or request."
    End If

    If Not IsMissing(pRemoteHost) Then
        v = pRemoteHost
        If Len(CStr(v)) > 0 Then gStates(idx).RemoteHost = CStr(v)
    End If
    If Not IsMissing(pRemotePort) Then
        v = pRemotePort
        If IsNumeric(v) Then
            If CLng(v) < 0 Or CLng(v) > 65535 Then
                Err.Raise sckOutOfRange, "Winsock.Connect", "Port must be between 0 and 65535."
            End If
            gStates(idx).RemotePort = CLng(v)
        End If
    End If

    If Len(gStates(idx).RemoteHost) = 0 Then
        Err.Raise sckAddressNotAvailable, "Winsock.Connect", "No remote host specified."
    End If

    If W2_CreateSocket(idx) <> 0 Then Exit Function

    If W2_IsIPLiteral(gStates(idx).RemoteHost) Then
        gStates(idx).RemoteHostIP = gStates(idx).RemoteHost
        gStates(idx).State = sckConnecting
        W2_DoConnect idx
    Else
        gStates(idx).State = sckResolvingHost
        W2_StartResolve idx
    End If
End Function

'================================================
' W2_Listen
'================================================
Private Function W2_Listen(ByRef pThis As W2_IFACE) As Long
    Dim idx As Long
    Dim r   As Long
    Dim t   As W2_IFACE

    idx = pThis.StateIndex
    If Not W2_ValidState(idx) Then Exit Function

    If gStates(idx).Protocol <> sckTCPProtocol Then
        Err.Raise sckWrongProtocol, "Winsock.Listen", "Listen requires the TCP protocol."
    End If
    If gStates(idx).State <> sckClosed And gStates(idx).State <> sckOpen Then
        Err.Raise sckBadState, "Winsock.Listen", "Wrong protocol or connection state for the requested transaction or request."
    End If

    If gStates(idx).SocketHandle = W2_INVALID_SOCKET Then
        If W2_CreateSocket(idx) <> 0 Then Exit Function
    End If

    If Not gStates(idx).Bound Then
        t.StateIndex = idx
        W2_Bind t
        If gStates(idx).State = sckError Then Exit Function
    End If

    r = listen(gStates(idx).SocketHandle, 5&)
    If r = W2_SOCKET_ERROR Then
        W2_SetError idx, WSAGetLastError()
        Exit Function
    End If

    gStates(idx).State = sckListening
    If Not W2_RegisterSocket(idx) Then
        W2_SetError idx, WSAGetLastError()
        Exit Function
    End If
    W2_UpdateLocalRemote idx
End Function

'================================================
' W2_Accept
'================================================
Private Function W2_Accept(ByRef pThis As W2_IFACE, ByVal pRequestID As Long) As Long
    Dim idx   As Long
    Dim key   As String
    Dim sock  As Long
    Dim owner As Long
    Dim v     As Variant

    idx = pThis.StateIndex
    If Not W2_ValidState(idx) Then Exit Function

    If gStates(idx).State <> sckClosed Or gStates(idx).SocketHandle <> W2_INVALID_SOCKET Then
        Err.Raise sckInvalidOp, "Winsock.Accept", "Accept must be called on a closed, unused Winsock object."
    End If

    key = CStr(pRequestID)
    On Error GoTo BadRequest
    v = gPendingSocks.Item(key)
    sock = CLng(v)
    v = gPendingOwners.Item(key)
    owner = CLng(v)
    On Error GoTo 0

    gPendingSocks.Remove key
    On Error Resume Next
    gPendingOwners.Remove key
    If W2_ValidState(owner) Then gStates(owner).Pending.Remove key
    On Error GoTo 0

    gStates(idx).SocketHandle = sock
    gStates(idx).Bound = True
    gStates(idx).Protocol = sckTCPProtocol
    gStates(idx).State = sckConnected

    If Not W2_RegisterSocket(idx) Then
        closesocket sock
        gStates(idx).SocketHandle = W2_INVALID_SOCKET
        gStates(idx).State = sckError
        W2_RaiseErrorEvent idx, WSAGetLastError(), W2_ErrorText(WSAGetLastError())
        Exit Function
    End If

    W2_UpdateLocalRemote idx
    Exit Function

BadRequest:
    Err.Raise sckInvalidArg, "Winsock.Accept", "The requestID is not valid or has expired."
End Function

' SendData counts bytes, not characters: the ANSI conversion can be shorter than
' Len(s), and using Len(s) made the send read past the converted array.
'================================================
' W2_SendData
'================================================
Private Function W2_SendData(ByRef pThis As W2_IFACE, ByVal pData As Variant) As Long
    Dim idx   As Long
    Dim s     As String
    Dim b()   As Byte
    Dim arr() As Byte
    Dim n     As Long
    Dim lb    As Long
    Dim t     As W2_IFACE

    idx = pThis.StateIndex
    If Not W2_ValidState(idx) Then Exit Function

    If gStates(idx).Protocol = sckUDPProtocol Then
        If gStates(idx).State = sckClosed Then
            t.StateIndex = idx
            W2_Bind t
            If gStates(idx).State = sckError Then Exit Function
        End If
        If gStates(idx).State <> sckOpen And gStates(idx).State <> sckConnected Then
            Err.Raise sckBadState, "Winsock.SendData", "Wrong protocol or connection state for the requested transaction or request."
        End If
    Else
        If gStates(idx).State <> sckConnected Then
            Err.Raise sckBadState, "Winsock.SendData", "Wrong protocol or connection state for the requested transaction or request."
        End If
    End If

    ' MSWINSCK ignores Null, Empty and Nothing instead of raising.
    If VarType(pData) = vbNull Or VarType(pData) = vbEmpty Then Exit Function
    If IsObject(pData) Then
        If pData Is Nothing Then Exit Function
    End If

    ' Byte arrays are handed to the send path as-is: copying them into a staging
    ' array first is the difference between 1 MB and 4 MB of memory traffic per
    ' SendData.
    Select Case VarType(pData)
    Case vbString
        s = CStr(pData)
        If Len(s) <= 0 Then Exit Function
        arr = StrConv(s, vbFromUnicode)
        n = UBound(arr) - LBound(arr) + 1
        If n <= 0 Then Exit Function
        W2_SendOut idx, arr, 0, n

    Case vbArray + vbByte
        b = pData
        n = W2_ArrayLen(b)
        If n <= 0 Then Exit Function
        lb = LBound(b)
        W2_SendOut idx, b, lb, n

    Case Else
        Err.Raise sckUnsupported, "Winsock.SendData", "SendData accepts a String or a Byte array."
    End Select
End Function

' Dispatches a send. When the socket is idle the bytes go straight to send();
' only an unsent remainder is staged in the write buffer.
'================================================
' W2_SendOut
'================================================
Private Sub W2_SendOut(ByVal idx As Long, ByRef src() As Byte, ByVal off As Long, ByVal n As Long)
    Dim sent As Long
    Dim e    As Long

    If Not W2_ValidState(idx) Then Exit Sub

    If gStates(idx).Protocol = sckUDPProtocol Then
        W2_SendUdp idx, src, off, n
        Exit Sub
    End If

    If gStates(idx).WriteSize > gStates(idx).WriteOffset Then
        W2_WriteAppend idx, src, off, n
        gStates(idx).SendPending = True
        W2_FlushSend idx
        Exit Sub
    End If

    sent = send(gStates(idx).SocketHandle, src(off), n, 0&)
    If sent = W2_SOCKET_ERROR Then
        e = WSAGetLastError()
        If e <> W2_WSAEWOULDBLOCK Then
            W2_SetError idx, e
            Exit Sub
        End If
        sent = 0
    ElseIf sent < 0 Then
        sent = 0
    End If

    If sent >= n Then
        gStates(idx).SendTotal = n
        W2_RaiseSimple idx, W2_EV_SENDPROGRESS, n, 0
        If W2_ValidState(idx) Then W2_RaiseSimple idx, W2_EV_SENDCOMPLETE, 0, 0
        Exit Sub
    End If

    ' Partial write: queue what is left and let FD_WRITE finish it.
    gStates(idx).SendTotal = sent
    gStates(idx).SendPending = True
    If sent < n Then W2_WriteAppend idx, src, off + sent, n - sent
End Sub

'================================================
' W2_GetData
'================================================
Private Function W2_GetData(ByRef pThis As W2_IFACE, ByRef pData As Variant, ByVal pType As Variant, ByVal pMaxLen As Variant) As Long
    W2_DoGetData pThis, pData, pType, pMaxLen, False
End Function

'================================================
' W2_PeekData
'================================================
Private Function W2_PeekData(ByRef pThis As W2_IFACE, ByRef pData As Variant, ByVal pType As Variant, ByVal pMaxLen As Variant) As Long
    W2_DoGetData pThis, pData, pType, pMaxLen, True
End Function

'================================================
' W2_Close
'================================================
Private Function W2_Close(ByRef pThis As W2_IFACE) As Long
    Dim idx As Long
    Dim e   As Long

    idx = pThis.StateIndex
    If Not W2_ValidState(idx) Then Exit Function

    With gStates(idx)
        Select Case .State
        Case sckConnected
            If .Protocol = sckTCPProtocol Then
                ' MSWINSCK is closed by the time Close returns, so the control can
                ' be reused immediately. Send a best-effort FIN, then release the
                ' socket now instead of waiting for the peer's FIN.
                If shutdown(.SocketHandle, W2_SD_SEND) <> 0 Then
                    e = WSAGetLastError()
                    If e <> W2_WSAENOTCONN And e <> W2_WSAENOTSOCK Then
                        ' ignore; close anyway
                    End If
                End If
                W2_DestroySocket idx
                .State = sckClosed
                .CloseRequested = False
            Else
                W2_DestroySocket idx
                .State = sckClosed
                .CloseRequested = False
            End If

        Case sckListening
            ' MSWINSCK keeps pending requests valid after the listener is
            ' closed, which is what makes the classic server pattern
            ' "Close" then "Accept requestID" work. Anything never accepted is
            ' closed when the object is freed.
            W2_DestroySocket idx
            .State = sckClosed

        Case sckResolvingHost, sckHostResolved, sckConnecting, sckConnectionPending
            W2_CancelDns idx
            W2_DestroySocket idx
            .State = sckClosed

        Case sckOpen, sckError, sckClosing
            W2_CancelDns idx
            W2_ClosePending idx
            W2_DestroySocket idx
            .State = sckClosed
            .CloseRequested = False

        Case sckClosed
            ' nothing to do
        End Select
    End With
End Function

'================================================
' W2_Bind
'================================================
Private Function W2_Bind(ByRef pThis As W2_IFACE, Optional ByVal pLocalPort As Variant, Optional ByVal pLocalIP As Variant) As Long
    Dim idx  As Long
    Dim addr As SOCKADDR
    Dim r    As Long
    Dim port As Long
    Dim ip   As String
    Dim a    As Long
    Dim v    As Variant

    idx = pThis.StateIndex
    If Not W2_ValidState(idx) Then Exit Function

    If gStates(idx).State <> sckClosed And gStates(idx).State <> sckOpen Then
        Err.Raise sckBadState, "Winsock.Bind", "Wrong protocol or connection state for the requested transaction or request."
    End If
    If gStates(idx).Bound Then
        Err.Raise sckInvalidOp, "Winsock.Bind", "The socket is already bound."
    End If

    If Not IsMissing(pLocalPort) Then
        v = pLocalPort
        If IsNumeric(v) Then
            port = CLng(v)
            If port < 0 Or port > 65535 Then
                Err.Raise sckOutOfRange, "Winsock.Bind", "Port must be between 0 and 65535."
            End If
            gStates(idx).LocalPort = port
        End If
    End If
    If Not IsMissing(pLocalIP) Then
        v = pLocalIP
        gStates(idx).LocalIP = CStr(v)
    End If

    If gStates(idx).SocketHandle = W2_INVALID_SOCKET Then
        If W2_CreateSocket(idx) <> 0 Then Exit Function
    End If

    ip = gStates(idx).LocalIP
    If Len(ip) > 0 Then
        a = inet_addr(ip)
        If a = W2_INADDR_NONE And ip <> "255.255.255.255" Then
            Err.Raise sckInvalidArg, "Winsock.Bind", "LocalIP is not a valid IPv4 address."
        End If
    Else
        a = W2_INADDR_ANY
    End If

    With addr
        .sin_family = W2_AF_INET
        .sin_addr = a
        .sin_port = htons(W2_ToInt(gStates(idx).LocalPort))
    End With

    r = bind(gStates(idx).SocketHandle, addr, Len(addr))
    If r = W2_SOCKET_ERROR Then
        W2_SetError idx, WSAGetLastError()
        Exit Function
    End If

    gStates(idx).Bound = True
    gStates(idx).State = sckOpen
    If Not W2_RegisterSocket(idx) Then
        W2_SetError idx, WSAGetLastError()
        Exit Function
    End If
    W2_UpdateLocalRemote idx
End Function

'================================================
' W2_GetRemoteHost
'================================================
Private Function W2_GetRemoteHost(ByRef pThis As W2_IFACE, ByRef pRetVal As String) As Long
    Dim idx As Long
    idx = pThis.StateIndex
    If W2_ValidState(idx) Then pRetVal = gStates(idx).RemoteHost
End Function

'================================================
' W2_LetRemoteHost
'================================================
Private Function W2_LetRemoteHost(ByRef pThis As W2_IFACE, ByVal pValue As String) As Long
    Dim idx As Long
    idx = pThis.StateIndex
    If W2_ValidState(idx) Then gStates(idx).RemoteHost = pValue
End Function

'==============================================================================
' IMSWinsockControl implementation helpers
'==============================================================================

'================================================
' W2_CreateSocket
'================================================
Private Function W2_CreateSocket(ByVal idx As Long) As Long
    Dim s As Long

    If Not W2_ValidState(idx) Then W2_CreateSocket = -1: Exit Function
    If gStates(idx).SocketHandle <> W2_INVALID_SOCKET Then Exit Function

    If gStates(idx).Protocol = sckUDPProtocol Then
        s = socket(W2_AF_INET, W2_SOCK_DGRAM, W2_IPPROTO_UDP)
    Else
        s = socket(W2_AF_INET, W2_SOCK_STREAM, W2_IPPROTO_TCP)
    End If

    If s = W2_INVALID_SOCKET Then
        W2_SetError idx, WSAGetLastError()
        W2_CreateSocket = -1
        Exit Function
    End If

    gStates(idx).SocketHandle = s
    gStates(idx).SocketRegistered = False

    If Not W2_RegisterSocket(idx) Then
        closesocket s
        gStates(idx).SocketHandle = W2_INVALID_SOCKET
        W2_SetError idx, WSAGetLastError()
        W2_CreateSocket = -1
        Exit Function
    End If
End Function

'================================================
' W2_RegisterSocket
'================================================
Private Function W2_RegisterSocket(ByVal idx As Long) As Boolean
    Dim mask As Long
    Dim s    As Long
    Dim key  As String

    If Not W2_ValidState(idx) Then Exit Function
    s = gStates(idx).SocketHandle
    If s = W2_INVALID_SOCKET Then Exit Function
    If gWin = 0 Then Exit Function

    If gStates(idx).Protocol = sckUDPProtocol Then
        mask = W2_FD_READ Or W2_FD_WRITE
        If gStates(idx).UdpConnected Then mask = mask Or W2_FD_CONNECT
    ElseIf gStates(idx).State = sckListening Then
        mask = W2_FD_ACCEPT Or W2_FD_CLOSE
    Else
        mask = W2_FD_CONNECT Or W2_FD_READ Or W2_FD_WRITE Or W2_FD_CLOSE
    End If

    If WSAAsyncSelect(s, gWin, gMsgSock, mask) = W2_SOCKET_ERROR Then
        W2_RegisterSocket = False
        Exit Function
    End If

    key = CStr(s)
    On Error Resume Next
    gSockets.Remove key
    gSockets.Add idx, key
    On Error GoTo 0

    gStates(idx).SocketRegistered = True
    W2_RegisterSocket = True
End Function

'================================================
' W2_DestroySocket
'================================================
Private Sub W2_DestroySocket(ByVal idx As Long)
    Dim s   As Long
    Dim key As String

    If Not W2_ValidState(idx) Then Exit Sub
    s = gStates(idx).SocketHandle
    If s = W2_INVALID_SOCKET Then Exit Sub

    key = CStr(s)
    On Error Resume Next
    If gWin <> 0 Then WSAAsyncSelect s, gWin, 0&, 0&
    gSockets.Remove key
    On Error GoTo 0

    closesocket s
    W2_DrainSocketMessages s
    gStates(idx).SocketHandle = W2_INVALID_SOCKET
    gStates(idx).SocketRegistered = False
    gStates(idx).Bound = False
    gStates(idx).UdpConnected = False

    Erase gStates(idx).WriteData
    gStates(idx).WriteCap = 0
    gStates(idx).WriteSize = 0
    gStates(idx).WriteOffset = 0
    gStates(idx).SendTotal = 0
    gStates(idx).SendPending = False
End Sub

' A socket handle can be reused while a notification posted for the previous
' socket is still queued. WSAAsyncSelect delivers to the window keyed by handle,
' so that stale message would be applied to the new owner - typically a spurious
' FD_CLOSE that tears down a brand new listener. Drain the queue for a handle the
' moment its socket is destroyed; nothing can legitimately arrive for it after
' that point.
'================================================
' W2_DrainSocketMessages
'================================================
Private Sub W2_DrainSocketMessages(ByVal s As Long)
    Dim m      As W2_MSG
    Dim keep() As W2_MSG
    Dim nKeep  As Long
    Dim i      As Long

    If gWin = 0 Or gMsgSock = 0 Then Exit Sub
    If s = W2_INVALID_SOCKET Then Exit Sub

    Do While PeekMessage_Any(m, gWin, gMsgSock, gMsgSock, W2_PM_REMOVE) <> 0
        If m.wParam <> s Then
            ReDim Preserve keep(0 To nKeep)
            keep(nKeep) = m
            nKeep = nKeep + 1
        End If
    Loop

    For i = 0 To nKeep - 1
        PostMessage gWin, keep(i).message, keep(i).wParam, keep(i).lParam
    Next
End Sub

'================================================
' W2_UpdateLocalRemote
'================================================
Private Sub W2_UpdateLocalRemote(ByVal idx As Long)
    Dim addr As SOCKADDR
    Dim l    As Long
    Dim ip   As String

    If Not W2_ValidState(idx) Then Exit Sub
    If gStates(idx).SocketHandle = W2_INVALID_SOCKET Then Exit Sub

    l = Len(addr)
    If getsockname(gStates(idx).SocketHandle, addr, l) = 0 Then
        gStates(idx).LocalPort = W2_ToUInt(ntohs(addr.sin_port))
        ip = W2_LongToIP(addr.sin_addr)
        If ip <> "0.0.0.0" Then gStates(idx).LocalIP = ip
    End If

    l = Len(addr)
    If getpeername(gStates(idx).SocketHandle, addr, l) = 0 Then
        gStates(idx).RemotePort = W2_ToUInt(ntohs(addr.sin_port))
        gStates(idx).RemoteHostIP = W2_LongToIP(addr.sin_addr)
    End If
End Sub

'================================================
' W2_DoConnect
'================================================
Private Sub W2_DoConnect(ByVal idx As Long)
    Dim addr As SOCKADDR
    Dim r    As Long
    Dim e    As Long

    If Not W2_ValidState(idx) Then Exit Sub

    With addr
        .sin_family = W2_AF_INET
        .sin_addr = inet_addr(gStates(idx).RemoteHostIP)
        .sin_port = htons(W2_ToInt(gStates(idx).RemotePort))
    End With

    r = Connect(gStates(idx).SocketHandle, addr, Len(addr))
    If r = 0 Then
        If gStates(idx).Protocol = sckUDPProtocol Then gStates(idx).UdpConnected = True
        gStates(idx).State = sckConnected
        W2_UpdateLocalRemote idx
        W2_RaiseSimple idx, W2_EV_CONNECT, 0, 0
    Else
        e = WSAGetLastError()
        If e = W2_WSAEWOULDBLOCK Or e = W2_WSAEINPROGRESS Or e = W2_WSAEALREADY Then
            gStates(idx).State = sckConnecting
        Else
            W2_SetError idx, e
        End If
    End If
End Sub

'==============================================================================
' Asynchronous DNS
'==============================================================================

'================================================
' W2_StartResolve
'================================================
Private Sub W2_StartResolve(ByVal idx As Long)
    Dim h       As Long
    Dim req     As Long
    Dim errCode As Long

    If Not W2_ValidState(idx) Then Exit Sub

    h = GlobalAlloc(W2_GMEM_FIXED, W2_MAXGETHOSTSTRUCT)
    If h = 0 Then
        gStates(idx).State = sckError
        W2_DestroySocket idx
        W2_RaiseErrorEvent idx, sckOutOfMemory, "Out of memory."
        Exit Sub
    End If

    req = WSAAsyncGetHostByName(gWin, gMsgDns, gStates(idx).RemoteHost, ByVal h, W2_MAXGETHOSTSTRUCT)
    If req = 0 Then
        errCode = WSAGetLastError()
        GlobalFree h
        If errCode = 0 Then errCode = W2_WSAHOST_NOT_FOUND
        W2_SetError idx, errCode
        Exit Sub
    End If

    gStates(idx).DnsReq = req
    gStates(idx).DnsHandle = h
    gStates(idx).DnsPending = True
    gDnsMap.Add idx, CStr(req)
End Sub

'================================================
' W2_OnResolve
'================================================
Private Sub W2_OnResolve(ByVal req As Long, ByVal e As Long)
    Dim idx   As Long
    Dim h     As Long
    Dim he    As HOSTENT
    Dim pList As Long
    Dim a     As Long
    Dim key   As String

    If Not gReady Then Exit Sub
    On Error Resume Next
    Err.Clear
    key = CStr(req)
    idx = CLng(gDnsMap.Item(key))
    If Err.number <> 0 Then
        Err.Clear
        Exit Sub
    End If
    If Not W2_ValidState(idx) Then Exit Sub

    W2_AddRefIndex idx
    gDnsMap.Remove key

    h = gStates(idx).DnsHandle

    If e = 0 Then
        MemoryCopy he, ByVal h, Len(he)
        MemoryCopy pList, ByVal he.hAddrList, 4
        MemoryCopy a, ByVal pList, 4
        gStates(idx).RemoteHostIP = W2_LongToIP(a)
        gStates(idx).State = sckHostResolved
        W2_DoConnect idx
    Else
        W2_SetError idx, e
    End If

    If h <> 0 Then GlobalFree h
    If W2_ValidState(idx) Then
        gStates(idx).DnsHandle = 0
        gStates(idx).DnsReq = 0
        gStates(idx).DnsPending = False
    End If
    W2_ReleaseIndex idx
End Sub

'================================================
' W2_CancelDns
'================================================
Private Sub W2_CancelDns(ByVal idx As Long)
    Dim req As Long
    Dim h   As Long
    Dim key As String

    If Not W2_ValidState(idx) Then Exit Sub
    If Not gStates(idx).DnsPending Then Exit Sub

    req = gStates(idx).DnsReq
    h = gStates(idx).DnsHandle
    key = CStr(req)

    On Error Resume Next
    If req <> 0 Then WSACancelAsyncRequest req
    If h <> 0 Then GlobalFree h
    gDnsMap.Remove key
    On Error GoTo 0

    gStates(idx).DnsReq = 0
    gStates(idx).DnsHandle = 0
    gStates(idx).DnsPending = False
End Sub

'==============================================================================
' Socket events
'==============================================================================

'================================================
' W2_OnSocketEvent
'================================================
Private Sub W2_OnSocketEvent(ByVal idx As Long, ByVal eventId As Long, ByVal e As Long)
    If Not W2_ValidState(idx) Then Exit Sub
    W2_AddRefIndex idx
    On Error GoTo EH

    If e <> 0 Then
        If eventId = W2_FD_CONNECT Then
            W2_OnConnectError idx, e
        ElseIf eventId = W2_FD_CLOSE Then
            W2_OnCloseEvent idx
        Else
            W2_SetError idx, e
        End If
    Else
        Select Case eventId
        Case W2_FD_READ:    W2_OnRead idx
        Case W2_FD_WRITE:   W2_OnWrite idx
        Case W2_FD_ACCEPT:  W2_OnAccept idx
        Case W2_FD_CONNECT: W2_OnConnected idx
        Case W2_FD_CLOSE:   W2_OnCloseEvent idx
        End Select
    End If

    W2_ReleaseIndex idx
    Exit Sub
EH:
    W2_ReleaseIndex idx
End Sub

'================================================
' W2_OnConnected
'================================================
Private Sub W2_OnConnected(ByVal idx As Long)
    If Not W2_ValidState(idx) Then Exit Sub
    If gStates(idx).State = sckConnected Then Exit Sub
    If gStates(idx).State <> sckConnecting And gStates(idx).State <> sckHostResolved Then Exit Sub
    If gStates(idx).Protocol = sckUDPProtocol Then gStates(idx).UdpConnected = True
    gStates(idx).State = sckConnected
    W2_UpdateLocalRemote idx
    W2_RaiseSimple idx, W2_EV_CONNECT, 0, 0
End Sub

'================================================
' W2_OnConnectError
'================================================
Private Sub W2_OnConnectError(ByVal idx As Long, ByVal e As Long)
    If Not W2_ValidState(idx) Then Exit Sub
    W2_SetError idx, e
End Sub

'================================================
' W2_OnRead
'================================================
Private Sub W2_OnRead(ByVal idx As Long)
    Dim buf() As Byte
    Dim n     As Long
    Dim e     As Long
    Dim addr  As SOCKADDR
    Dim l     As Long
    Dim avail As Long
    Dim d()   As Byte
    Dim v     As Variant

    If Not W2_ValidState(idx) Then Exit Sub
    If gStates(idx).SocketHandle = W2_INVALID_SOCKET Then Exit Sub

    If gStates(idx).Protocol = sckUDPProtocol Then
        If gUdpCap < 65536 Then
            ReDim gUdpBuf(0 To 65535)
            gUdpCap = 65536
        End If
        Do
            l = Len(addr)
            n = recvfrom(gStates(idx).SocketHandle, gUdpBuf(0), 65536, 0&, addr, l)
            If n > 0 Then
                ReDim d(0 To n - 1)
                MemoryCopy d(0), gUdpBuf(0), n
                gStates(idx).UdpData.Add d

                gStates(idx).RemotePort = W2_ToUInt(ntohs(addr.sin_port))
                gStates(idx).RemoteHostIP = W2_LongToIP(addr.sin_addr)
                gStates(idx).RemoteHost = gStates(idx).RemoteHostIP
                gStates(idx).BytesReceived = n

                W2_RaiseSimple idx, W2_EV_DATA, n, 0
                If Not W2_ValidState(idx) Then Exit Sub
                If gStates(idx).SocketHandle = W2_INVALID_SOCKET Then Exit Sub
            Else
                e = WSAGetLastError()
                If e <> W2_WSAEWOULDBLOCK Then W2_SetError idx, e
                Exit Do
            End If
        Loop
        Exit Sub
    End If

    ReDim buf(0 To 8191)
    Do
        n = recv(gStates(idx).SocketHandle, buf(0), 8192, 0&)
        If n > 0 Then
            W2_ReadAppend idx, buf, n
        ElseIf n = 0 Then
            Exit Do
        Else
            e = WSAGetLastError()
            If e = W2_WSAEWOULDBLOCK Then Exit Do
            W2_SetError idx, e
            Exit Sub
        End If
    Loop

    avail = W2_Available(idx)
    If avail > 0 Then
        W2_RaiseSimple idx, W2_EV_DATA, avail, 0
    End If
End Sub

'================================================
' W2_OnWrite
'================================================
Private Sub W2_OnWrite(ByVal idx As Long)
    If Not W2_ValidState(idx) Then Exit Sub
    If Not gStates(idx).SendPending Then Exit Sub
    If gStates(idx).WriteSize > gStates(idx).WriteOffset Then
        W2_FlushSend idx
    End If
End Sub

'================================================
' W2_OnAccept
'================================================
Private Sub W2_OnAccept(ByVal idx As Long)
    Dim addr As SOCKADDR
    Dim l    As Long
    Dim s    As Long
    Dim e    As Long
    Dim req  As Long
    Dim key  As String

    If Not W2_ValidState(idx) Then Exit Sub
    If gStates(idx).State <> sckListening Then Exit Sub

    l = Len(addr)
    s = accept(gStates(idx).SocketHandle, addr, l)
    If s = W2_INVALID_SOCKET Then
        e = WSAGetLastError()
        If e <> W2_WSAEWOULDBLOCK Then
            W2_RaiseErrorEvent idx, e, W2_ErrorText(e)
        End If
        Exit Sub
    End If

    gNextRequestID = gNextRequestID + 1
    If gNextRequestID <= 0 Then gNextRequestID = 1
    req = gNextRequestID
    key = CStr(req)

    gPendingSocks.Add s, key
    gPendingOwners.Add idx, key
    gStates(idx).Pending.Add key

    W2_RaiseSimple idx, W2_EV_CONNREQ, req, 0
End Sub

'================================================
' W2_OnCloseEvent
'================================================
Private Sub W2_OnCloseEvent(ByVal idx As Long)
    Dim avail      As Long
    Dim localClose As Boolean

    If Not W2_ValidState(idx) Then Exit Sub

    localClose = gStates(idx).CloseRequested

    If gStates(idx).Protocol = sckTCPProtocol And gStates(idx).SocketHandle <> W2_INVALID_SOCKET Then
        W2_DrainRead idx
        avail = W2_Available(idx)
        If avail > 0 Then
            W2_RaiseSimple idx, W2_EV_DATA, avail, 0
            If Not W2_ValidState(idx) Then Exit Sub
        End If
    End If

    W2_DestroySocket idx
    Erase gStates(idx).ReadData
    gStates(idx).ReadCap = 0
    gStates(idx).ReadSize = 0
    gStates(idx).ReadOffset = 0
    gStates(idx).BytesReceived = 0
    gStates(idx).State = sckClosed
    gStates(idx).CloseRequested = False

    ' The OCX raises Close for a remote close, not for a close initiated by
    ' the local application.
    If Not localClose Then
        W2_RaiseSimple idx, W2_EV_CLOSE, 0, 0
    End If
End Sub

'================================================
' W2_DrainRead
'================================================
Private Sub W2_DrainRead(ByVal idx As Long)
    Dim buf() As Byte
    Dim n     As Long

    If Not W2_ValidState(idx) Then Exit Sub
    If gStates(idx).SocketHandle = W2_INVALID_SOCKET Then Exit Sub

    ReDim buf(0 To 8191)
    Do
        n = recv(gStates(idx).SocketHandle, buf(0), 8192, 0&)
        If n > 0 Then
            W2_ReadAppend idx, buf, n
        Else
            Exit Do
        End If
    Loop
End Sub

'================================================
' W2_SetError
'================================================
Private Sub W2_SetError(ByVal idx As Long, ByVal e As Long)
    If Not W2_ValidState(idx) Then Exit Sub
    gStates(idx).State = sckError
    W2_DestroySocket idx
    W2_RaiseErrorEvent idx, e, W2_ErrorText(e)
End Sub

'==============================================================================
' TCP buffers and send
'==============================================================================

'================================================
' W2_ReadAppend
'================================================
Private Sub W2_ReadAppend(ByVal idx As Long, ByRef src() As Byte, ByVal n As Long)
    Dim need As Long

    If n <= 0 Then Exit Sub
    If Not W2_ValidState(idx) Then Exit Sub

    If gStates(idx).ReadOffset > 0 Then W2_ReadCompact idx

    need = gStates(idx).ReadSize + n
    If need > gStates(idx).ReadCap Then
        gStates(idx).ReadCap = W2_GrowCap(need)
        If gStates(idx).ReadSize = 0 And gStates(idx).ReadCap > 0 Then
            ReDim gStates(idx).ReadData(0 To gStates(idx).ReadCap - 1)
        Else
            ReDim Preserve gStates(idx).ReadData(0 To gStates(idx).ReadCap - 1)
        End If
    End If

    MemoryCopy gStates(idx).ReadData(gStates(idx).ReadSize), src(0), n
    gStates(idx).ReadSize = gStates(idx).ReadSize + n
    gStates(idx).BytesReceived = gStates(idx).ReadSize - gStates(idx).ReadOffset
End Sub

'================================================
' W2_ReadCompact
'================================================
Private Sub W2_ReadCompact(ByVal idx As Long)
    Dim n As Long
    If Not W2_ValidState(idx) Then Exit Sub
    n = gStates(idx).ReadSize - gStates(idx).ReadOffset
    If n > 0 Then
        MemoryCopy gStates(idx).ReadData(0), gStates(idx).ReadData(gStates(idx).ReadOffset), n
    End If
    gStates(idx).ReadSize = n
    gStates(idx).ReadOffset = 0
    gStates(idx).BytesReceived = n
End Sub

'================================================
' W2_ReadConsume
'================================================
Private Sub W2_ReadConsume(ByVal idx As Long, ByVal n As Long)
    If Not W2_ValidState(idx) Then Exit Sub
    gStates(idx).ReadOffset = gStates(idx).ReadOffset + n
    If gStates(idx).ReadOffset >= gStates(idx).ReadSize Then
        gStates(idx).ReadOffset = 0
        gStates(idx).ReadSize = 0
        gStates(idx).ReadCap = 0
        Erase gStates(idx).ReadData
    ElseIf gStates(idx).ReadOffset > 4096 And (gStates(idx).ReadOffset * 2) > gStates(idx).ReadSize Then
        W2_ReadCompact idx
    End If
    gStates(idx).BytesReceived = gStates(idx).ReadSize - gStates(idx).ReadOffset
End Sub

'================================================
' W2_Available
'================================================
Private Function W2_Available(ByVal idx As Long) As Long
    If Not W2_ValidState(idx) Then Exit Function
    If gStates(idx).Protocol = sckUDPProtocol Then
        W2_Available = W2_UdpFirstSize(idx)
    Else
        W2_Available = gStates(idx).ReadSize - gStates(idx).ReadOffset
    End If
End Function

'================================================
' W2_UdpFirstSize
'================================================
Private Function W2_UdpFirstSize(ByVal idx As Long) As Long
    Dim d() As Byte

    If Not W2_ValidState(idx) Then Exit Function
    If gStates(idx).UdpData.Count = 0 Then Exit Function

    On Error Resume Next
    Err.Clear
    d = gStates(idx).UdpData.Item(1)
    If Err.number <> 0 Then
        Err.Clear
        Exit Function
    End If
    If UBound(d) >= LBound(d) Then W2_UdpFirstSize = UBound(d) - LBound(d) + 1
End Function

'================================================
' W2_WriteAppend
'================================================
Private Sub W2_WriteAppend(ByVal idx As Long, ByRef src() As Byte, ByVal off As Long, ByVal n As Long)
    Dim need As Long

    If n <= 0 Then Exit Sub
    If Not W2_ValidState(idx) Then Exit Sub

    If gStates(idx).WriteOffset > 0 Then W2_WriteCompact idx

    need = gStates(idx).WriteSize + n
    If need > gStates(idx).WriteCap Then
        gStates(idx).WriteCap = W2_GrowCap(need)
        If gStates(idx).WriteSize = 0 And gStates(idx).WriteCap > 0 Then
            ReDim gStates(idx).WriteData(0 To gStates(idx).WriteCap - 1)
        Else
            ReDim Preserve gStates(idx).WriteData(0 To gStates(idx).WriteCap - 1)
        End If
    End If

    MemoryCopy gStates(idx).WriteData(gStates(idx).WriteSize), src(off), n
    gStates(idx).WriteSize = gStates(idx).WriteSize + n
End Sub

'================================================
' W2_WriteCompact
'================================================
Private Sub W2_WriteCompact(ByVal idx As Long)
    Dim n As Long
    If Not W2_ValidState(idx) Then Exit Sub
    n = gStates(idx).WriteSize - gStates(idx).WriteOffset
    If n > 0 Then
        MemoryCopy gStates(idx).WriteData(0), gStates(idx).WriteData(gStates(idx).WriteOffset), n
    End If
    gStates(idx).WriteSize = n
    gStates(idx).WriteOffset = 0
End Sub

'================================================
' W2_WriteReset
'================================================
Private Sub W2_WriteReset(ByVal idx As Long)
    If Not W2_ValidState(idx) Then Exit Sub
    Erase gStates(idx).WriteData
    gStates(idx).WriteCap = 0
    gStates(idx).WriteSize = 0
    gStates(idx).WriteOffset = 0
    gStates(idx).SendTotal = 0
    gStates(idx).SendPending = False
End Sub

'================================================
' W2_FlushSend
'================================================
Private Sub W2_FlushSend(ByVal idx As Long)
    Dim n         As Long
    Dim e         As Long
    Dim remaining As Long
    Dim total     As Long

    If Not W2_ValidState(idx) Then Exit Sub
    If gStates(idx).SocketHandle = W2_INVALID_SOCKET Then Exit Sub

    If gStates(idx).WriteOffset > 0 Then W2_WriteCompact idx
    total = gStates(idx).SendTotal

    Do While gStates(idx).WriteOffset < gStates(idx).WriteSize
        n = send(gStates(idx).SocketHandle, gStates(idx).WriteData(gStates(idx).WriteOffset), gStates(idx).WriteSize - gStates(idx).WriteOffset, 0&)
        If n = W2_SOCKET_ERROR Then
            e = WSAGetLastError()
            If e = W2_WSAEWOULDBLOCK Then Exit Do
            W2_WriteReset idx
            W2_SetError idx, e
            Exit Sub
        ElseIf n <= 0 Then
            Exit Do
        End If
        gStates(idx).WriteOffset = gStates(idx).WriteOffset + n
        total = total + n
        gStates(idx).SendTotal = total
    Loop

    remaining = gStates(idx).WriteSize - gStates(idx).WriteOffset

    If remaining <= 0 Then
        W2_WriteReset idx
        W2_RaiseSimple idx, W2_EV_SENDPROGRESS, total, 0
        If W2_ValidState(idx) Then W2_RaiseSimple idx, W2_EV_SENDCOMPLETE, 0, 0
    Else
        If total > 0 Then
            W2_RaiseSimple idx, W2_EV_SENDPROGRESS, total, remaining
        End If
    End If
End Sub

'================================================
' W2_SendUdp
'================================================
Private Sub W2_SendUdp(ByVal idx As Long, ByRef src() As Byte, ByVal off As Long, ByVal n As Long)
    Dim addr  As SOCKADDR
    Dim r     As Long
    Dim e     As Long
    Dim a     As Long
    Dim he    As Long
    Dim host  As HOSTENT
    Dim pList As Long

    If Not W2_ValidState(idx) Then Exit Sub
    If gStates(idx).RemotePort <= 0 Then
        Err.Raise sckAddressNotAvailable, "Winsock.SendData", "RemotePort is not set."
    End If
    If Len(gStates(idx).RemoteHost) = 0 And Len(gStates(idx).RemoteHostIP) = 0 Then
        Err.Raise sckAddressNotAvailable, "Winsock.SendData", "RemoteHost is not set."
    End If

    If Len(gStates(idx).RemoteHostIP) = 0 Then
        If W2_IsIPLiteral(gStates(idx).RemoteHost) Then
            gStates(idx).RemoteHostIP = gStates(idx).RemoteHost
        Else
            he = gethostbyname(gStates(idx).RemoteHost)
            If he = 0 Then
                W2_RaiseErrorEvent idx, W2_WSAHOST_NOT_FOUND, W2_ErrorText(W2_WSAHOST_NOT_FOUND)
                Exit Sub
            End If
            MemoryCopy host, ByVal he, Len(host)
            MemoryCopy pList, ByVal host.hAddrList, 4
            MemoryCopy a, ByVal pList, 4
            gStates(idx).RemoteHostIP = W2_LongToIP(a)
        End If
    End If

    a = inet_addr(gStates(idx).RemoteHostIP)
    With addr
        .sin_family = W2_AF_INET
        .sin_addr = a
        .sin_port = htons(W2_ToInt(gStates(idx).RemotePort))
    End With

    If gStates(idx).UdpConnected And gStates(idx).State = sckConnected Then
        r = send(gStates(idx).SocketHandle, src(off), n, 0&)
    Else
        r = sendto(gStates(idx).SocketHandle, src(off), n, 0&, addr, Len(addr))
    End If

    If r = W2_SOCKET_ERROR Then
        e = WSAGetLastError()
        W2_RaiseErrorEvent idx, e, W2_ErrorText(e)
        Exit Sub
    End If

    ' MSWINSCK raises no SendProgress/SendComplete for datagrams: a datagram is
    ' delivered atomically, so there is nothing to report progress against.
End Sub

'==============================================================================
' GetData / PeekData
'==============================================================================

'================================================
' W2_DoGetData
'================================================
Private Sub W2_DoGetData(ByRef pThis As W2_IFACE, ByRef pData As Variant, ByVal pType As Variant, ByVal pMaxLen As Variant, ByVal peek As Boolean)
    Dim idx      As Long
    Dim avail    As Long
    Dim vt       As Long
    Dim maxLen   As Long
    Dim vtSize   As Long
    Dim n        As Long
    Dim s        As String
    Dim b()      As Byte
    Dim d()      As Byte
    Dim v        As Variant
    Dim srcPtr   As Long
    Dim srcIsUdp As Boolean
    Dim udpWhole As Boolean
    Dim tmpByte  As Byte
    Dim tmpInt   As Integer
    Dim tmpLng   As Long
    Dim tmpSng   As Single
    Dim tmpDbl   As Double
    Dim tmpCur   As Currency
    Dim tmpDat   As Date
    Dim tmpBoo   As Boolean

    idx = pThis.StateIndex
    If Not W2_ValidState(idx) Then Exit Sub

    avail = W2_Available(idx)
    If avail <= 0 Then
        Err.Raise sckBadState, "Winsock.GetData", "No data is available in the receive buffer."
    End If

    If IsMissing(pType) Then
        If VarType(pData) = vbEmpty Then
            vt = vbString
        Else
            vt = VarType(pData)
        End If
    ElseIf IsEmpty(pType) Then
        If VarType(pData) = vbEmpty Then
            vt = vbString
        Else
            vt = VarType(pData)
        End If
    Else
        vt = CLng(pType)
    End If

    maxLen = avail
    If Not IsMissing(pMaxLen) Then
        If Not IsEmpty(pMaxLen) Then
            On Error Resume Next
            Err.Clear
            maxLen = CLng(pMaxLen)
            If Err.number <> 0 Then
                Err.Clear
                On Error GoTo 0
                Err.Raise sckInvalidArg, "Winsock.GetData", "MaxLen is not a valid number."
            End If
            On Error GoTo 0
        End If
    End If
    If maxLen <= 0 Then maxLen = avail
    If maxLen > avail Then maxLen = avail

    If gStates(idx).Protocol = sckUDPProtocol Then
        srcIsUdp = True
        d = gStates(idx).UdpData.Item(1)
        If UBound(d) >= LBound(d) Then
            srcPtr = VarPtr(d(LBound(d)))
            udpWhole = ((UBound(d) - LBound(d) + 1) = n)
        Else
            srcPtr = 0
        End If
    Else
        srcIsUdp = False
        If gStates(idx).ReadSize > gStates(idx).ReadOffset Then
            srcPtr = VarPtr(gStates(idx).ReadData(gStates(idx).ReadOffset))
        Else
            srcPtr = 0
        End If
    End If

    If srcPtr = 0 Then
        Err.Raise sckBadState, "Winsock.GetData", "No data is available in the receive buffer."
    End If

    vtSize = W2_TypeSize(vt)
    If vtSize > 0 Then
        If avail < vtSize Then
            Err.Raise sckBadState, "Winsock.GetData", "Not enough data for the requested type."
        End If
        n = vtSize
    Else
        n = maxLen
    End If

    Select Case vt
    Case vbString
        s = String$(n, 0)
        CopyStrAny s, ByVal srcPtr, n
        pData = s

    Case vbArray + vbByte
        If udpWhole Then
            ' Whole datagram: d is already a private copy, hand it over.
            pData = d
        Else
            ReDim b(0 To n - 1)
            MemoryCopy b(0), ByVal srcPtr, n
            pData = b
        End If

    Case vbByte
        MemoryCopy tmpByte, ByVal srcPtr, 1
        pData = tmpByte

    Case vbInteger
        MemoryCopy tmpInt, ByVal srcPtr, 2
        pData = tmpInt

    Case vbLong
        MemoryCopy tmpLng, ByVal srcPtr, 4
        pData = tmpLng

    Case vbSingle
        MemoryCopy tmpSng, ByVal srcPtr, 4
        pData = tmpSng

    Case vbDouble
        MemoryCopy tmpDbl, ByVal srcPtr, 8
        pData = tmpDbl

    Case vbCurrency
        MemoryCopy tmpCur, ByVal srcPtr, 8
        pData = tmpCur

    Case vbDate
        MemoryCopy tmpDat, ByVal srcPtr, 8
        pData = tmpDat

    Case vbBoolean
        MemoryCopy tmpBoo, ByVal srcPtr, 2
        pData = tmpBoo

    Case Else
        Err.Raise sckUnsupported, "Winsock.GetData", "Unsupported data type."
    End Select

    If Not peek Then
        If srcIsUdp Then
            On Error Resume Next
            gStates(idx).UdpData.Remove 1
            On Error GoTo 0
            gStates(idx).BytesReceived = W2_UdpFirstSize(idx)
        Else
            W2_ReadConsume idx, n
        End If
    End If
End Sub

'==============================================================================
' IConnectionPointContainer / IConnectionPoint
'==============================================================================

'================================================
' W2_EnumConnectionPoints
'================================================
Private Function W2_EnumConnectionPoints(ByRef pThis As W2_IFACE, ByRef ppEnum As IEnumConnectionPoints) As Long
    W2_EnumConnectionPoints = W2_E_NOTIMPL
End Function

'================================================
' W2_FindConnectionPoint
'================================================
Private Function W2_FindConnectionPoint(ByRef pThis As W2_IFACE, ByRef riid As CLSID, ByRef ppCP As IConnectionPoint) As Long
    Dim idx As Long
    Dim p   As Long

    idx = pThis.StateIndex
    If Not W2_ValidState(idx) Then
        W2_FindConnectionPoint = W2_E_FAIL
        Exit Function
    End If

    If W2_IsEqualIID(riid, gIID_Events) Then
        p = gStates(idx).HeaderPtr + 16
        MemoryCopy ppCP, p, 4
        W2_AddRefIndex idx
        W2_FindConnectionPoint = 0
    Else
        W2_FindConnectionPoint = W2_CONNECT_E_NOCONNECTION
    End If
End Function

'================================================
' W2_GetConnectionInterface
'================================================
Private Function W2_GetConnectionInterface(ByRef pThis As W2_IFACE, ByRef piid As CLSID) As Long
    MemoryCopy piid, gIID_Events, 16
End Function

'================================================
' W2_GetConnectionPointContainer
'================================================
Private Function W2_GetConnectionPointContainer(ByRef pThis As W2_IFACE, ByRef ppCPC As IConnectionPointContainer) As Long
    Dim idx As Long
    Dim p   As Long

    idx = pThis.StateIndex
    If Not W2_ValidState(idx) Then
        W2_GetConnectionPointContainer = W2_E_FAIL
        Exit Function
    End If

    p = gStates(idx).HeaderPtr + 8
    MemoryCopy ppCPC, p, 4
    W2_AddRefIndex idx
End Function

'================================================
' W2_Advise
'================================================
Private Function W2_Advise(ByRef pThis As W2_IFACE, ByVal pObjectPtr As DMSWinsockControlEvents, ByRef pdwCookie As Long) As Long
    Dim idx    As Long
    Dim cookie As Long

    idx = pThis.StateIndex
    If Not W2_ValidState(idx) Then
        W2_Advise = W2_E_FAIL
        Exit Function
    End If
    If pObjectPtr Is Nothing Then
        W2_Advise = W2_E_POINTER
        Exit Function
    End If

    gStates(idx).NextCookie = gStates(idx).NextCookie + 1
    If gStates(idx).NextCookie <= 0 Then gStates(idx).NextCookie = 1
    cookie = gStates(idx).NextCookie

    On Error GoTo AdviseFail
    gStates(idx).Sinks.Add pObjectPtr, CStr(cookie)
    pdwCookie = cookie
    W2_Advise = 0
    Exit Function
AdviseFail:
    W2_Advise = W2_E_FAIL
End Function

'================================================
' W2_Unadvise
'================================================
Private Function W2_Unadvise(ByRef pThis As W2_IFACE, ByVal dwCookie As Long) As Long
    Dim idx As Long

    idx = pThis.StateIndex
    If Not W2_ValidState(idx) Then
        W2_Unadvise = W2_E_FAIL
        Exit Function
    End If

    On Error GoTo NotThere
    gStates(idx).Sinks.Remove CStr(dwCookie)
    W2_Unadvise = 0
    Exit Function
NotThere:
    W2_Unadvise = W2_CONNECT_E_NOCONNECTION
End Function

'================================================
' W2_EnumConnections
'================================================
Private Function W2_EnumConnections(ByRef pThis As W2_IFACE, ByRef ppEnum As IEnumConnections) As Long
    W2_EnumConnections = W2_E_NOTIMPL
End Function

'==============================================================================
' IEnumConnections stubs (v2 hook)
'==============================================================================

'================================================
' W2_EnumNext
'================================================
Private Function W2_EnumNext(ByRef pThis As W2_IFACE, ByVal pcConnections As Long, ByVal prgcd As Long, ByVal lpcFetched As Long) As Long
    W2_EnumNext = W2_E_NOTIMPL
End Function

'================================================
' W2_EnumSkip
'================================================
Private Function W2_EnumSkip(ByRef pThis As W2_IFACE, ByVal pcConnections As Long) As Long
    W2_EnumSkip = W2_E_NOTIMPL
End Function

'================================================
' W2_EnumReset
'================================================
Private Function W2_EnumReset(ByRef pThis As W2_IFACE) As Long
    W2_EnumReset = W2_E_NOTIMPL
End Function

'================================================
' W2_EnumClone
'================================================
Private Function W2_EnumClone(ByRef pThis As W2_IFACE, ByRef ppEnum As Long) As Long
    W2_EnumClone = W2_E_NOTIMPL
End Function

'==============================================================================
' Events
'==============================================================================
' Raises one event on one sink with by-value VARIANT arguments.
' A VB6 form sink reads a byref VARIANT as the pointer (the requestID arrived
' as 1702696), so the call is made through IDispatch::Invoke by dispid.

'================================================
' W2_RaiseSimple
'================================================
Private Sub W2_RaiseSimple(ByVal idx As Long, ByVal ev As Long, ByVal a1 As Long, ByVal a2 As Long)
    Dim n     As Long
    Dim k     As Long
    Dim arr() As DMSWinsockControlEvents
    Dim evObj As DMSWinsockControlEvents

    If Not W2_ValidState(idx) Then Exit Sub
    If gStates(idx).Sinks Is Nothing Then Exit Sub

    n = gStates(idx).Sinks.Count
    If n <= 0 Then Exit Sub

    ReDim arr(1 To n)
    For k = 1 To n
        Set arr(k) = gStates(idx).Sinks.Item(k)
    Next

    W2_AddRefIndex idx
    For k = 1 To n
        If Not W2_ValidState(idx) Then Exit For
        Set evObj = arr(k)
        On Error Resume Next
        ' The parentheses matter: without them VB passes a variable by
        ' reference to the dispinterface, and a VB6 form sink reads that
        ' pointer as the argument value (requestID arrived as 1702696).
        ' A parenthesised argument becomes a temporary, i.e. a by-value
        ' VARIANT, which is what the form's compiled sink expects.
        Select Case ev
        Case W2_EV_CONNECT:      evObj.Connect
        Case W2_EV_DATA:         evObj.DataArrival (a1)
        Case W2_EV_CONNREQ:      evObj.ConnectionRequest (a1)
        Case W2_EV_SENDPROGRESS: evObj.SendProgress (a1), (a2)
        Case W2_EV_SENDCOMPLETE: evObj.SendComplete
        Case W2_EV_CLOSE:        evObj.Close
        End Select
        On Error GoTo 0
        Set evObj = Nothing
    Next

    For k = 1 To n
        Set arr(k) = Nothing
    Next
    Erase arr
    W2_ReleaseIndex idx
End Sub

'================================================
' W2_RaiseErrorEvent
'================================================
Private Sub W2_RaiseErrorEvent(ByVal idx As Long, ByVal number As Long, ByVal desc As String)
    Dim n      As Long
    Dim k      As Long
    Dim cancel As Boolean
    Dim arr()  As DMSWinsockControlEvents
    Dim evObj  As DMSWinsockControlEvents

    cancel = False

    If W2_ValidState(idx) Then
        If Not gStates(idx).Sinks Is Nothing Then
            n = gStates(idx).Sinks.Count
            If n > 0 Then
                ReDim arr(1 To n)
                For k = 1 To n
                    Set arr(k) = gStates(idx).Sinks.Item(k)
                Next

                W2_AddRefIndex idx
                For k = 1 To n
                    If Not W2_ValidState(idx) Then Exit For
                    Set evObj = arr(k)
                    On Error Resume Next
                    ' MSWINSCK reports Scode as 0x800A0000 Or <winsock error>,
                    ' e.g. -2146818227 (0x800A274D) for 10061.
                    evObj.Error CInt(number), desc, &H800A0000 Or number, "MSWinsock", "", 0&, cancel
                    On Error GoTo 0
                    Set evObj = Nothing
                Next

                For k = 1 To n
                    Set arr(k) = Nothing
                Next
                Erase arr
                W2_ReleaseIndex idx
            End If
        End If
    End If

    ' MSWINSCK opens a modal box when the host does not cancel the Error event.
    ' A library must not block the host's message loop from a socket callback, so
    ' the box is only shown when no sink is attached and the host explicitly asks
    ' for dialogs through the standard VB6 App.LogMode setting (2 = message box).
    If Not cancel And Not gInErrorBox Then
        If n = 0 And W2_ValidState(idx) Then
            If App.LogMode = 2 Then
                gInErrorBox = True
                On Error Resume Next
                MsgBox desc, vbExclamation, "Winsock"
                On Error GoTo 0
                gInErrorBox = False
            End If
        End If
    End If
End Sub

'==============================================================================
' Object destruction
'==============================================================================

'================================================
' W2_DestroyObject
'================================================
Private Sub W2_DestroyObject(ByVal idx As Long)
    Dim p As Long

    If Not W2_ValidState(idx) Then Exit Sub

    W2_CancelDns idx
    W2_ClosePending idx
    W2_DestroySocket idx

    p = gStates(idx).HeaderPtr

    Set gStates(idx).UdpData = Nothing
    Set gStates(idx).Sinks = Nothing
    Set gStates(idx).Pending = Nothing
    gStates(idx).RemoteHost = ""
    gStates(idx).RemoteHostIP = ""
    gStates(idx).LocalIP = ""
    gStates(idx).LocalHostName = ""
    Erase gStates(idx).ReadData
    Erase gStates(idx).WriteData
    gStates(idx).ReadCap = 0
    gStates(idx).ReadSize = 0
    gStates(idx).ReadOffset = 0
    gStates(idx).WriteCap = 0
    gStates(idx).WriteSize = 0
    gStates(idx).WriteOffset = 0
    gStates(idx).HeaderPtr = 0
    gStates(idx).InUse = False
    gStates(idx).NextFree = gFreeHead
    gFreeHead = idx

    If p <> 0 Then HeapFree GetProcessHeap(), 0, ByVal p

    gObjCount = gObjCount - 1
    If gObjCount < 0 Then gObjCount = 0
    ' Process-wide state (WSAStartup, message window, IIDs, vtables, state pool)
    ' is deliberately kept for the lifetime of the process. Tearing it down when
    ' the last object dies makes every create/release cycle pay for WSAStartup,
    ' CreateWindowEx and two RegisterWindowMessage calls again.
End Sub

'================================================
' W2_ClosePending
'================================================
Private Sub W2_ClosePending(ByVal idx As Long)
    Dim keys() As String
    Dim n      As Long
    Dim i      As Long
    Dim key    As String
    Dim s      As Long

    If Not W2_ValidState(idx) Then Exit Sub
    If gStates(idx).Pending Is Nothing Then Exit Sub

    n = gStates(idx).Pending.Count
    If n <= 0 Then Exit Sub

    ReDim keys(1 To n)
    For i = 1 To n
        keys(i) = CStr(gStates(idx).Pending.Item(i))
    Next

    For i = 1 To n
        key = keys(i)
        On Error Resume Next
        Err.Clear
        s = CLng(gPendingSocks.Item(key))
        If Err.number = 0 And s <> 0 Then closesocket s
        Err.Clear
        gPendingSocks.Remove key
        gPendingOwners.Remove key
        On Error GoTo 0
    Next

    Set gStates(idx).Pending = New Collection
End Sub

'==============================================================================
' Message window
'==============================================================================

'================================================
' W2_CreateMessageWindow
'================================================
Private Sub W2_CreateMessageWindow()
    gWin = CreateWindowEx(0&, "STATIC", "W2SOCK", 0&, 0&, 0&, 0&, 0&, W2_HWND_MESSAGE, 0&, App.hInstance, ByVal 0&)
    If gWin = 0 Then
        Err.Raise sckOutOfMemory, "New_Winsock2", "Unable to create the Winsock message window."
    End If

    gOldProc = SetWindowLong(gWin, W2_GWL_WNDPROC, AddressOf W2_WindowProc)
    gMsgSock = RegisterWindowMessage(App.EXEName & ".W2.Socket")
    gMsgDns = RegisterWindowMessage(App.EXEName & ".W2.Dns")

    If gMsgSock = 0 Then
        W2_DestroyMessageWindow
        Err.Raise sckOutOfMemory, "New_Winsock2", "Unable to register the Winsock window message."
    End If
End Sub

'================================================
' W2_DestroyMessageWindow
'================================================
Private Sub W2_DestroyMessageWindow()
    If gWin = 0 Then Exit Sub
    On Error Resume Next
    SetWindowLong gWin, W2_GWL_WNDPROC, gOldProc
    DestroyWindow gWin
    On Error GoTo 0
    gWin = 0
End Sub

'================================================
' W2_WindowProc
'================================================
Private Function W2_WindowProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim idx As Long
    Dim ev  As Long
    Dim e   As Long

    On Error Resume Next

    If gMsgSock <> 0 And uMsg = gMsgSock Then
        ev = lParam And &HFFFF&
        e = (lParam \ &H10000) And &HFFFF&
        idx = W2_FindSocket(wParam)
        If idx >= 0 Then W2_OnSocketEvent idx, ev, e
        W2_WindowProc = 0
        Exit Function
    End If

    If gMsgDns <> 0 And uMsg = gMsgDns Then
        e = (lParam \ &H10000) And &HFFFF&
        W2_OnResolve wParam, e
        W2_WindowProc = 0
        Exit Function
    End If

    W2_WindowProc = CallWindowProc(gOldProc, hWnd, uMsg, wParam, lParam)
End Function

'================================================
' W2_FindSocket
'================================================
Private Function W2_FindSocket(ByVal s As Long) As Long
    Dim v As Variant
    W2_FindSocket = -1
    If Not gReady Then Exit Function
    If gSockets Is Nothing Then Exit Function
    On Error Resume Next
    Err.Clear
    v = gSockets.Item(CStr(s))
    If Err.number = 0 Then W2_FindSocket = CLng(v)
End Function

'==============================================================================
' Local host information
'==============================================================================

'================================================
' W2_LoadLocalInfo
'================================================
Private Sub W2_LoadLocalInfo()
    Dim b()   As Byte
    Dim i     As Long
    Dim a()   As Byte
    Dim he    As Long
    Dim host  As HOSTENT
    Dim pList As Long
    Dim addr  As Long

    ReDim b(0 To 255)
    If gethostname_Any(b(0), 256&) = 0 Then
        i = 0
        Do While i <= 255
            If b(i) = 0 Then Exit Do
            i = i + 1
        Loop
        If i > 0 Then
            ReDim a(0 To i - 1)
            MemoryCopy a(0), b(0), i
            gLocalHostName = StrConv(a, vbUnicode)
        End If
    End If

    If Len(gLocalHostName) > 0 Then
        he = gethostbyname(gLocalHostName)
        If he <> 0 Then
            MemoryCopy host, ByVal he, Len(host)
            MemoryCopy pList, ByVal host.hAddrList, 4
            MemoryCopy addr, ByVal pList, 4
            gLocalIP = W2_LongToIP(addr)
        End If
    End If
End Sub

'==============================================================================
' Error text
'==============================================================================

'================================================
' W2_ErrorText
'================================================
Private Function W2_ErrorText(ByVal e As Long) As String
    Select Case e
    Case W2_WSAEWOULDBLOCK:         W2_ErrorText = "Resource temporarily unavailable."
    Case W2_WSAEINPROGRESS:         W2_ErrorText = "Operation now in progress."
    Case W2_WSAEALREADY:            W2_ErrorText = "Operation already in progress."
    Case W2_WSAENOTSOCK:            W2_ErrorText = "Socket operation on nonsocket."
    Case W2_WSAEDESTADDRREQ:        W2_ErrorText = "Destination address required."
    Case W2_WSAEMSGSIZE:            W2_ErrorText = "Message too long."
    Case 10041:                     W2_ErrorText = "Protocol wrong type for socket."
    Case 10042:                     W2_ErrorText = "Bad protocol option."
    Case 10043:                     W2_ErrorText = "Protocol not supported."
    Case 10044:                     W2_ErrorText = "Socket type not supported."
    Case 10045:                     W2_ErrorText = "Operation not supported."
    Case 10046:                     W2_ErrorText = "Protocol family not supported."
    Case 10047:                     W2_ErrorText = "Address family not supported by protocol family."
    Case W2_WSAEADDRINUSE:          W2_ErrorText = "Address already in use."
    Case W2_WSAEADDRNOTAVAIL:       W2_ErrorText = "Cannot assign requested address."
    Case W2_WSAENETDOWN:            W2_ErrorText = "Network is down."
    Case W2_WSAENETUNREACH:         W2_ErrorText = "Network is unreachable."
    Case W2_WSAENETRESET:           W2_ErrorText = "Network dropped connection on reset."
    Case W2_WSAECONNABORTED:        W2_ErrorText = "Software caused connection abort."
    Case W2_WSAECONNRESET:          W2_ErrorText = "Connection reset by peer."
    Case W2_WSAENOBUFS:             W2_ErrorText = "No buffer space available."
    Case W2_WSAEISCONN:             W2_ErrorText = "Socket is already connected."
    Case W2_WSAENOTCONN:            W2_ErrorText = "Socket is not connected."
    Case W2_WSAESHUTDOWN:           W2_ErrorText = "Cannot send after socket shutdown."
    Case W2_WSAETIMEDOUT:           W2_ErrorText = "Connection timed out."
    Case W2_WSAECONNREFUSED:        W2_ErrorText = "Connection is forcefully rejected"
    Case W2_WSAEHOSTUNREACH:        W2_ErrorText = "No route to host."
    Case W2_WSANOTINITIALISED:      W2_ErrorText = "Successful WSAStartup not yet performed."
    Case W2_WSAHOST_NOT_FOUND:      W2_ErrorText = "Host not found."
    Case W2_WSATRY_AGAIN:           W2_ErrorText = "Nonauthoritative host not found."
    Case W2_WSANO_RECOVERY:         W2_ErrorText = "This is a nonrecoverable error."
    Case W2_WSANO_DATA:             W2_ErrorText = "Valid name, no data record of requested type."
    Case Else:                      W2_ErrorText = "Winsock error " & e & "."
    End Select
End Function

#If W2DIAG Then

'==============================================================================
' Diagnostics
'==============================================================================
' Read-only counters for the test harness. Kept out of release builds: the
' constant is defined only by the FeatureTests project, so the shipped DLL
' exposes exactly the original public surface.
'================================================
' W2_Diagnostics
'================================================
Public Function W2_Diagnostics() As String
    Dim i     As Long
    Dim used  As Long
    Dim socks As Long
    Dim pend  As Long
    Dim udp   As Long
    Dim dns   As Long

    For i = 0 To gHighWater - 1
        If gStates(i).InUse Then
            used = used + 1
            If gStates(i).SocketHandle <> W2_INVALID_SOCKET Then socks = socks + 1
            If Not gStates(i).Pending Is Nothing Then pend = pend + gStates(i).Pending.Count
            If Not gStates(i).UdpData Is Nothing Then udp = udp + gStates(i).UdpData.Count
        End If
    Next
    If Not gDnsMap Is Nothing Then dns = gDnsMap.Count

    W2_Diagnostics = "objects=" & gObjCount & " slots=" & used & " high=" & gHighWater & _
                     " sockets=" & socks & " pending=" & pend & " udp=" & udp & _
                     " dns=" & dns & " window=" & gWin
End Function

#End If
