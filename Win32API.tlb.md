# Win32API.tlb — Type Library Reference

This document describes **only** the `Win32API.tlb` type library and how to
consume its declarations. It is intentionally independent of any application
that happens to use the file.

> `Win32API.tlb` is a declaration-only OLE type library, not a DLL/OCX. It
> contains no executable code. It is a large, general-purpose declaration
> library (Win32, COM, Winsock, DirectX, GDI+, audio, etc.), not a minimal
> Win32-only header replacement.

---

## 0. TL;DR

| Item | Value |
| --- | --- |
| File name | `Win32API.tlb` |
| Original copy | `Win32API.tlb` (same folder as this document) |
| Size | 697,868 bytes |
| SHA-256 | `B4ABBC10D5885EC41ED5757A84CFC7CFC2163750477285925AC88EFDC1773DA9` |
| File timestamp | 2025-11-14 09:38:39 (local) |
| Library name | `Win32Tlb_Lib` |
| LIBID | `{35904C2D-5F6A-4D7F-924D-570B700F7D42}` |
| Version | 1.0 |
| LCID | 1036 (`0x40C`, French) |
| syskind | 1 (`SYS_WIN32`) |
| Types | 677 |
| Modules | 41 |
| Module functions | 1,353 |
| Author/description | `EBArtSoft@ - Win32 API Declaration TypeLib For VB` |
| VB6 reference | `Reference=*\G{35904C2D-5F6A-4D7F-924D-570B700F7D42}#1.0#40c#Win32API.tlb#EBArtSoft@ - Win32 API Declaration TypeLib For VB` |

It is a **big** type library. Do not assume a function/constant is present just
because it exists in the Windows SDK; and do not assume a name is absent just
because it is not in this document — use the inspector in §9.

---

## 1. What it is / isn't

**It is:**

- An OLE/COM type library (`ITypeLib`) with 677 types and 41 modules.
- A VB6-friendly set of declarations for many Windows DLLs.
- A source of types such as `CLSID`, `SOCKADDR`, `HOSTENT`, `WSADATA`,
  `SAFEARRAY1D`, `ITypeInfo`, `ITypeLib`, `IProvideClassInfo`,
  connection-point interfaces, `IWinHttpRequest`, DirectX/DirectSound
  interfaces, FMOD, LUA, GDI+, and more.
- A re-declaration of the MSWINSCK interfaces/coclass with the same GUIDs
  (`IMSWinsockControl`, `DMSWinsockControlEvents`, `Winsock`) — see the
  appendix.

**It is not:**

- The `MSWINSCK.OCX` implementation. `New Winsock` uses CLSID
  `{248DD896-BB45-11CF-9ABC-0080C7E7B78D}` and therefore requires the OCX to be
  installed/registered. The tlb itself has no code.
- A complete replacement for the Windows SDK headers. Some APIs and constants
  are missing (see §7); the tlb is also full of non-Win32 declarations.
- Something to edit by hand. It is a binary type library.

The tlb is authored by EBArtSoft. There is no license file next to it; check
redistribution rights before shipping it in another project.

---

## 2. Library metadata

```
LIBNAME  = Win32Tlb_Lib
LIBGUID  = 35904c2d-5f6a-4d7f-924d-570b700f7d42
LCID     = 1036 (0x40C)
SYSKIND  = 1 (SYS_WIN32)
VERSION  = 1.0
```

The `.vbp` reference description is:

```
EBArtSoft@ - Win32 API Declaration TypeLib For VB
```

---

## 3. Type inventory

| Kind | Count |
| --- | ---: |
| `TKIND_ENUM` | 271 |
| `TKIND_RECORD` | 279 |
| `TKIND_MODULE` | 41 |
| `TKIND_INTERFACE` | 71 |
| `TKIND_COCLASS` | 7 |
| `TKIND_DISPATCH` | 5 |
| `TKIND_ALIAS` | 3 |
| **Total** | **677** |

### Coclasses (7)

`CShellLink`, `Winsock`, `WinHttpRequest`, `ActiveDesktop`, `TaskbarList`,
`directSound`, `DirectSound8`.

### Dispinterfaces (5)

`DMSWinsockControlEvents`, `IWinHttpRequest`, `VB6AddinFactory`, `VB6Addin`,
`VB6AddinEvent`.

### Aliases (3)

`HTTPREQUEST_PROXY_SETTING`, `HTTPREQUEST_SETCREDENTIALS_FLAGS`,
`VB6AddinFlagsEnum`.

### Modules and function counts (41 modules, 1,353 functions)

```
KERNEL32 (171)      NTDLL (1)           USER32 (224)
ADVAPI32 (38)       GDI32 (141)         COMDLG32 (10)
Version (6)         MSVBVM60 (36)       OLE32 (16)
OLEAUT32 (18)       OLEPRO32 (8)        WINMM (59)
LZ32 (10)           SHELL32 (17)        COMCTL32 (13)
ZLIB (9)            SHLWAPI (6)         WS2_32 (38)
IPHLPAPI (5)        PSAPI (17)          MSIMG32 (5)
UNRAR (7)           IJL11 (7)           GDIPLUS (79)
AVICAP32 (2)        CARDS (5)           FMOD (227)
DESDLL (6)          ASM (5)             LIBMAD (12)
LAME_ENC (5)        DINPUT8 (2)         IMAGEHLP (1)
LUA (145)           Constants (0)       DXMACRO_CONST (0)
D3d9 (2)            DSALG_GUID_CONST (0)  DSFX_GUID_CONST (0)
DSCFX_GUID_CONST (0)  DSInterfaces_GUID_CONST (0)
```

Notable: the tlb is not Win32-only. It contains **FMOD** (227 functions),
**LUA** (145), **GDIPLUS** (79), **USER32** (224), **KERNEL32** (171), and
various DirectX/DirectSound/audio/graphics declarations.

Key interfaces include `IUnknown`, `IDispatch`, `ITypeInfo`, `ITypeLib`,
`IProvideClassInfo`, `IConnectionPoint`, `IConnectionPointContainer`,
`IEnumConnections`, `IEnumConnectionPoints`, `IWinHttpRequest`,
`IWinHttpRequestEvents`, `IMSWinsockControl`, plus many DirectX/Sound/GDI+
interfaces.

---

## 4. Referencing it from VB6

Project in the same folder as the tlb:

```
Reference=*\G{35904C2D-5F6A-4D7F-924D-570B700F7D42}#1.0#40c#Win32API.tlb#EBArtSoft@ - Win32 API Declaration TypeLib For VB
```

Project in a subfolder:

```
Reference=*\G{35904C2D-5F6A-4D7F-924D-570B700F7D42}#1.0#40c#..\Win32API.tlb#EBArtSoft@ - Win32 API Declaration TypeLib For VB
```

Use LCID `40c` (1036) in the reference line. If the path is wrong, VB6 can
sometimes resolve the type library by GUID if it is registered; do not rely on
that — reference the file by path.

The declarations are VB6-friendly. Handles/pointers are generally `Long`,
strings are `String`, and records are `Type` declarations. Use VB6's Object
Browser (F2) to see the exact declared signature of any module function/type.
For a fabricated COM object, the tlb supplies the interface types and GUIDs;
see the appendix.

---

## 5. Win32 API coverage (curated)

The tlb declares 1,353 module functions. The most relevant Windows modules and
their sizes are:

| Module | Funcs | Typical contents |
| --- | ---: | --- |
| `KERNEL32` | 171 | memory, heaps, files, processes, threads, modules, strings, resources |
| `USER32` | 224 | windows, messages, dialogs, input, icons/cursors, rectangles |
| `GDI32` | 141 | DCs, bitmaps, brushes, pens, text, regions, paths, metafiles |
| `ADVAPI32` | 38 | registry, services, crypto, user name |
| `COMDLG32` | 10 | open/save/color/font/print dialogs |
| `Version` | 6 | file version information |
| `SHELL32` | 17 | ShellExecute, file info, file operations, icons, drag/drop |
| `COMCTL32` | 13 | InitCommonControls + flat scrollbar helpers (limited) |
| `SHLWAPI` | 6 | color conversion, path checks, autocomplete |
| `WINMM` | 59 | sound, MCI, wave I/O, joystick, multimedia timers |
| `PSAPI` | 17 | process/module/working-set information |
| `IPHLPAPI` | 5 | interface information |
| `MSIMG32` | 5 | AlphaBlend, GradientFill, TransparentBlt |
| `OLE32` | 16 | COM initialization, class creation, GUIDs, memory/streams |
| `OLEAUT32` | 18 | type libraries, BSTRs, dispatch/error info |
| `MSVBVM60` | 36 | VB runtime helpers (`ArrPtr`, etc.) |
| `WS2_32` | 38 | Winsock 1.1/2 API |
| `GDIPLUS` | 79 | GDI+ |
| `FMOD` | 227 | FMOD audio |
| `LUA` | 145 | Lua |
| `AVICAP32`, `CARDS`, `DESDLL`, `ASM`, `LIBMAD`, `LAME_ENC`, `DINPUT8`, `IMAGEHLP`, `ZLIB`, `UNRAR`, `IJL11`, `NTDLL`, `OLEPRO32`, `D3d9` | 1–12 each | niche/audio/graphics/compression |

### 5.1 KERNEL32 notable APIs

**Memory/heaps/pointers**

`GetProcessHeap`, `HeapAlloc`, `HeapFree`, `HeapReAlloc`, `HeapSize`,
`HeapValidate`, `HeapCreate`, `HeapDestroy`, `HeapLock`, `HeapUnlock`,
`HeapCompact`, `GlobalAlloc`, `GlobalFree`, `GlobalLock`, `GlobalUnlock`,
`GlobalSize`, `GlobalReAlloc`, `LocalAlloc`, `LocalFree`, `LocalLock`,
`LocalUnlock`, `VirtualAlloc`, `VirtualFree`, `VirtualProtect`,
`VirtualQuery`, `VirtualLock`, `VirtualUnlock`, `VirtualAllocEx`,
`VirtualProtectEx`, `IsBadReadPtr`, `IsBadWritePtr`, `IsBadCodePtr`,
`IsBadStringPtr`, `IsBadHugeReadPtr`, `IsBadHugeWritePtr`.

**Memory-copy helpers exposed by this tlb**

`MemoryCopy`, `MemoryFill`, `MemoryFillStr`, `MemoryZero`, `MemoryZeroStr`,
`CopyStrAny`, `CopyAnyStr`, `CopyArrAny`, `CopyAnyArr`, `CopyArrArr`,
`CopyStrStrW`.

**Files**

`CreateFile`, `ReadFile`, `ReadFileStr`, `CreatePipe`, `WriteFile` (see the
Object Browser for the exact overloads present), `CloseHandle`, `DeleteFile`,
`GetFileSize`, `GetFileSizeEx`, `CreateFileMapping`, `MapViewOfFile`,
`UnmapViewOfFile`, `GetFullPathName`, `GetTempPath`, `GetFileAttributes`,
`GetFileTime`, `FileTimeToSystemTime`, `GetSystemDirectory`.

**Processes/threads**

`GetCurrentProcess`, `GetCurrentProcessId`, `OpenProcess`, `CreateProcess`,
`TerminateProcess`, `GetExitCodeProcess`, `WaitForSingleObject`,
`GetCurrentThread`, `GetCurrentThreadId`, `OpenThread`, `ResumeThread`,
`SuspendThread`, `GetThreadContext`, `SetThreadContext`, `GetExitCodeThread`,
`ReadProcessMemory`, `WriteProcessMemory`, `TlsAlloc`, `TlsFree`,
`TlsGetValue`, `TlsSetValue`, `CreateToolhelp32Snapshot`, `ExitProcess`.

**Modules/resources/strings**

`LoadLibrary`, `LoadLibraryEx`, `FreeLibrary`, `GetProcAddress`,
`GetModuleHandle`, `GetModuleFileName`, `LoadResource`, `LockResource`,
`FreeResource`, `FindResource`, `FindResourceEx`, `SizeofResource`,
`EnumResourceNames`, `EnumResourceTypes`, `EnumResourceLanguages`,
`UpdateResource`, `BeginUpdateResource`, `EndUpdateResource`,
`lstrlen`, `lstrlenW`, `lstrcpy`, `lstrcpyn`, `lstrcat`, `lstrcmp`, `lstrcmpi`,
`strlen`, `strlenW`, `strcpy`, `strcpyW`, `strcat`, `strcatW`, `strcmp`,
`strcmpW`, `strcmpi`, `strcmpiW`, `strcpyn`, `strcpynW`.

**Misc**

`Sleep`, `SleepEx`, `GetLastError`, `GetTickCount`,
`QueryPerformanceCounter`, `QueryPerformanceFrequency`, `SetErrorMode`,
`GetComputerName`, `GetTempPath`, `FreeConsole`.

### 5.2 USER32 notable APIs

**Windows**

`CreateWindowEx`, `DestroyWindow`, `CreateMDIWindow`, `ShowWindow`,
`MoveWindow`, `SetWindowPos`, `FindWindow`, `FindWindowEx`, `EnumWindows`,
`EnumThreadWindows`, `GetWindow`, `GetTopWindow`, `GetDesktopWindow`,
`GetShellWindow`, `WindowFromPoint`, `WindowFromDC`, `IsWindow`, `IsChild`,
`IsIconic`, `IsZoomed`, `IsWindowVisible`, `IsWindowEnabled`,
`IsWindowUnicode`, `EnableWindow`, `CloseWindow`, `SetParent`, `GetParent`,
`GetWindowThreadProcessId`, `GetWindowText`, `SetWindowText`,
`GetWindowTextLength`, `GetWindowRect`, `GetClientRect`, `GetClassName`,
`GetWindowLong`, `SetWindowLong`, `GetClassLong`, `SetClassLong`,
`GetClassWord`, `SetClassWord`, `GetWindowRgn`, `SetWindowRgn`,
`GetProp`, `SetProp`, `RemoveProp`, `SetCapture`, `GetCapture`,
`ReleaseCapture`, `FlashWindow`, `FlashWindowEx`, `LockWindowUpdate`,
`UpdateWindow`, `InvalidateRect`, `PrintWindow`, `TileWindows`,
`CascadeWindows`, `GetActiveWindow`, `SetActiveWindow`.

**Messages**

`GetMessage`, `PeekMessage`, `PostMessage`, `SendMessage`, `PostQuitMessage`,
`DispatchMessage`, `TranslateMessage`, `WaitMessage`, `WaitForInputIdle`,
`DefWindowProc`, `CallWindowProc`.

**Dialogs/input**

`MessageBox`, `MessageBoxEx`, `MessageBoxLong`, `MessageBeep`, `DialogBoxParam`,
`EndDialog`, `GetKeyState`, `GetKeyboardState`, `GetAsyncKeyState`,
`ToAscii`, `ToAsciiEx`, `ToUnicode`, `ToUnicodeEx`, `mouse_event`,
`keybd_event`.

**Icons/cursors/rectangles/monitors**

`LoadIcon`, `LoadCursor`, `LoadCursorI`, `LoadImage`, `CopyIcon`, `DrawIcon`,
`DrawIconEx`, `CreateIcon`, `CreateIconIndirect`, `CreateIconFromResource`,
`GetIconInfo`, `DestroyIcon`, `SetCursor`, `GetCursor`, `ShowCursor`,
`ClipCursor`, `GetClipCursor`, `GetCursorPos`, `SetCursorPos`,
`CreateCursor`, `DestroyCursor`, `GetCursorInfo`, `SetRect`, `CopyRect`,
`PtInRect`, `UnionRect`, `EqualRect`, `OffsetRect`, `InflateRect`,
`IsRectEmpty`, `SetRectEmpty`, `SubtractRect`, `IntersectRect`,
`GetMonitorInfo`, `MonitorFromRect`, `MonitorFromPoint`, `MonitorFromWindow`,
`GetSystemMetrics`, `SystemParametersInfo`, `GetSysColor`,
`GetSysColorBrush`, `ClientToScreen`, `ScreenToClient`, `MapWindowPoints`.

### 5.3 GDI32 notable APIs

**DCs/text**

`CreateDC`, `CreateCompatibleDC`, `DeleteDC`, `SaveDC`, `RestoreDC`,
`ResetDC`, `CancelDC`, `GetDeviceCaps`, `TextOut`, `GetTextExtentPoint32`,
`GetTextAlign`, `SetTextAlign`, `SetTextColor`, `GetTextColor`,
`SetBkColor`, `GetBkColor`, `SetBkMode`, `GetBkMode`, `GetStockObject`,
`SelectObject`, `DeleteObject`, `GetObjectA`, `GetObjectType`,
`GetCurrentObject`, `CreateFontIndirect`.

**Bitmaps/transfers**

`CreateBitmap`, `CreateBitmapIndirect`, `CreateCompatibleBitmap`,
`CreateDIBSection`, `CreateDIBitmap`, `GetBitmapBits`, `SetBitmapBits`,
`GetBitmapDimensionEx`, `SetBitmapDimensionEx`, `SetDIBits`, `GetDIBits`,
`StretchDIBits`, `SetDIBitsToDevice`, `SetDIBColorTable`,
`GetDIBColorTable`, `BitBlt`, `PatBlt`, `PlgBlt`, `MaskBlt`, `StretchBlt`,
`GetPixel`, `SetPixel`, `SetPixelV`, `GetBrushOrgEx`, `SetBrushOrgEx`.

**Brushes/pens/shapes**

`CreateSolidBrush`, `CreateHatchBrush`, `CreatePatternBrush`,
`CreateBrushIndirect`, `CreateDIBPatternBrush`, `CreateDIBPatternBrushPt`,
`CreatePen`, `ExtCreatePen`, `CreatePenIndirect`, `Pie`, `Chord`, `Ellipse`,
`Polygon`, `Rectangle`, `RoundRect`, `PolyBezier`, `PolyPolygon`, `Arc`,
`ArcTo`, `AngleArc`, `LineTo`, `Polyline`, `MoveToEx`.

**Regions/paths/metafiles**

`PtInRegion`, `FillRgn`, `SelectClipRgn`, `SelectClipPath`,
`BeginPath`, `EndPath`, `FillPath`, `StrokePath`, `StrokeAndFillPath`,
`AbortPath`, `WidenPath`, `CloseFigure`, `FlattenPath`, `PathToRegion`,
`GetMiterLimit`, `SetMiterLimit`, `CreateEnhMetaFile`, `CloseEnhMetaFile`,
`DeleteEnhMetaFile`, `PlayEnhMetaFile`, `EnumEnhMetaFile`,
`GetEnhMetaFile`, `CopyEnhMetaFile`, `GetEnhMetaFileBits`,
`SetEnhMetaFileBits`, `CreateMetaFile`, `CloseMetaFile`, `DeleteMetaFile`,
`PlayMetaFile`, `EnumMetaFile`.

### 5.4 ADVAPI32 (38 functions)

`GetUserName`, registry:
`RegOpenKey`, `RegOpenKeyEx`, `RegCreateKey`, `RegCreateKeyEx`,
`RegQueryValue`, `RegQueryValueEx`, `RegQueryInfoKey`, `RegSetValue`,
`RegSetValueEx`, `RegDeleteKey`, `RegDeleteValue`, `RegEnumKey`,
`RegEnumKeyEx`, `RegEnumValue`, `RegCloseKey`, `RegFlushKey`, `RegSaveKey`,
`RegLoadKey`, `RegUnLoadKey`, `RegReplaceKey`, `RegRestoreKey`,
`RegConnectRegistry`, `RegNotifyChangeKeyValue`, `RegSetKeySecurity`,
`RegGetKeySecurity`.

Crypto:
`CryptAcquireContext`, `CryptReleaseContext`, `CryptCreateHash`,
`CryptHashData`, `CryptGetHashParam`, `CryptDestroyHash`.

Services:
`OpenSCManager`, `CreateService`, `OpenService`, `StartService`,
`DeleteService`, `CloseServiceHandle`.

### 5.5 COMDLG32, Version, SHELL32, COMCTL32, SHLWAPI

**COMDLG32 (10):**
`GetOpenFileName`, `GetSaveFileName`, `GetFileTitle`, `ChooseColor`,
`ChooseFont`, `PrintDlg`, `PageSetupDlg`, `FindText`, `ReplaceText`,
`CommDlgExtendedError`.

**Version (6):**
`GetFileVersionInfo`, `GetFileVersionInfoSize`, `VerQueryValue`,
`VerFindFile`, `VerInstallFile`, `VerLanguageName`.

**SHELL32 (17):**
`ShellExecute`, `ShellExecuteW`, `ShellAbout`, `SHGetFileInfo`,
`SHFileOperation`, `SHBrowseForFolder`, `SHGetPathFromIDList`,
`SHGetNewLinkInfo`, `SHAppBarMessage`, `Shell_NotifyIcon`, `ExtractIcon`,
`ExtractIconEx`, `SHFreeNameMappings`, `DragAcceptFiles`, `DragFinish`,
`DragQueryFile`, `DragQueryPoint`.

**COMCTL32 (13):**
`InitCommonControls`, `InitializeFlatSB`, `UninitializeFlatSB`, and the
`FlatSB_*` scrollbar helpers, `LBItemFromPt`. This is a **limited** common
controls set; many ComCtl32 APIs are not declared here.

**SHLWAPI (6):**
`ColorRGBToHLS`, `ColorHLSToRGB`, `SHAutoComplete`, `PathIsRelative`,
`PathIsDirectory`, `PathIsDirectoryEmpty`.

### 5.6 WINMM, PSAPI, IPHLPAPI, MSIMG32

**WINMM (59):**
`PlaySound`, `sndPlaySound`, MCI (`mciExecute`, `mciSendString`,
`mciSendCommand`, `mciGetDeviceID`, `mciGetErrorString`, ...), joystick
(`joyGetPos`, `joyGetPosEx`, `joySetCapture`, `joyGetDevCaps`,
`joyGetNumDevs`, ...), timers (`timeGetTime`, `timeSetEvent`,
`timeKillEvent`, `timeBeginPeriod`, `timeEndPeriod`, `timeGetDevCaps`,
`timeGetSystemTime`), wave input/output (`waveIn*`, `waveOut*`).

**PSAPI (17):**
`EnumProcesses`, `EnumProcessModules`, `GetModuleBaseName`,
`GetModuleFileNameEx`, `GetModuleInformation`, `GetProcessMemoryInfo`,
`GetPerformanceInfo`, `EmptyWorkingSet`, `GetMappedFileName`,
`GetProcessImageFileName`, `GetDeviceDriverBaseName`,
`GetDeviceDriverFileName`, `EnumDeviceDrivers`, `QueryWorkingSet`,
`EnumPageFiles`, `GetWsChanges`, `InitializeProcessForWsWatch`.

**IPHLPAPI (5):**
`GetInterfaceInfo`, `GetNumberOfInterfaces`, `SetAdapterIPAddress`,
`AllocateAndGetTcpExTableFromStack`, `AllocateAndGetUdpExTableFromStack`.

**MSIMG32 (5):**
`AlphaBlend`, `GradientFill`, `TransparentBlt`, plus two initialization
exports.

### 5.7 COM: OLE32, OLEAUT32, MSVBVM60

**OLE32 (16):**

`CoInitialize`, `CoInitializeEx`, `CoUninitialize`, `CoGetClassObject`,
`CoCreateInstance`, `CoCreateInstanceEx`, `CoCreateInstanceExLong`,
`CoCreateGuid`, `OleInitialize`, `IIDFromString`, `CLSIDFromString`,
`CLSIDFromProgID`, `StringFromGUID2`, `CoTaskMemFree`,
`CreateStreamOnHGlobal`, `GetHGlobalFromStream`.

**OLEAUT32 (18):**

`LoadTypeLib`, `LoadTypeLibEx`, `LoadRegTypeLib`, `RegisterTypeLib`,
`UnRegisterTypeLib`, `QueryPathOfRegTypeLib`, `CreateTypeLib`,
`CreateTypeLib2`, `SysAllocString`, `SysAllocStringLen`,
`SysAllocStringByteLen`, `VarCmp`, `SystemTimeToVariantTime`,
`CreateErrorInfo`, `CreateStdDispatch`, `CreateDispTypeInfo`,
`LHashValOfNameSys`, `LHashValOfNameSysA`.

**MSVBVM60 (36):** VB runtime helpers, including `ArrPtr` (used by old code to
get a SAFEARRAY pointer).

### 5.8 WS2_32 (38 functions, present but secondary here)

`gethostbyaddr`, `gethostbyname`, `gethostname`, `getservbyname`,
`getprotobynumber`, `getprotobyname`, `getservbyport`, `inet_addr`,
`inet_ntoa`, `htons`, `htonl`, `ntohl`, `ntohs`, `socket`, `closesocket`,
`connect`, `getsockname`, `getpeername`, `bind`, `selects`, `recv`, `send`,
`shutdown`, `listen`, `accept`, `setsockopt`, `getsockopt`, `sendto`,
`recvfrom`, `ioctlsocket`, `WSAIoctl`, `WSAStartup`, `WSACleanup`,
`WSAAsyncSelect`, `WSAGetLastError`, `WSAAsyncGetHostByAddr`,
`WSAAsyncGetHostByName`, `WSACancelBlockingCall`.

Note: `WSACancelAsyncRequest` is **not** declared (see §7).

### 5.9 Other modules

`GDIPLUS` (79), `FMOD` (227), `LUA` (145), `DESDLL` (6), `ASM` (5),
`LIBMAD` (12), `LAME_ENC` (5), `DINPUT8` (2), `IMAGEHLP` (1), `ZLIB` (9),
`UNRAR` (7), `IJL11` (7), `AVICAP32` (2), `CARDS` (5), `D3d9` (2),
`NTDLL` (1), `OLEPRO32` (8), and GUID/constant-only modules.

Use the inspector (§9) to get the complete function list for a module; this
section is a curated index, not an exhaustive dump.

---

## 6. Enums and constants present in the tlb

The tlb has 271 enums. The following are the ones verified in detail.

### 6.1 ProtocolConstants

```
sckTCPProtocol = 0
sckUDPProtocol = 1
```

### 6.2 StateConstants

```
sckClosed            = 0
sckOpen              = 1
sckListening         = 2
sckConnectionPending = 3
sckResolvingHost     = 4
sckHostResolved      = 5
sckConnecting        = 6
sckConnected         = 7
sckClosing           = 8
sckError             = 9
```

### 6.3 ErrorConstants (37 values, exact)

```
sckInvalidPropertyValue = 380
sckGetNotSupported      = 394
sckSetNotSupported      = 383
sckOutOfMemory          = 7
sckBadState             = 40006
sckInvalidArg           = 40014
sckSuccess              = 40017     <-- unusual; not 40000
sckUnsupported          = 40018
sckInvalidOp            = 40020
sckOutOfRange           = 40021     <-- unusual; not 40022
sckWrongProtocol        = 40026
sckOpCanceled           = 10004
sckInvalidArgument      = 10014
sckWouldBlock           = 10035
sckInProgress           = 10036
sckAlreadyComplete      = 10037
sckNotSocket            = 10038
sckMsgTooBig            = 10040
sckPortNotSupported     = 10043
sckAddressInUse         = 10048
sckAddressNotAvailable  = 10049
sckNetworkSubsystemFailed = 10050
sckNetworkUnreachable   = 10051
sckNetReset             = 10052
sckConnectAborted       = 10053
sckConnectionReset      = 10054
sckNoBufferSpace        = 10055
sckAlreadyConnected     = 10056
sckNotConnected         = 10057
sckSocketShutdown       = 10058
sckTimedout             = 10060
sckConnectionRefused    = 10061
sckNotInitialized       = 10093
sckHostNotFound         = 11001
sckHostNotFoundTryAgain = 11002
sckNonRecoverableError  = 11003
sckNoData               = 11004
```

Use these constants instead of guessing the `sck*` numbers.

### 6.4 Address family — note the typo

The enum is named **`vbADDRESS_FAMILLY_CONSTANTS`** (double **L** in FAMILY).
Key values:

```
AF_UNSPEC = 0
AF_UNIX   = 1
AF_INET   = 2
AF_INET6  = 23
AF_MAX    = 32
```

It also contains `AF_IMPLINK`, `AF_PUP`, `AF_CHAOS`, `AF_NS`, `AF_IPX = 0`,
`AF_ISO`, `AF_OSI = 0`, `AF_ECMA`, `AF_DATAKIT`, `AF_CCITT`, `AF_SNA`,
`AF_DECnet`, `AF_DLI`, `AF_LAT`, `AF_HYLINK`, `AF_APPLETALK`, `AF_NETBIOS`,
`AF_VOICEVIEW`, `AF_FIREFOX`, `AF_UNKNOWN1`, `AF_BAN`, `AF_ATM`, `AF_CLUSTER`,
`AF_12844`, `AF_IRDA`, `AF_NETDES`, `AF_TCNPROCESS`, `AF_TCNMESSAGE`,
`AF_ICLFXBM`.

### 6.5 Socket type/protocol constants

```
vbSOCKET_TYPE_CONSTANTS:
  SOCK_STREAM=1, SOCK_DGRAM=2, SOCK_RAW=3, SOCK_RDM=4, SOCK_SEQPACKET=5

vbSOCKET_PROTOCOL_CONSTANTS:
  IPPROTO_IP=0, IPPROTO_ICMP=1, IPPROTO_IGMP=2, IPPROTO_GGP=3,
  IPPROTO_TCP=6, IPPROTO_PUP=12, IPPROTO_UDP=17, IPPROTO_IDP=22,
  IPPROTO_ND=77, IPPROTO_RAW=255, IPPROTO_MAX=256
```

### 6.6 Window constants

`GWL_CONSTANTS` contains:

```
GWL_WNDPROC=-4, GWL_HINSTANCE=-6, GWL_HWNDPARENT=-8, GWL_ID=-12,
GWL_STYLE=-16, GWL_EXSTYLE=-20, GWL_USERDATA=-21,
DWL_MSGRESULT=0, DWL_DLGPROC=4, DWL_USER=8
```

Other constants may be in the `Constants` module or other enums. Use the
inspector if a name is not listed here.

---

## 7. Missing constants and APIs — workarounds

These names were **not found** in the tlb. Define them privately in the project
or add private `Declare` aliases.

### 7.1 Common Winsock constants not in the tlb

```vb
' FD_* event masks
FD_READ    = &H1&
FD_WRITE   = &H2&
FD_OOB     = &H4&
FD_ACCEPT  = &H8&
FD_CONNECT = &H10&
FD_CLOSE   = &H20&

' shutdown()
SD_RECEIVE = 0
SD_SEND    = 1
SD_BOTH    = 2

' addresses/sockets
INADDR_ANY     = 0&
INADDR_NONE    = -1&
INVALID_SOCKET = -1&
SOCKET_ERROR   = -1&

' socket options
SOL_SOCKET     = &HFFFF&
SO_DEBUG       = &H1&
SO_ACCEPTCONN  = &H2&
SO_REUSEADDR   = &H4&
SO_KEEPALIVE   = &H8&
SO_DONTROUTE   = &H10&
SO_BROADCAST   = &H20&
SO_LINGER      = &H80&
SO_OOBINLINE   = &H100&
SO_SNDBUF      = &H1001&
SO_RCVBUF      = &H1002&
SO_ERROR       = &H1007&
SO_TYPE        = &H1008&

' window
HWND_MESSAGE = -3&
```

Common WSA error codes (also not in the tlb):

```vb
WSAEINTR=10004, WSAEBADF=10009, WSAEACCES=10013, WSAEFAULT=10014, WSAEINVAL=10022,
WSAEMFILE=10024, WSAEWOULDBLOCK=10035, WSAEINPROGRESS=10036,
WSAEALREADY=10037, WSAENOTSOCK=10038, WSAEDESTADDRREQ=10039,
WSAEMSGSIZE=10040, WSAEPROTOTYPE=10041, WSAENOPROTOOPT=10042,
WSAEPROTONOSUPPORT=10043, WSAESOCKTNOSUPPORT=10044,
WSAEOPNOTSUPP=10045, WSAEPFNOSUPPORT=10046, WSAEAFNOSUPPORT=10047,
WSAEADDRINUSE=10048, WSAEADDRNOTAVAIL=10049, WSAENETDOWN=10050,
WSAENETUNREACH=10051, WSAENETRESET=10052, WSAECONNABORTED=10053,
WSAECONNRESET=10054, WSAENOBUFS=10055, WSAEISCONN=10056,
WSAENOTCONN=10057, WSAESHUTDOWN=10058, WSAETOOMANYREFS=10059,
WSAETIMEDOUT=10060, WSAECONNREFUSED=10061, WSAELOOP=10062,
WSAENAMETOOLONG=10063, WSAEHOSTDOWN=10064, WSAEHOSTUNREACH=10065,
WSASYSNOTREADY=10091, WSAVERNOTSUPPORTED=10092,
WSANOTINITIALISED=10093, WSAEDISCON=10101,
WSAHOST_NOT_FOUND=11001, WSATRY_AGAIN=11002,
WSANO_RECOVERY=11003, WSANO_DATA=11004
```

### 7.2 Missing APIs

```vb
Private Declare Function CancelAsyncRequest Lib "ws2_32.dll" _
    Alias "WSACancelAsyncRequest" (ByVal hAsyncTaskHandle As Long) As Long

' gethostname itself IS in the tlb; this alias is optional and avoids
' ANSI/Unicode buffer surprises when using a Byte array:
Private Declare Function gethostnameBytes Lib "ws2_32.dll" _
    Alias "gethostname" (ByRef name As Byte, ByVal namelen As Long) As Long
```

Other verified gaps:

- `CoTaskMemAlloc` is **not** declared; `CoTaskMemFree` is (OLE32). Use
  `HeapAlloc`/`HeapFree`, or add a `Declare` for `CoTaskMemAlloc`.
- `SysFreeString` is **not** declared, although `SysAllocString` is (OLEAUT32).
- `SafeArrayCreate` / `SafeArrayDestroy` are not declared.
- `ISupportErrorInfo` is not a type in the tlb.
- `WSACancelAsyncRequest` is not declared.
- `FD_*`, WSA errors, `SO_*`, `SD_*`, `INADDR_*`, `INVALID_SOCKET`,
  `SOCKET_ERROR`, and `HWND_MESSAGE` are not declared.

Before adding a declaration, first search the tlb with the inspector; many
unexpected APIs are present (e.g., `selects` rather than `select`).

---

## 8. Type/record notes

- `CLSID`: `Data1 As Long`, `Data2 As Integer`, `Data3 As Integer`,
  `Data4(0 To 7) As Byte`. Use `CLSIDFromString` (OLE32) or
  `IIDFromString`/`CLSIDFromProgID`.
- `SOCKADDR`: `sin_family As Integer`, `sin_port As Integer` (network byte
  order), `sin_addr As Long`; `Len()` is 16. Use `htons`/`ntohs` and
  unsigned conversions for ports.
- `HOSTENT`: `hName As Long`, `hAliases As Long`, `hAddrType As Integer`,
  `hLength As Integer`, `hAddrList As Long`. `hAddrList` points to an array of
  address pointers; copy 4 bytes at `ByVal hAddrList` to get the first address.
- `WSADATA`: contains startup data including `iMaxUdpDg`; use it if you need the
  maximum UDP datagram size.
- `SAFEARRAY1D`: used by `ArraySize`-style helpers (`cDims`,
  `Bounds.cElements`).
- Interfaces: `IUnknown`, `IDispatch`, `ITypeInfo`, `ITypeLib`,
  `IProvideClassInfo`, `IConnectionPoint`, `IConnectionPointContainer`,
  `IEnumConnections`, `IEnumConnectionPoints`, `IWinHttpRequest`, and many
  DirectX/Sound/GDI+ interfaces.
- `IWinHttpRequest`/`WinHttpRequest` and `IWinHttpRequestEvents` are present if
  an HTTP client is needed.
- Standard interface IIDs in this tlb:

| Interface | IID |
| --- | --- |
| IUnknown | `{00000000-0000-0000-C000-000000000046}` |
| IDispatch | `{00020400-0000-0000-C000-000000000046}` |
| ITypeInfo | `{00020401-0000-0000-C000-000000000146}` |
| ITypeLib | `{00020402-0000-0000-C000-000000000046}` |
| IProvideClassInfo | `{B196B283-BAB4-101A-B69C-00AA00341D07}` |
| IConnectionPointContainer | `{B196B284-BAB4-101A-B69C-00AA00341D07}` |
| IEnumConnectionPoints | `{B196B285-BAB4-101A-B69C-00AA00341D07}` |
| IConnectionPoint | `{B196B286-BAB4-101A-B69C-00AA00341D07}` |
| IEnumConnections | `{B196B287-BAB4-101A-B69C-00AA00341D07}` |
| IWinHttpRequest | `{016FE2EC-B2C8-45F8-B23B-39E53A75396B}` |
| IWinHttpRequestEvents | `{F97F4E15-B787-4212-80D1-D380CBBF982E}` |

---

## 9. How to inspect the tlb quickly

### 9.1 32-bit vs 64-bit `oVft` (critical)

`ITypeInfo.GetFuncDesc` returns `oVft` in bytes. The value is **scaled by the
calling process bitness** by `oleaut32`:

- 32-bit process: 4-byte vtable slots.
  - IUnknown: QI=0, AddRef=4, Release=8.
  - IDispatch: 12, 16, 20, 24.
  - IMSWinsockControl: AboutBox=28, Protocol get=32, put=36, ... RemoteHost put=124.
  - `cbSizeVft` = 128.
- 64-bit process: 8-byte vtable slots.
  - IUnknown: 0, 8, 16.
  - IDispatch: 24, 32, 40, 48.
  - IMSWinsockControl: AboutBox=56, ... RemoteHost put=248.
  - `cbSizeVft` = 256.

So a 64-bit dump showing `AboutBox oVft=56` is **not** a malformed type library;
it is the 64-bit view. Run the inspector in 32-bit PowerShell for real 4-byte
offsets:

```
%WINDIR%\SysWOW64\WindowsPowerShell\v1.0\powershell.exe -NoProfile -File dump.ps1
```

### 9.2 Copy-paste inspector (C# via PowerShell)

```powershell
$cs = @'
using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using CT = System.Runtime.InteropServices.ComTypes;

[StructLayout(LayoutKind.Sequential)]
public struct LibAttr {
    public Guid guid;
    public int lcid;
    public int syskind;
    public ushort major;
    public ushort minor;
    public ushort flags;
}

public class TlbDump {
    [DllImport("oleaut32.dll", CharSet=CharSet.Unicode, PreserveSig=false)]
    public static extern void LoadTypeLibEx(string file, int regkind, out ITypeLib tlb);

    public static void Dump(string file, string typeName) {
        ITypeLib tlb; LoadTypeLibEx(file, 0, out tlb);

        string n, d, h; int c;
        tlb.GetDocumentation(-1, out n, out d, out c, out h);

        IntPtr lp; tlb.GetLibAttr(out lp);
        LibAttr a = (LibAttr)Marshal.PtrToStructure(lp, typeof(LibAttr));
        tlb.ReleaseTLibAttr(lp);

        Console.WriteLine(n + " " + a.guid +
            " lcid=" + a.lcid + " syskind=" + a.syskind +
            " ver=" + a.major + "." + a.minor +
            " types=" + tlb.GetTypeInfoCount());

        for (int i = 0; i < tlb.GetTypeInfoCount(); i++) {
            ITypeInfo ti; tlb.GetTypeInfo(i, out ti);
            string tn, td, th; int tc;
            ti.GetDocumentation(-1, out tn, out td, out tc, out th);

            IntPtr pa; ti.GetTypeAttr(out pa);
            CT.TYPEATTR ta = (CT.TYPEATTR)Marshal.PtrToStructure(pa, typeof(CT.TYPEATTR));

            if (typeName == null || tn == typeName) {
                Console.WriteLine(ta.typekind + " " + tn +
                    " funcs=" + ta.cFuncs + " vars=" + ta.cVars);

                for (int f = 0; f < ta.cFuncs; f++) {
                    IntPtr pf; ti.GetFuncDesc(f, out pf);
                    CT.FUNCDESC fd = (CT.FUNCDESC)Marshal.PtrToStructure(pf, typeof(CT.FUNCDESC));
                    string[] names = new string[fd.cParams + 1];
                    int pc;
                    ti.GetNames(fd.memid, names, names.Length, out pc);
                    Console.WriteLine("  [" + f + "] oVft=" + fd.oVft +
                        " slot=" + (fd.oVft / 4) +
                        " " + string.Join(",", names));
                    ti.ReleaseFuncDesc(pf);
                }
            }
            ti.ReleaseTypeAttr(pa);
        }
    }
}
'@

Add-Type -TypeDefinition $cs
[TlbDump]::Dump("<path>\Win32API.tlb", "IMSWinsockControl")
```

Pass `$null` instead of `"IMSWinsockControl"` to list every type. Use
`IMSWinsockControl`, `DMSWinsockControlEvents`, `Winsock`,
`IConnectionPoint`, `ITypeInfo`, `ITypeLib`, or any module name such as
`KERNEL32`, `USER32`, `GDI32`, `ADVAPI32`, `OLE32`, `OLEAUT32` to inspect it.

For enum values, use `ITypeInfo.GetVarDesc`; `VARDESC.desc.lpvarValue` points to
a `VARIANT`. For `VT_I4`, read `Int32` at offset **8** of that VARIANT. This is
how the `sck*`, `AF_*`, `SOCK_*`, `IPPROTO_*`, and `GWL_*` values in this
document were obtained.

---

## 10. Gotchas for using this type library

1. **Huge namespace.** 677 types and 41 modules. Public constants from the tlb
   can collide with project declarations. Use private, prefixed constants for
   missing names.
2. **`AF_INET` is in `vbADDRESS_FAMILLY_CONSTANTS`** (typo double L).
   `GWL_WNDPROC` is in `GWL_CONSTANTS`.
3. **Not every Windows API/constant is present.** `FD_*`, WSA errors, `SO_*`,
   `SD_*`, `INADDR_*`, `INVALID_SOCKET`, `SOCKET_ERROR`, `HWND_MESSAGE`,
   `WSACancelAsyncRequest`, `CoTaskMemAlloc`, `SysFreeString`,
   `SafeArrayCreate/Destroy`, and `ISupportErrorInfo` are missing. Define/alias
   them.
4. **`New Winsock` is not provided by the tlb.** It uses the MSWINSCK CLSID and
   requires the OCX. The tlb is declaration-only.
5. **A 64-bit `oVft` is scaled.** Inspect in 32-bit for real slot offsets.
6. **32-bit registry view.** If you use the MSWINSCK OCX, its registration is
   under `WOW6432Node`; 64-bit registry queries miss it.
7. **The tlb is compile-time for most uses.** A compiled VB6 EXE does not embed
   it. If you use `LoadTypeLib` for IDispatch/typeinfo, deploy the tlb (or use a
   registered typelib).
8. **LCID.** Reference the tlb with `40c` (1036).
9. **The tlb is third-party (EBArtSoft).** Verify licensing/redistribution before
   shipping it.
10. **Do not edit the binary.** Regenerate from its source if a change is needed.

---

## 11. Appendix: MSWINSCK-compatible declarations included in the tlb

These are part of the type library (not a separate OCX in the tlb). They are
documented here because another project may encounter them.

### 11.1 IMSWinsockControl

| Item | Value |
| --- | --- |
| Type | `TKIND_INTERFACE` |
| IID | `{248DD892-BB45-11CF-9ABC-0080C7E7B78D}` |
| Functions | 25 |
| Vtable size | 128 bytes / 32 slots in 32-bit |

Vtable map (32-bit slots; IUnknown 0–2, IDispatch 3–6):

| Slot | Member | Kind |
| ---: | --- | --- |
| 7 | `AboutBox` | method |
| 8 | `Protocol` | get `Long` |
| 9 | `Protocol` | put `Long` |
| 10 | `RemoteHostIP` | get `String` |
| 11 | `LocalHostName` | get `String` |
| 12 | `LocalIP` | get `String` |
| 13 | `SocketHandle` | get `Long` |
| 14 | `_RemoteHost` | hidden get `String` |
| 15 | `_RemoteHost` | hidden put `String` |
| 16 | `RemotePort` | get `Long` |
| 17 | `RemotePort` | put `Long` |
| 18 | `LocalPort` | get `Long` |
| 19 | `LocalPort` | put `Long` |
| 20 | `state` | get `Short`/`Integer` |
| 21 | `BytesReceived` | get `Long` |
| 22 | `connect` | `[RemoteHost], [RemotePort]` |
| 23 | `listen` | — |
| 24 | `accept` | `requestID As Long` |
| 25 | `SendData` | `Data As Variant` |
| 26 | `GetData` | `Data, [Type], [maxLen]` |
| 27 | `PeekData` | `Data, [Type], [maxLen]` |
| 28 | `Close` | — |
| 29 | `bind` | `[LocalPort], [LocalIP]` |
| 30 | `RemoteHost` | get `String` |
| 31 | `RemoteHost` | put `String` |

Property getters are HRESULT methods with a retval parameter, not IDispatch
`propget` signatures.

### 11.2 DMSWinsockControlEvents

| Item | Value |
| --- | --- |
| Type | `TKIND_DISPATCH` (dispinterface) |
| IID | `{248DD893-BB45-11CF-9ABC-0080C7E7B78D}` |
| Functions | 7 |

| Dispid | Event | Signature |
| ---: | --- | --- |
| 0 | `DataArrival` | `(bytesTotal As Long)` |
| 1 | `connect` | `()` |
| 2 | `ConnectionRequest` | `(requestID As Long)` |
| 3 | `SendProgress` | `(bytesSent As Long, bytesRemaining As Long)` |
| 4 | `SendComplete` | `()` |
| 5 | `Close` | `()` |
| 6 | `error` / `Error` | `(Number As Integer, Description As String, Scode As Long, Source As String, HelpFile As String, HelpContext As Long, CancelDisplay As Boolean)` |

### 11.3 Winsock coclass

| Item | Value |
| --- | --- |
| Type | `TKIND_COCLASS`, `FCONTROL` |
| CLSID | `{248DD896-BB45-11CF-9ABC-0080C7E7B78D}` |
| Default interface | `IMSWinsockControl` |
| Source interface | `DMSWinsockControlEvents` |

---

## 12. Quick checklist for a new project

1. Add the `Reference=` line with the correct relative path and `40c`.
2. Use VB6's Object Browser or the §9 inspector to confirm a function/type is
   declared and get its exact signature.
3. Define missing constants privately (`FD_*`, WSA errors, `SO_*`, `SD_*`,
   `INADDR_*`, `INVALID_SOCKET`, `SOCKET_ERROR`, `HWND_MESSAGE`).
4. Add private aliases for missing APIs (`WSACancelAsyncRequest`,
   `CoTaskMemAlloc`, `SysFreeString`, SafeArray helpers, if needed).
5. Do not assume the tlb is Win32-only; check the module list before searching
   for a name.
6. For the MSWINSCK-compatible types, use the exact IIDs/vtable in §11; otherwise
   ignore that appendix.
7. If you load the type library at runtime for type info, deploy the tlb file;
   otherwise it is compile-time only.
8. Verify the EBArtSoft tlb's licensing/redistribution before shipping it.
