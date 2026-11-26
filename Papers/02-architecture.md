# 02 — Architecture

## The three artefacts

```
                 +------------------------------------------------------+
 host EXE  --->  | winshim.dll        (C, TCC, ~11 KB)                  |
 (VB6)           |  - DllGetClassObject for {248DD896-...}              |
                 |  - IClassFactory / IClassFactory2                    |
                 |  - OLE control layer around the VB object            |
                 +---------------------------+--------------------------+
                                             | LoadLibrary + GetProcAddress
                                             v
                 +------------------------------------------------------+
                 | WinsockVB6.dll     (VB6 ActiveX DLL, ~32 KB)         |
                 |  - Common.bas: export New_WinsockVB6()               |
                 |  - modWinsock2: IMSWinsockControl, ICPC/ICP, events  |
                 +------------------------------------------------------+
```

`TypeLib\MSWinsockLib.tlb` is the extracted OCX type library (same LIBID,
interfaces, dispids). It is the authority for the object model and the source of
the dispids used by the dispatch layer. `MSWinsockLib.idl` is the decompiled
source used to regenerate/verify it.

## Three call paths

**A. Raw DLL export (no COM).** `IRL-Test`'s `USE_VBWINSOCK = True` path:

```vb
Declare Function New_WinsockVB6 Lib "WinsockVB6.dll" () As Winsock
Set mSock = New_WinsockVB6()
```

`New_WinsockVB6` is **not** a COM-style `HRESULT f(IUnknown **)`. It takes no
arguments and returns the interface pointer in `EAX` (`void *__stdcall f(void)`).
No registry, no manifest, no shim involved.

**B. Control created from code.** `USE_MSWINSOCK = True`, or `New Winsock` /
`CreateObject("MSWinsock.Winsock")` with the manifest/registration in place:

```
CoCreateInstance({248DD896}) -> winshim.dll -> New_WinsockVB6()
host QI(IUnknown, IMSWinsockControl, IPersistStreamInit, IPersistPropertyBag, ICPC)
```

The host binds to `IMSWinsockControl` and calls methods through the vtable. This
path never calls `IOleObject::SetClientSite`.

**C. Control on a VB6 form.**

```
QI IQuickActivate      (refused: slow path)
QI IOleObject          -> SetClientSite, extents, DoVerb
QI IOleControl         -> FreezeEvents, ambient changes
QI IPersistStreamInit  -> InitNew (only after the site exists)
QI IViewObject         -> SetAdvise
QI IConnectionPointContainer -> the inner object's, for events
QI IViewObjectEx       (refused)
QI IPointerInactive    (refused)
QI IDispatch           -> forwarded to the VB object (dispids)
```

Properties are called through `IDispatch` with MSWinsockLib dispids; events flow
from the VB object's connection point to the form's sink.

## `modWinsock2.bas` — the replacement

* **No registration.** The object is created by `New_Winsock2()` and typed as
  `Winsock` from `Win32API.tlb`; VB sees a normal object reference. A conditional
  area (`#If NO_TYPELIB`) is reserved for local Win32 declarations so the module
  can eventually build without that type library; it is empty and the work is
  parked (see `06-state-of-development.md`, P1).
* **Fabricated COM.** `W2_BuildVtables` fills `gV1(0..31)` (IMSWinsockControl:
  IUnknown + IDispatch + 25 members), `gV2` (IConnectionPointContainer), `gV3`
  (IConnectionPoint) and `gV4` (IEnumConnections, not exposed) with `AddressOf`
  pointers. `W2_AllocHeaders` heap-allocates a header whose first field points at
  `gV1`, plus the other interface headers, and returns that pointer.
* **State pool.** `gStates()` is a growable array of fixed slots with a free list
  (`InUse`, `NextFree`, `RefCount`, `HeaderPtr`). Each object holds a
  `StateIndex`; `pThis.StateIndex` is read directly from the fabricated header.
* **Lifetime.** The module keeps a locked object alive (`StartVBRuntime`,
  `mLockedObject`) so the VB runtime is not unloaded while COM still holds
  objects. Object references are counted per slot; the module state outlives
  individual objects.
* **Sockets.** One hidden message window with `WSAAsyncSelect`; per-state socket
  handles, read/write buffers with compaction, UDP data queue, pending-accept
  collection keyed by a request id string, DNS map for asynchronous resolution.
* **Events.** `gStates(idx).Sinks` holds `DMSWinsockControlEvents` references.
  `W2_RaiseSimple` copies the sink list, AddRefs the state, and invokes each sink
  by value (see `05-discoveries.md` for why the arguments are parenthesised).
* **Errors.** `Err.Raise sckXxx, "Winsock.X", "text"` inside typed members. The
  Error event carries `Scode = &H800A0000 Or Number`, matching the OCX.
* **Diagnostics.** `W2_Diagnostics()` reports live objects, used slots, sockets,
  pending accepts, datagrams and DNS requests. It is gated behind
  `#If W2DIAG`; only `FeatureTests.vbp` defines that constant. (`DEBUG` cannot be
  used: it collides with the reserved `Debug` object.)

### `IMSWinsockControl` vtable map (slots in `gV1`)

| Slot | Member | Slot | Member |
| ---: | --- | ---: | --- |
| 0–2 | IUnknown | 17 | `LetRemotePort(long)` |
| 3 | `GetTypeInfoCount` | 18 | `GetLocalPort(long*)` |
| 4 | `GetTypeInfo` | 19 | `LetLocalPort(long)` |
| 5 | `GetIDsOfNames` | 20 | `GetState(short*)` |
| 6 | `Invoke` | 21 | `GetBytesReceived(long*)` |
| 7 | `AboutBox` | 22 | `Connect(VARIANT, VARIANT)` |
| 8 | `GetProtocol(long*)` | 23 | `Listen` |
| 9 | `LetProtocol(long)` | 24 | `Accept(long)` |
| 10 | `GetRemoteHostIP(BSTR*)` | 25 | `SendData(VARIANT)` |
| 11 | `GetLocalHostName(BSTR*)` | 26 | `GetData(VARIANT*, VARIANT, VARIANT)` |
| 12 | `GetLocalIP(BSTR*)` | 27 | `PeekData(VARIANT*, VARIANT, VARIANT)` |
| 13 | `GetSocketHandle(long*)` | 28 | `Close` |
| 14 | `Get_RemoteHost(BSTR*)` | 29 | `Bind(VARIANT, VARIANT)` |
| 15 | `Let_RemoteHost(BSTR)` | 30 | `GetRemoteHost(BSTR*)` |
| 16 | `GetRemotePort(long*)` | 31 | `LetRemoteHost(BSTR)` |

### Dispid map (MSWinsockLib, used by `W2_Invoke`)

| Dispid | Member | Dispid | Member |
| ---: | --- | ---: | --- |
| 0 | `_RemoteHost` (get/put) | 9 | `BytesReceived` (get) |
| 1 | `RemotePort` (get/put) | 10 | `RemoteHost` (get/put) |
| 2 | `LocalPort` (get/put) | 64 | `Connect` |
| 3 | `Protocol` (get/put) | 65 | `Listen` |
| 4 | `RemoteHostIP` (get) | 66 | `Accept` |
| 5 | `LocalHostName` (get) | 67 | `SendData` |
| 6 | `LocalIP` (get) | 68 | `GetData` |
| 7 | `SocketHandle` (get) | 69 | `PeekData` |
| 8 | `State` (get) | 70 | `Close` |
| | | 71 | `Bind` |
| | | −552 | `AboutBox` |

`W2_GetIDsOfNames` maps the same names (case-insensitive) for `CallByName` and
other late-bound callers. `GetTypeInfoCount` still returns 0 and `GetTypeInfo`
returns `DISP_E_BADINDEX`: the object has automation *dispatch* but no type
information of its own.

### Event dispids (`DMSWinsockControlEvents`, IID `248DD893`)

| Dispid | Event | Arguments |
| ---: | --- | --- |
| 0 | `DataArrival` | `bytesTotal As Long` |
| 1 | `Connect` | — |
| 2 | `ConnectionRequest` | `requestID As Long` |
| 3 | `SendProgress` | `bytesSent As Long, bytesRemaining As Long` |
| 4 | `SendComplete` | — |
| 5 | `Close` | — |
| 6 | `Error` | `Number As Integer, Description As String, Scode As Long, Source As String, HelpFile As String, HelpContext As Long, CancelDisplay As Boolean` |

## `AsDll\WinsockVB6.dll` — the VB6 wrapper

* Project type `OleDll`, one dummy class (unused), `modWinsock2`, `Common.bas` (renamed from `modDllMain.bas` in `7705b45`).
* `New_WinsockVB6()` calls `StartVBRuntime` (locks the runtime by creating an
  object through a patched `DllGetClassObject`) and returns `New_Winsock2()`.
* The export is produced by `[VBCompiler] LinkSwitches=/EXPORT:New_WinsockVB6`.
* The calling convention is raw: pointer in `EAX`, no arguments. Verified with
  `winshim\probe.exe`: the COM-style call returns a garbage HRESULT and a null
  object; the raw call returns a valid object.

## `winshim.dll` — identity and OLE

* **Factory.** Static `IClassFactory`/`IClassFactory2`; `DllGetClassObject`
  accepts only the MSWINSCK CLSID. `IClassFactory2` answers `GetLicInfo`
  (runtime key available, license verified), `RequestLicKey` (a dummy BSTR) and
  `CreateInstanceLic` (any key), because VB6 asks for it and creates the control
  through `CreateInstanceLic`.
* **Wrapper (`Ctl`).** Holds the inner `IUnknown*`, the client site, advise
  sinks, extent, and one embedded sub-object per interface (`IOleObject`,
  `IViewObject`, `IPersistStreamInit`, `IPersistPropertyBag`, `IOleControl`).
* **QueryInterface policy.**

  | IID | Result |
  | --- | --- |
  | `IUnknown` | the wrapper (`&c->ole`), preserving identity |
  | `IOleObject` | the wrapper's |
  | `IViewObject` / `IViewObject2` | the wrapper's |
  | `IOleControl` | the wrapper's |
  | `IPersistStreamInit` | the wrapper's **only if a client site exists** |
  | `IPersistPropertyBag` | the wrapper's **only if a client site exists** |
  | `IMSWinsockControl`, `IDispatch`, `IConnectionPointContainer`, everything else | the inner object's pointer |
  | `IQuickActivate`, `IViewObjectEx`, `IPointerInactive` | `E_NOINTERFACE` |

* **Persistence.** `GetClassID` returns the MSWINSCK CLSID; `InitNew`/`Load`/
  `Save` are accepted, `IsDirty` returns `S_FALSE`, `GetSizeMax` returns 0.
  Property values are **not** replayed (see the design-time finding).
* **Dispatch is not in C.** `IDispatch` is answered by the VB object. A VB error
  raised inside a method only becomes a normal HRESULT when the call is made from
  VB; through a C frame the raise escapes as `RPC_E_DISCONNECTED (0x80010108)`.
* **Logging.** Every `QueryInterface` and every OLE/persistence call is appended
  to `winshim.log` next to the host EXE. The log is the primary instrument for
  diagnosing "did the shim serve this process?". It is compiled out by defining
  `WINSHIM_NO_LOG` (VS2019 Release); Debug and TCC builds log.

## CLSIDs and IIDs

| Name | GUID |
| --- | --- |
| Type library / `MSWinsockLib` | `{248DD890-BB45-11CF-9ABC-0080C7E7B78D}` |
| `IMSWinsockControl` | `{248DD892-BB45-11CF-9ABC-0080C7E7B78D}` |
| `DMSWinsockControlEvents` | `{248DD893-BB45-11CF-9ABC-0080C7E7B78D}` |
| Winsock coclass | `{248DD896-BB45-11CF-9ABC-0080C7E7B78D}` |
| `IQuickActivate` | `{CF51ED10-62FE-11CF-BF86-00A0C9034836}` |
| `IViewObjectEx` | `{3AF24292-0C96-11CE-A0CF-00AA00600AB8}` |
| `IPointerInactive` | `{55980BA0-35AA-11CF-B671-00AA004CD6D8}` |
| `IPersistPropertyBag` | `{37D84F60-42CB-11CE-8135-00AA004BB851}` |
| `IClassFactory2` | `{B196B28F-BAB4-101A-B69C-00AA00341D07}` |
