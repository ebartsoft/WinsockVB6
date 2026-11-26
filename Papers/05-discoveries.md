# 05 — Discoveries

Findings that were not obvious from the source and that shaped the design. Each
entry states the observation, the evidence, and what was done about it.

---

## 1. `New_WinsockVB6` is not a COM-style export

**Observation.** The export returns a pointer in `EAX` and takes no arguments:
`void *__stdcall f(void)`. Calling it as `HRESULT f(IUnknown **)` yields a garbage
HRESULT and a null object.

**Evidence.** `winshim\probe.exe` calls it both ways:
`com-abi: hr=0x0076C608 obj=00000000` vs `raw: obj=006BA650`, and the raw object
answers `QI IMSWinsockControl hr=0`, `QI IConnectionPointContainer hr=0`.

**Consequence.** The shim calls the raw form. This is also why
`New_WinsockVB6() As Winsock` works from VB6: the runtime uses the raw
convention.

---

## 2. The activation sequence VB6 uses for a form-hosted control

**Observation (measured, not assumed).**

```
CreateInstance
QI IUnknown
QI IQuickActivate            -> E_NOINTERFACE (accepted; slow path)
QI IOleObject                -> SetClientSite
QI IOleControl               -> FreezeEvents(TRUE)
QI IPersistStreamInit        -> InitNew
QI IViewObject               -> SetAdvise
QI IConnectionPointContainer (from the inner object)
QI IViewObjectEx             -> refused
QI IPointerInactive          -> refused
QI IDispatch
```

**Consequence.** `IOleObject` and `IViewObject` are mandatory; `IQuickActivate`,
`IViewObjectEx` and `IPointerInactive` may be refused without preventing the
control from loading. `IViewObjectEx` *is* used when present (VB6 called
`GetRect`), but implementing it as stubs crashed the form, so it was reverted.

---

## 3. Licensing: VB6 asks for `IClassFactory2`

**Observation.** The class factory receives
`QI {B196B28F-BAB4-101A-B69C-00AA00341D07}` (IClassFactory2) and then creates the
control through `CreateInstanceLic`, i.e. it treats the class as a licensed
control.

**Consequence.** `winshim.dll` implements `IClassFactory2`: `GetLicInfo` reports
a runtime key and a verified license, `RequestLicKey` returns a dummy BSTR, and
`CreateInstanceLic` accepts any key. It is used on every start.

---

## 4. Design-time property values are never handed to the control

**Observation.** A `.frm` with `Protocol`, `RemoteHost` and `RemotePort` compiles
those values into the EXE (a saved record with the BSTR `127.0.0.1`, the long
`65000` and interface GUID `248DD892` is present in the binary), and the real OCX
restores them at run time. Our object receives only
`IPersistStreamInit::InitNew`: no `Load`, and no `propput` through `IDispatch`.

**Ruled out by measurement:**

* manifest/SxS vs plain registry activation (temporary per-user `InprocServer32`);
* typelib registration (per-user, identical TLB);
* type information — VB6 never calls `GetTypeInfoCount`/`GetTypeInfo` and never
  QIs `IProvideClassInfo`, even when available;
* licensing (`IClassFactory2` present, no change);
* persistence negotiation — refusing `IPersistStreamInit` makes the form fail
  with error 372; `InitNew` must return `S_OK`; `S_FALSE` changes nothing;
* missing control interfaces — `IViewObjectEx`/`IPointerInactive` stubs, then
  reverted (form died before `Form_Load`);
* a logging proxy around the real OCX — not viable: the wrapper and the OCX do
  not share COM identity (QI(IUnknown) on the OCX's interfaces returns the OCX),
  and the form dies during load even with a pass-through wrapper.

**Decision.** Accept and document. Configure the control at run time
(`Form_Load`). Design-time values remain cosmetic in the IDE.

---

## 5. A VB error crossing a C frame becomes `RPC_E_DISCONNECTED`

**Observation.** With the dispatch layer implemented in C, a method that raises
(e.g. `RemotePort = 70000`) surfaced in the VB6 caller as
`0x80010108 "The object invoked has disconnected from its clients"` instead of
`40021`. The same call made from VB returns `0x800A9C55`, which VB6 maps back to
`40021`.

**Consequence.** `IDispatch` had to move into `modWinsock2.bas`. `W2_Invoke`
catches the raise (`On Error GoTo Failed`), fills `EXCEPINFO` and returns
`&H800A0000 Or (Err.Number And &HFFFF&)`. No C code sits between the client and
the VB method on the success/error path.

---

## 6. VB6 passes dispinterface arguments by reference; form sinks read the pointer

**Observation.** Raising `ConnectionRequest a1` from VB6 delivered `1702696`
(a pointer) to a form handler declared `ByVal requestID As Long`, while a
`WithEvents` class sink received the correct `1`. The form's compiled sink reads
the VARIANT's value without dereferencing a byref argument.

**Evidence.** Logging at the sink, then changing the call:

```vb
evObj.ConnectionRequest (a1)   ' parentheses force a temporary -> by-value VARIANT
```

The form then logged `connreq 1` and `bytes=5`.

**Rejected alternatives.** `CallByName evObj, "ConnectionRequest", VbMethod, a1`
loses the event entirely: a form sink does not answer `GetIDsOfNames`.
`DispCallFunc` needs VARIANT-shaped arguments (`VARIANTARG**`), and a wrong
description crashed the call.

---

## 7. `Close` + `Accept requestID` requires pending requests to survive `Close`

**Observation.** The canonical VB6 server pattern is

```vb
Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
    Winsock1.Close
    Winsock1.Accept requestID
End Sub
```

Closing the listener used to discard its pending accepted sockets, so `Accept`
raised `40014`. The OCX keeps them.

**Fix.** `W2_Close` no longer calls `W2_ClosePending` for a listening socket;
requests that are never accepted are closed in `W2_FreeState`.

---

## 8. Code-created controls abort if persistence is offered

**Observation.** `WebServer.exe` (`USE_MSWINSOCK = True`) creates the control as a
class (`New Winsock`). Its trace ended at `IPersistStreamInit::GetSizeMax` and the
process exited silently; `Bind` was never reached. With persistence refused it
continued and bound port 80.

**Cause.** For this path VB6 probes `IPersistStreamInit` and
`IPersistPropertyBag` and enters a handshake the host never completes.

**Fix.** The shim exposes both persistence interfaces **only after
`IOleObject::SetClientSite`**: a form host calls it before asking; a code-created
object never does, gets `E_NOINTERFACE`, and proceeds.

---

## 9. External manifests are unreliable; embedded ones have strict rules

**Observations.**

* An external `<exe>.exe.manifest` is sometimes ignored for a given path after a
  recompile; the app then silently uses the real OCX. A fresh path (renamed EXE)
  or an embedded manifest fixes it.
* An embedded manifest is parsed strictly: the resource must be a multiple of four
  bytes (RC pads with NULs → "Invalid Xml syntax"), and `<comClass>` must be on a
  single line.
* A new EXE name triggers UAC/firewall prompts; the test EXE name is therefore
  kept stable.

**Consequence.** `winshim.log` next to the host is the honest indicator of which
server ran. The form test embeds its manifest; deployed apps use the external one.

---

## 10. Toolchain traps

* **PowerShell 5.1 argument quoting.** Native arguments are not quoted, so
  `"Default Project"` truncates. Use `ProcessStartInfo.Arguments` with explicit
  quotes.
* **CRLF.** `.frm`/`.vbp` with LF endings lose their `Attribute` lines; VB6 then
  reports `Form1` as an undefined variable in `modMain`.
* **TCC exports.** `__stdcall` names are decorated and TCC has no `.def` support;
  `build.ps1` patches the four export strings in place and verifies them.
* **`Set-ItemProperty -Name '(default)'`** creates a value named `(default)`
  instead of the default value; COM ignores it. Use the .NET registry API.

---

## 11. Behaviours found and fixed in `modWinsock2` by the tests

| Finding | Fix |
| --- | --- |
| `SendData` used `Len(s)` (characters) as the byte count, so ANSI conversion could read past the buffer | count bytes from the converted array |
| `Scode` in the Error event did not match the OCX | `Scode = &H800A0000 Or number` |
| UDP raised `SendProgress`/`SendComplete`; the OCX does not | suppressed for UDP |
| `Close` returned before the socket was released | close is synchronous (best-effort FIN, then release) |
| Stale `WSAAsyncSelect` messages could arrive after re-binding a slot | drained |
| The VB runtime could unload the module while COM still held objects | module lifetime pinned |
| Empty, `Null` and `Nothing` sends raised; the OCX ignores them | ignored |
| `10061` description text differed | matched |
| `BytesReceived`, `GetData` and `Bind` after close/in-use crashed the OCX | W2 handles them (probes W2-only) |

---

## 12. Oracle defects (recorded, not imitated)

`MSWINSCK.OCX 6.01.9782` raises an access violation for:

* `BytesReceived` after `Close`;
* `GetData` without a connection;
* `Bind` to a port already in use;
* wrong-order calls such as `Connect`/`Listen` while connected.

It also drops most of a 4,000-datagram burst (≈96% loss on this host), while W2
delivers all of them.

---

## 13. UDP read timing and the `Error.Source` difference

* The OCX discards a datagram that is not read inside `DataArrival`; W2 keeps it
  buffered. The harness reads inside the event (the documented pattern for both),
  so this does not affect verdicts.
* The OCX reports its module path as `Error.Source`, W2 reports `MSWinsock`.
  `Number`, `Description` and `Scode` match.

---

## 14. `WebServer.exe` has two independent socket modes

* `USE_VBWINSOCK = True` → raw `Declare Function New_WinsockVB6` (no COM, no
  manifest; uses `WinsockVB6.dll` from the app folder).
* `USE_MSWINSOCK = True` → COM creation through the CLSID (manifest or
  registration; `winshim.dll` + `WinsockVB6.dll`).

This explains why the manifest "matters only when the CLSID is requested", and why
the application test must run with `USE_MSWINSOCK = True` from a copy of the app.
