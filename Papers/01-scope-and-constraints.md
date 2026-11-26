# 01 — Scope, goals and constraints

## Problem

The application under test (`IRL-Test\WebServer.exe`, a VB6 web server) was built
against Microsoft's `MSWINSCK.OCX` and also carries a home-grown replacement
(`modWinsock.bas` + `modWinsock2.bas` + `WinsockVB6.dll`). The goal that grew out
of the original analysis was not merely "fix the clone": it was to make the pure
VB6 implementation usable **as the control**, i.e. under the same CLSID as the
OCX, for both code-created and form-hosted use.

Two facts shape the problem:

1. A VB6 ActiveX DLL can only answer `DllGetClassObject` for its own classes. It
   cannot serve `{248DD896-BB45-11CF-9ABC-0080C7E7B78D}` (the MSWINSCK coclass)
   directly.
2. A control dropped on a form is not just the socket API. VB6 asks it for the
   OLE control interfaces (`IOleObject`, `IViewObject`, `IOleControl`,
   persistence), and calls its properties through `IDispatch` with the dispids
   from `MSWinsockLib`.

Hence the architecture: a shim DLL owns the identity and the OLE surface, and
forwards to the VB object.

## v1 goals (all reached unless marked)

1. **API parity** — every `IMSWinsockControl` member and all seven events behave
   like `MSWINSCK.OCX 6.01.9782` for documented use, verified differentially.
2. **Code-created control** — `New Winsock`, `CreateObject("MSWinsock.Winsock")`
   and the raw DLL export all reach the same implementation.
3. **Form-hosted control** — the control can be dropped on a form, create its
   windowless shell, exchange data, and raise events with correct arguments.
4. **Registry-free deployment** — a per-application manifest next to the host
   EXE is enough to redirect the CLSID.
5. **No new public surface** — no extra public methods or members on the Winsock
   object; the vtable layout and typelib remain authoritative.
6. **The tested module stays clean** — test harnesses link to `modWinsock2.bas`;
   diagnostic code is behind conditional compilation (`W2DIAG`).

## Non-goals (deferred or excluded)

| Item | Status |
| --- | --- |
| VB6 IDE designer / toolbox | v2. Needs a registered typelib, `IProvideClassInfo`, Control/Insertable categories, toolbox bitmap, and design-time licensing. |
| Restoring design-time property values at run time | Accepted limitation; see `05-discoveries.md` and `06-state-of-development.md`. Configure at run time. |
| 64-bit builds | Impossible: VB6 is 32-bit, the shim is x86. On x64 Windows the CLSID lives in `WOW6432Node`. |
| IPv6, TLS, other transports | Out of scope; the OCX does not offer them either. |
| Redistributing `mswinsck.ocx` | Not allowed. The shim reuses only the identity; the object is our code. |
| Registering the typelib machine-wide | Approved in principle for v2, postponed; not needed for the manifest route. |

## Hard constraints and how they shaped the work

* **No new public members.** The dispatch layer had to map the OCX's dispids onto
  the *existing* typed members; it could not add automation-only entry points.
  `W2_Invoke` and `W2_GetIDsOfNames` are existing `IDispatch` vtable slots, not
  new members.
* **`modWinsock2.bas` is the implementation under test.** It started as
  "never modified"; that held until IDispatch and two socket behaviours had to
  live in VB (see `05-discoveries.md`). All harness code still lives outside it.
* **`IRL-Test\` is read-only.** Every reproduction of the application runs from a
  copy under `%TEMP%`. Nothing in `IRL-Test\` was modified.
* **Differential testing, not unit tests.** The oracle is the real OCX on the
  same machine. A behaviour is "correct" when it matches the OCX (or when the
  OCX itself is broken, which is recorded).
* **Expensive sweeps are not casual.** The 199-cell matrix takes about nine
  minutes; it is run deliberately, not on every change.
* **No unapproved downloads.** Tooling had to come from what was already on the
  machine (`vb6_ide`, `Tools\vb98`, `Tools\tcc`, `Tools\mingit`, `Tools\midl`).

## Terminology

| Term | Meaning |
| --- | --- |
| W2, the replacement | `modWinsock2.bas` and the object it creates |
| OCX, oracle | `<OCX path>`, registered in the 32-bit view |
| Inner object | The VB socket object returned by `New_WinsockVB6` |
| Ctl / wrapper | The C object in `winshim.dll` that adds OLE control interfaces and COM identity |
| Shim | `winshim.dll` as a whole (class factory + wrapper + registration) |
| Dispid | Automation member id from `MSWinsockLib`, e.g. `State` = 8, `Close` = 70 |
| Sink | The object a control raises events on (a VB6 form or a `WithEvents` class) |
| UserControl path | Form-hosted control: VB6 calls the OLE control interfaces |
| Code-created path | `New Winsock` in a VB6 EXE: `IMSWinsockControl` + persistence probes, no client site |
| Manifest route | Per-application SxS redirection of the CLSID to `winshim.dll` |
| Oracle-limited | A matrix cell the OCX cannot answer (crash, hang, dialog, no output) |
