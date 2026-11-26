# 06 — State of development

Snapshot: 2026-11-25, `main` at `7705b45` ("Some cleanup").

## Status by layer

| Layer | State | Evidence |
| --- | --- | --- |
| `modWinsock2` API surface (`IMSWinsockControl`) | Working; differential suite green | 32 PASS / 0 DIFF / 0 FAIL / 1 EXPECTED, 33 scenarios |
| Events (7) | Working, including form sinks | `connreq 1`, `bytes=5`, `hello`/`world` round trip |
| `IDispatch` / dispid mapping | Working | `CallByName(sock, "State")` matches the OCX |
| Raw DLL export | Working | `New_WinsockVB6` ABI verified; `IRL-Test` VBWINSOCK mode |
| Shim identity (`DllGetClassObject`, factory) | Working | `testhost.exe` |
| OLE control surface (form) | Working | two controls on a form, full loopback |
| Code-created control via COM | Working | `WebServer.exe` binds `0.0.0.0:80` |
| Manifest deployment | Working | per-app manifest; app verified |
| Design-time property values | **Not applied** | accepted limitation |
| No Win32API.tlb build | **Parked** | `NO_TYPELIB` area exists, empty |
| VB6 IDE designer / toolbox | **Not supported** | v2 |
| 64-bit | N/A | VB6/x86 only |

The source has since been reformatted to the project's VB6 conventions
(declaration order, single-line statements, aligned `Dim` groups, a banner above
every procedure). The canonical token stream is unchanged, so none of the
behavioural statements above changed with the formatting.

## Verified end-to-end

1. **Form test** (`winshim\testform`, embedded manifest):
   `listen → connect → connreq 1 → accepted → 'hello' → 'world' → closed 0/0`.
2. **Real app** (`IRL-Test\WebServer.exe`, run from a copy under `%TEMP%`,
   `USE_MSWINSOCK = True`): stays running, listens on `0.0.0.0:80` through the
   shim; `winshim.log` shows the expected QIs and refused persistence.
3. **Feature suite** `FeatureTests.exe /autorun`: 32 PASS / 1 EXPECTED in ~17 s.
4. **Matrix** (historical, 199 cells): MATCH 101, DIVERGENCE 42, ORACLE-LIMITED
   56, W2_BROKEN 0. The 56 are harness gaps, not replacement failures.

## Open items

### A. Design-time property values (accepted)

The compiled EXE contains the saved record, but VB6 only calls `InitNew` and
sends no `propput`; every plausible cause was eliminated (see
`05-discoveries.md` §4). Workaround: set `Protocol`, `LocalPort`, `RemoteHost`,
`RemotePort` in `Form_Load`. Reopening this needs a different instrument (for
example a debugger trace of the OCX inside VB6), not more COM surface guessing.

### B. Error descriptions are generic

`Err.Number` matches the OCX (`40014`, `40021`, `10061`…), but the description is
usually "Application-defined or object-defined error" because the HRESULT is
returned without `IErrorInfo`. `EXCEPINFO` alone was not enough (VB6 showed
`440`). Low priority: numbers are what applications branch on.

### C. `RemotePort = 70000` parity

W2 raises `40021` (`sckOutOfRange`); the OCX raises `40014`. One-line change in
`W2_LetRemotePort` if exact parity is wanted; left as is because the matrix treats
it as a known difference.

### D. Matrix divergences still open

From the 42 divergence cells (families `conn`, `pg`, `dns`, `send`, `close`,
`evt`, `abort`, `conc`):

| Area | Difference |
| --- | --- |
| `Connect`/DNS validation | W2 reports `10049` for an empty host; the OCX is silent in some variants |
| Properties during connecting / after a failed connect | snapshot differences |
| State after peer FIN | W2 reports `0` (closed); OCX reports `8`/`9` (`sckClosing`/`sckError`) |
| Reset classification | W2 raises `Close`; OCX raises `Error 10053/10054` (contradicts the original `IRL-Test` symptom — must be reconciled before changing) |
| `udpNoRemotePort` | send without a remote port differs |
| `SendProgress` aggregation | `dataCoalescing`, `close:midSend` differ |

### E. Interface surface still missing

| Interface | Why it is absent | Impact |
| --- | --- | --- |
| `IQuickActivate` | deliberately refused (slow path) | none observed |
| `IViewObjectEx` | stubs crashed the form; reverted | none observed in form use |
| `IPointerInactive` | refused | none observed |
| `IPersistPropertyBag` replay | accepted but does not apply values | related to item A |
| `IProvideClassInfo` | not implemented | type info unavailable (VB6 does not ask) |

### F. Harness gaps

* 56 `ORACLE_NO_OUTPUT` cells (`ps`, `accept`, `listen` families with
  `impl=ocx`).
* The `get` family measures the no-data path instead of the `MaxLen`/type matrix.
* The suite covers the OCX and W2 only; third-party engines are not part of the
  1:1 runs (see parked work).

## Parked work

### P1. Build without `Win32API.tlb`

`modWinsock2.bas` already has the intended area:

```vb
#If NO_TYPELIB Then
'PUT DATA HERE
#End If
```

It is empty. The inventory was completed:

* **Local UDTs (4):** `CLSID`, `SOCKADDR`, `WSADATA`, `HOSTENT` (layouts known
  from the type library, e.g. `WSADATA` is 396 bytes with 256/128-byte fixed
  strings).
* **Local declares (33):** 21 `ws2_32` functions, 6 `kernel32`, 5 `user32`
  (`A` variants), 1 `ole32` (`CLSIDFromString`). Six further APIs are already
  declared locally (`W2_*`).
* **No TLB constants are used**; the module only uses its own `W2_*` constants.
* **Blocker:** seven COM types cannot be declared in a `.bas`:
  `Winsock`, `DMSWinsockControlEvents`, `ITypeInfo`, `IConnectionPoint`,
  `IConnectionPointContainer`, `IEnumConnectionPoints`, `IEnumConnections`.

Options considered: (A) reference `TypeLib\MSWinsockLib.tlb` for the Winsock
object/event types and drop `Win32API.tlb`; (B) replace the seven types with
`Object` (late binding), which is module-only but changes the factory typing and
`Common.bas`. The work is paused pending a decision.

### P2. Third-party implementation comparison (`wqweto/VbAsyncSocket`)

> [!IMPORTANT]
> @wqweto: Although I’ve had the idea since the days of VB5, your work inspired me and gave me the motivation to finish this project. Thank you.

Analysed but not adopted and not in the bench. Summary:

* `src\cAsyncSocket.cls` is a thin VB6 async socket engine with a different API
  (`Create`, `AsyncSelect`, `Send`/`Receive`, `OnConnect`/`OnReceive`…), no
  MSWINSCK shape.
* `contrib\ctxWinsock.ctl` is the intended MSWINSCK control replacement: same
  protocol/state/error values, same method names except `Close_`, same events
  except `CloseEvent`, missing `BytesReceived` and `AboutBox`, plus TLS/timeout
  extras.
* Its socket creation passes `SocketType:=(m_eProtocol And 1)` (0/1), while the
  engine’s enum is `1=stream`, `2=datagram`; this needs verification before it
  could be trusted in a benchmark.
* Adding it to the **fast** performance suite (not the 1:1 scenarios/matrix)
  would require an adapter in `FeatureTests` and, for the control, hosting a
  `UserControl` on the test form.

## Deferred to v2 (design-time)

1. Register the MSWinsockLib-compatible typelib machine-wide and/or load it from
   the DLL.
2. Implement `IProvideClassInfo` returning the coclass's `ITypeInfo`.
3. Implement `IClassFactory2` licensing as a real design-time surface.
4. Register the CLSID with `Control`/`Insertable` categories, `MiscStatus`,
   `ToolboxBitmap32`, `DefaultIcon`, and the typelib reference.
5. Only then expect the toolbox to accept the control and the IDE to host it.

## Risks and mitigations

| Risk | Mitigation |
| --- | --- |
| Manifest silently ignored → app uses the real OCX | Check `winshim.log`; prefer embedded manifests where possible |
| Machine-wide registration affects other software | Use manifests; `regsvr32 /u` restores the OCX |
| VB raise crossing C kills the host | Keep dispatch in VB; never call VB methods from C with error-capable semantics |
| Identity shim (QI forwarding to the inner) | Works for forms and code-created objects; do not reintroduce a C forwarder for methods |
| Type-library swap breaks `As Winsock` users | Decide P1 route; verify with `As Object` coercion before committing |
| Stale documentation | Papers are authoritative; READMEs updated alongside |

## Acceptance criteria for v1 (met)

* Code-created and form-hosted use work through the manifest.
* The documented API and all seven events match the OCX in the differential suite
  (modulo the classified `EXPECTED` wrong-order scenario).
* Deployment needs only a manifest and two DLLs.
* The implementation is our own code; no OCX bits are redistributed.

## Suggested next steps (in order)

1. **Reconcile the reset/`Close` classification** (§D) — the one divergence that
   contradicts the original application symptom.
2. **Fix the matrix harness gaps** (`ORACLE_NO_OUTPUT`, `get` family) so the 42
   divergences can be trusted and closed.
3. **Decide P1** (no-TLB build) if a self-contained module is required.
4. **v2 typelib/designer work** if design-time hosting is required.
5. **Error descriptions** (`IErrorInfo`) as polish.
