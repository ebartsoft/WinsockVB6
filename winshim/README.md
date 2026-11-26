# winshim — publish WinsockVB6 under the MSWINSCK.OCX identity

A host compiled against `mswinsck.ocx` asks COM for CLSID
`{248DD896-BB45-11CF-9ABC-0080C7E7B78D}`. A VB6 ActiveX DLL can only answer
`DllGetClassObject` for the CLSIDs of its own classes, so it cannot serve that
one directly. `winshim.dll` is an ~11 KB in-process server that accepts the
MSWINSCK CLSID and hands back the object created by
`WinsockVB6.dll!New_WinsockVB6`.

```
host  --CoCreateInstance({248DD896…})-->  winshim.dll
                                             |
                                             +-- LoadLibrary("WinsockVB6.dll")
                                             +-- New_WinsockVB6()  -> IMSWinsockControl
```

For a control **dropped on a VB6 form**, the object also has to be a full OLE
control. `winshim.dll` adds that layer around the socket object:

| Interface | What it does |
| --- | --- |
| `IOleObject` | client site, extents, `DoVerb`, advise |
| `IViewObject` / `IViewObject2` | `Draw` is a no-op (empty shell); advise stored |
| `IOleControl` | `GetControlInfo`, `FreezeEvents`, ambient changes |
| `IPersistStreamInit` | `InitNew`/`Load`/`Save` accepted; no properties of its own |
| `IDispatch` | forwarded to the VB object, which maps the MSWinsockLib dispids |

`IQuickActivate` and `IPointerInactive` are deliberately not offered, so VB6
takes the slow, documented activation path.

The dispatch layer lives in `modWinsock2.bas`, not in this DLL: a VB error
raised inside a method only reaches the caller as a normal HRESULT when the call
is made from VB. Through a C frame the raise escapes as
`RPC_E_DISCONNECTED (0x80010108)`, so a VB6 form saw "Automation error / object
disconnected" instead of `40021`.

## Files

| File | Purpose |
| --- | --- |
| `winshim.c` | The shim: `DllGetClassObject`, `DllCanUnloadNow`, `DllRegisterServer`, `DllUnregisterServer`, one static `IClassFactory`, plus the OLE control layer |
| `probe.c` | Determines the calling convention of the VB6 export (see below) |
| `testhost.c` | End-to-end check: class object → `CreateInstance` → `get_State` |
| `build.ps1` | Builds all three with TCC and patches the export names |
| `testform\` | VB6 project with two Winsock controls on a form; decompiled by `/make` and used to verify the control surface |
| `app.exe.manifest` | Per-application, registry-free deployment |
| `winshim.reg` | Reference registry layout for the machine-wide route |
| `out\` | Build output (not committed) |

## Build

```powershell
powershell -ExecutionPolicy Bypass -File build.ps1           # logging build
powershell -ExecutionPolicy Bypass -File build.ps1 -NoLog    # silent build
```

Requires Tiny C Compiler; the script uses `%TCC%` or
`<TOOLS>\tcc\tcc.exe`. It copies `..\AsDll\WinsockVB6.dll`
into `out\` so the tests can run immediately. `-NoLog` passes
`-DWINSHIM_NO_LOG` and verifies that the log strings are gone; the default
build still writes `winshim.log` next to the host.

### Visual Studio 2019

`vs2019\winshim.sln` builds the same DLL with the MSVC v142 toolset (Win32 only).
The TCC script above remains the reference build; see `vs2019\README.md` for
the output path and the `.def` export mapping. The VS **Release** configuration
compiles the trace log out (`WINSHIM_NO_LOG`): no `winshim.log` is written. The
**Debug** configuration logs, as does the TCC build by default; pass `-NoLog` to
`build.ps1` to silence the TCC build too.

TCC decorates `__stdcall` exports (`_DllGetClassObject@12`) and has no `.def`
support, so `build.ps1` rewrites those four strings in the finished DLL. They
only ever get shorter and are NUL padded in place, which leaves every RVA in the
export directory valid; the script verifies all four names afterwards.

## Verified on this machine

```
probe.exe            -> com-abi: hr=0x0076C608  obj=00000000      (wrong ABI)
probe.exe raw        -> raw: obj=006BA650                          (correct)
                         QI IMSWinsockControl          hr=0
                         QI IConnectionPointContainer  hr=0

testhost.exe         -> DllGetClassObject(IID_IClassFactory) hr=0  cf=10002064
                         CreateInstance(IMSWinsockControl)     hr=0  obj=006E9E58
                         get_State                             hr=0  state=0
                         get_Protocol                          hr=0  protocol=0
                         QI IConnectionPointContainer          hr=0
```

The first line matters: **`New_WinsockVB6` is not a COM-style
`HRESULT f(IUnknown **)`**. It returns the interface pointer in `EAX` and takes
no arguments (`void *__stdcall f(void)`). Calling it the COM way yields a
garbage HRESULT and a null object. The shim uses the raw form.

## Form-hosted control (verified)

`testform\` contains a VB6 project with two `MSWinsockLib.Winsock` controls and
a timer. One listens on 65000, the other connects, and the server echoes a
message. Compiled to `testform\winshimtest.exe`, run with
`winshimtest.exe.manifest` next to it:

```
Form_Load state=0 proto=0 localport=65000 handle=-1
listen state=2 err=0
connect state=6 err=0
client connect event state=7
client send err=0
server connreq 1
server accepted state=7 err=0
server got 'hello' bytes=5 err=0
client got 'world' bytes=5 err=0
closed states 0/0
```

So the control creates its window, its properties and methods work through
`IDispatch`, events carry correct arguments, and the canonical
`Close` + `Accept requestID` pattern succeeds. `winshim.log` next to the EXE
confirms the shim was the server; if that file is absent, Windows ignored the
manifest (usually an invalid `winshim.dll`) and VB6 silently used the real OCX.

## Controls created from code (e.g. WebServer.exe)

A VB6 program can also create the control as a class (`New Winsock`) with no
form. That path asks for `IMSWinsockControl`, then `IPersistStreamInit` and
`IPersistPropertyBag`, and calls `GetSizeMax`; if persistence is offered there
the host aborts silently right after `GetSizeMax` and the program exits.

The shim therefore offers the two persistence interfaces **only after
`IOleObject::SetClientSite`**. A form host calls that first, so it still gets
persistence; a code-created object gets `E_NOINTERFACE`, skips the handshake and
goes on to configure the socket. Verified with `IRL-Test`'s `WebServer.exe`
(`USE_MSWINSOCK = True`): it stays running and binds `0.0.0.0:80` through the
shim.

Building it with VB6 needs care on this machine:

* PowerShell 5.1 does not quote native arguments, so `"Default Project"`
  truncates the path and VB6 opens `<parent>\Default.vbp`. Use
  `ProcessStartInfo.Arguments = '/make "<full path>\winshimtest.vbp"'`.
* `.frm`/`.vbp` files must keep CRLF line endings; with LF the `Attribute`
  lines are not parsed and `Form1` is reported as an undefined variable.
* The test EXE **embeds** its manifest (`winshimtest.rc` + `.res`, linked via
  `ResFile32` in the `.vbp`) instead of relying on `winshimtest.exe.manifest`.
  External manifests stop being honoured for this path after some recompiles
  and the app then silently loads the real OCX. An embedded manifest is parsed
  strictly: the XML must be a multiple of four bytes long (RC pads with NULs
  otherwise and SxS reports "Invalid Xml syntax"), and `<comClass>` must be on
  one line.
* Keep the EXE name stable when running it: a new name makes Windows treat it
  as a new program (UAC and firewall prompts).

Rebuild the resource after editing the manifest:

```powershell
& "$env:ProgramFiles\..\Documents\Tools\vb98\Wizards\RC.EXE" /fo winshimtest.res winshimtest.rc
```

## Deploy per application (recommended)

Copy next to the host EXE, and rename the manifest to match it:

```
WebServer.exe
WebServer.exe.manifest      <- app.exe.manifest
winshim.dll
WinsockVB6.dll
```

No registry, no admin, no effect on other software. Uninstall = delete the three
files.

## Deploy machine wide

```bat
regsvr32 <path>\winshim.dll     :: writes the MSWINSCK CLSID + ProgID keys
regsvr32 /u %WINDIR%\SysWOW64\mswinsck.ocx   :: undo, re-registers the real control
```

`DllRegisterServer` registers HKLM (needs an elevated prompt) and deliberately
does **not** delete or replace the `Control` / `Implemented Categories` /
`TypeLib` entries: those belong to the OLE control surface, which this shim does
not implement. The real typelib registration stays in place so VB6 can still
compile projects that reference `mswinsck.ocx`.

## Scope and limits

* **Code-created sockets and form-hosted controls both work.**
  `New Winsock`, `CreateObject("MSWinsock.Winsock")` and a control dropped on a
  form all reach the same object.
* **Design-time property values are not applied.** The compiled EXE does carry
  the saved record (the BSTR `127.0.0.1`, the long `65000`, interface GUID
  `248DD892`), but VB6 calls `IPersistStreamInit::InitNew` on our object and
  never `Load`, and sends no `propput` either. Ruled out: manifest/SxS vs
  registry activation, typelib registration, type info (never requested),
  licensing (`IClassFactory2` is now implemented), and the missing
  `IViewObjectEx`/`IPointerInactive` interfaces (tried, then reverted). A
  logging proxy around the real OCX was attempted and is not viable: the OCX
  and the wrapper do not share COM identity, and the form dies during load.
  Configure the control at run time (`Protocol`, `LocalPort`, `RemoteHost`,
  `RemotePort` in `Form_Load`).
* **Event arguments are delivered by value.** VB6 passes a variable to a
  dispinterface by reference, and a form sink then read the pointer as the
  value (`requestID` arrived as 1702696). The events are raised with
  parenthesised arguments, which forces a temporary and a by-value VARIANT; the
  form test now logs `connreq 1` and `bytes=5`.
* **`Close` + `Accept requestID` works.** Closing a listener keeps its pending
  requests, as the OCX does; a request that is never accepted is closed when the
  object is freed.
* **32-bit only.** VB6 hosts are 32-bit and the shim is x86; on 64-bit Windows
  the CLSID lives in `WOW6432Node`.
* **Machine-wide registration affects every process** that uses the control,
  including programs that are not being tested. Prefer the manifest.
* **Design-time hosting is not covered.** The VB6 designer still needs a registered typelib and the toolbox/`Control` categories; only the compiled-EXE path is supported.
* **Do not redistribute `mswinsck.ocx`.** The shim only reuses its identity;
  the object behind it is our own code.
