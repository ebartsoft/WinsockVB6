# 09 — Inventory

Repository root: `<repo>` (the `WinsockVB6` project folder). No machine-specific
paths, user names or account identifiers are recorded; `<TOOLS>`, `<VB6>` and
`<OCX path>` stand for the local tool, VB6 and registered-OCX locations.

## Repository root

| Path | Purpose |
| --- | --- |
| `modWinsock2.bas` | **The replacement** (112,075 bytes, 3,220 lines): fabricated COM object, sockets, events, dispatch |
| `modWinsock.bas` | The older clone (analysis target, not used by the replacement) |
| `modUtils.bas`, `frmMain.frm`/`.frx`, `WinsockVB6.vbp` | Original demo application (`modMain.bas` was removed in `7705b45`) |
| `REPORT.md` | Original analysis + §8 on the OLE/form work |
| `Win32API.tlb` | Third-party VB6 type library, replaced in `7705b45` (814,256 bytes) |
| `Win32API.tlb.md`, `Win32API.dump.txt` | Older documentation/dump of that library (predates the replacement file) |
| `TypeLib\` | Extracted `MSWinsockLib` typelib + decompiled IDL |
| `AsDll\` | VB6 ActiveX DLL wrapper and its build output |
| `FeatureTests\` | Differential feature/performance suite and matrix |
| `Tests\` | Early, smaller VB6 test harness for `modWinsock2` |
| `winshim\` | C shim (identity + OLE control layer), probes, form test |
| `Papers\` | This documentation set |
| `IRL-Test\` | The real application under test — **read-only** |
| `Licence\` | Local OCX registration material (user-owned) |
| `backup\` | Timestamped ZIP snapshot (git-ignored) |

`.gitattributes` sets `* -text` (preserve working-tree line endings) and marks
binaries (`*.dll`, `*.exe`, `*.tlb`, `*.lib`, `*.exp`, `*.zip`, `*.frx`).

## `AsDll\`

| File | Purpose |
| --- | --- |
| `WinsockVB6.vbp` | `Type=OleDll`, `modWinsock2` + `Common` + `Dummy.cls`, export `New_WinsockVB6` |
| `Common.bas` | `New_WinsockVB6()` and `StartVBRuntime` (renamed from `modDllMain.bas` in `7705b45`) |
| `Dummy.cls` | Empty class used to keep the DLL a valid ActiveX project |
| `WinsockVB6.dll` (32,768 bytes) | Committed deliverable |
| `WinsockVB6.exp`, `.lib` | Linker artefacts (committed with the DLL) |

## `winshim\`

| File | Purpose |
| --- | --- |
| `winshim.c` (~34 KB) | Class factory, `IClassFactory2`, OLE control layer, QI policy, logging |
| `build.ps1` | TCC build + export-name patching/verification + copy of the VB DLL |
| `probe.c` / `probe.exe` | Compares COM-style and raw calls to `New_WinsockVB6` |
| `testhost.c` / `testhost.exe` | End-to-end factory test |
| `app.exe.manifest` | Deployment manifest template |
| `winshim.reg` | Reference machine-wide registration layout |
| `README.md` | Operational notes for the shim |
| `out\` | Build output (git-ignored) |
| `testform\` | Form host (committed sources, ignored binaries) |
| `testform2\` | User-supplied reference form project (git-ignored) |
| `vs2019\` | `winshim.sln` / `.vcxproj` / `winshim.def` for the MSVC v142 Win32 build |

`testform\` contains `Form1.frm`, `modMain.bas`, `winshimtest.vbp`,
`winshimtest.exe.manifest`, `winshimtest.rc`, `winshimtest.res`; `.exe`, `.dll`,
`.vbw` and logs are ignored.

## `FeatureTests\`

| File | Purpose |
| --- | --- |
| `FeatureTests.vbp`, `frmFeatureTests.frm` | UI and `/autorun` entry |
| `modTestEngine.bas` | Scenarios, comparison, performance, reports |
| `modTestUtil.bas` | Timing, pump, formatting/IO |
| `modMatrix.bas`, `modMatrixCells.bas`, `modMatrixCells2.bas` | 199-cell matrix and child-process runner |
| `clsEndpoint.cls` | `WithEvents` wrapper for either implementation |
| `clsRun.cls`, `clsResult.cls`, `clsPerfRow.cls`, `clsCell.cls` | Result records |
| `results\` | Generated reports (git-ignored) |
| `README.md` | Harness guide |

## `Tests\`

Early harness (`W2Tests.vbp`, `frmW2Tests.frm`, `modW2Test.bas`,
`clsW2TestClient.cls`, `README.md`). Superseded by `FeatureTests`; kept for
history.

## External tools (not part of the repository)

| Tool | Version | Location |
| --- | --- | --- |
| VB6 | 6.00.8176 | `<VB6>\Vb6.exe` |
| TCC | 0.9.27 | `<TOOLS>\tcc\tcc.exe` |
| RC | 5.00.1641.1 | `<TOOLS>\vb98\Wizards\RC.EXE` |
| MIDL | 6.0.6001.16529 | `<TOOLS>\midl\midl.exe` |
| Git | 2.56.1 | `<TOOLS>\mingit\cmd\git.exe` |
| OCX oracle | 6.01.9782 | `<OCX path>` (registered in the 32-bit view) |

Nothing here is installed globally; tools are invoked by absolute path or via
environment variables.

## Generated / ignored artefacts

| Pattern | Why |
| --- | --- |
| `winshim\out\` | TCC build output |
| `winshim\testform\*.exe`, `*.dll`, `*.vbw` | Test binaries |
| `FeatureTests\results\` | Reports and raw logs |
| `FeatureTests\*.exe`, `*.vbw` | Test binary/project state |
| `*.log` | `winshim.log`, `form.log`, build/raw logs |
| `backup\` | Snapshot ZIPs |
| `winshim\testform2\`, `testform3\` | Local/reference copies |

## Runtime files produced by tests

* `winshim.log` — appended next to the **host EXE**; primary evidence of whether
  the shim served the process.
* `form.log` — the form test's own log.
* `FeatureTests\results\latest.*` and timestamped copies.
* Temporary instrumentation logs (`w2bind.log`, `w2invoke.log`) from earlier
  investigations; not part of any build.

## Ownership / do-not-touch

* `IRL-Test\` — the application under test. Never edited; only copied.
* `Licence\MSWINSCK.OCX.REG` — user-owned registration material.
* `backup\` — snapshots.
* `TypeLib\MSWinsockLib.tlb` — extracted from the OCX; authoritative binary.

## Security note

All committed Markdown uses placeholders for local paths and contains no user
names, credentials, API keys, account identifiers or absolute paths. Test
reports under `results\` are generated artefacts and may quote the OCX's module
path; they are not committed.
