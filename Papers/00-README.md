# WinsockVB6 — Papers

Technical documentation for the project that replaces `MSWINSCK.OCX` with a pure
VB6 implementation and publishes it under the original control's CLSID.

The papers record **what the project is, how it is built, how it is tested, what
was discovered, and where development stands**. They are meant to be readable
without the conversation history that produced them.

## The project in one paragraph

`modWinsock2.bas` implements `IMSWinsockControl` (the `Winsock` control's
interface) in pure VB6 by fabricating COM vtables in memory. `AsDll\WinsockVB6.dll`
wraps it behind a single export, `New_WinsockVB6`, which returns the object as a
raw pointer. `winshim\winshim.dll` is a small C in-process server that answers the
MSWINSCK CLSID `{248DD896-BB45-11CF-9ABC-0080C7E7B78D}`, adds the OLE control
layer VB6 expects when a control is dropped on a form, and forwards everything
else to the VB object. Deployment is registry-free through a per-application
manifest, or machine-wide through `regsvr32`.

## Current status

Snapshot: **2026-11-25**, `main` at `7705b45` ("Some cleanup").

| Area | State |
| --- | --- |
| `IMSWinsockControl` behaviour | Implemented; differential suite **32 PASS / 0 DIFF / 0 FAIL / 1 EXPECTED** (33 scenarios) |
| Control created from code (`New Winsock`) | Works: `WebServer.exe` runs through the shim and binds `0.0.0.0:80` |
| Control dropped on a VB6 form | Works: full loopback round trip through two controls on a form |
| Event arguments to form sinks | Correct (by-value marshalling) |
| `Close` + `Accept requestID` | Works (pending requests survive closing the listener) |
| Design-time property values | **Not applied** (documented limitation; configure at run time) |
| No Win32API.tlb build | **Planned, parked**: the `NO_TYPELIB` area exists but is empty |
| VB6 IDE designer / toolbox | **Not supported** (v2: typelib, Control categories, toolbox bitmap) |
| 64-bit | Not applicable: VB6 hosts are 32-bit |

Latest reference run: `FeatureTests.exe /autorun` on 2026-11-25 01:55 —
`PASS=32 DIFF=0 FAIL=0 EXPECTED/SKIP=1 Scenarios=33`, about 17 s.

Recent landmarks:

* the module was reformatted to the project's VB6 style rules (declaration order,
  single-line statements, aligned `Dim` groups, a banner per procedure);
* machine-independent paths: all documentation uses placeholders (`<repo>`, `<TOOLS>`,
  `<VB6>`, `<OCX path>`) instead of absolute paths;
* `AsDll\modDllMain.bas` became `AsDll\Common.bas`, `modMain.bas` was removed and
  the delivery DLL was rebuilt (32 KB);
* `Win32API.tlb` was replaced with a newer, larger build (814 KB).

## Known size facts

| Artefact | Size |
| --- | ---: |
| `modWinsock2.bas` | 112,075 bytes, 3,220 lines |
| `AsDll\WinsockVB6.dll` | 32,768 bytes |
| `winshim\winshim.dll` | 11,264 bytes |
| `FeatureTests\FeatureTests.exe` | ~400 KB |
| `Win32API.tlb` | 814,256 bytes |

## Document index

| File | Contents |
| --- | --- |
| `01-scope-and-constraints.md` | Problem statement, goals, non-goals, hard constraints, terminology |
| `02-architecture.md` | Components, call paths, COM layout, vtable and dispid reference |
| `03-toolchain-and-builds.md` | Tools and versions, exact build recipes, manifests, deployment |
| `04-testing-and-oracle.md` | Differential test method, harness, matrix, verdicts, performance |
| `05-discoveries.md` | Every non-obvious finding, with evidence and consequence |
| `06-state-of-development.md` | Feature status, open items, parked work, risks |
| `07-decisions-and-history.md` | Decision log and commit timeline, including rejected approaches |
| `08-reproduction-and-troubleshooting.md` | Step-by-step rebuild/reverify recipes and a fault table |
| `09-inventory.md` | Repository tree, artefacts, external dependencies, generated files |

## Conventions

* Placeholders are used instead of machine-specific locations: `<repo>` is the
  project root, `<TOOLS>` the external tools folder, `<VB6>` the VB6 installation,
  `<OCX path>` the registered `mswinsck.ocx` and `%WINDIR%` the Windows folder.
  No user names, absolute paths, credentials or account identifiers appear in
  these documents.
* "W2" means `modWinsock2.bas` / the replacement; "OCX" is the real
  `MSWINSCK.OCX` used as the oracle.
* "Inner" is the VB socket object; "wrapper" or "Ctl" is the C object in
  `winshim.dll` that adds the OLE control surface around it.
* Test results under `FeatureTests\results\` are run artefacts. They are
  git-ignored, so numbers quoted here are snapshots, not live data.
