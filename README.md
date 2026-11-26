# WinsockVB6

This project is a replacement module for `MSWINSCK.OCX`, written entirely in Visual Basic 6.
(shim dll, to also replace the `Control on a Form`, is in C, because it was not the project heart).
It can be used standalone within one of your own programs, or it can completely
replace an existing control in an already compiled executable—without requiring any changes
to the code or access to the original source code.

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

## Usage

1. Import `modWinsock2.bas` in your vb6 project.
2. Add reference to `Win32API.tlb`.
3. Create a new winsock with :
```
Dim sock As Winsock
Set sock = New_WinsockVB6()
```

## Credits

WinsockVB6 
by EBArtSoft@  
Created with Microsoft Visual Basic 6  
AI assistance: OpenCode + DeepSeek V4.1 Flash 1017

> [!IMPORTANT]
> Distributed VB6 executables may appear as viruses; simply let your antivirus software remove them and compile the project you need from Visual Studio. Never trust a .EXE from unknown source.

> [!CAUTION]
> The OCX replacement using the shim dll can break your existing vb6 projects do not register it system wide, use *.manifest instead

> [!NOTE]
> This project does not claim to be a 100% `MSWINSCK.OCX` replacement. Its primary goal is to provide a single .bas module to avoid using distributing OCX with your apps and... without too much work.

