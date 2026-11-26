# winshim — Visual Studio 2019 project

Builds `winshim.dll` (32-bit, in-process COM server) from `..\winshim.c` with the
MSVC v142 toolset. The TCC build in `build.ps1` remains the reference build.

## Requirements

* Visual Studio 2019 (16.x) with **Desktop development with C++**
* A Windows 10 SDK (the project asks for `10.0` and takes the newest installed
  one)
* Platform **Win32 (x86)** only — the shim is a 32-bit COM server and the `.def`
  file exports `__stdcall`-decorated names, which only exist on x86

## Build

1. Open `winshim.sln`.
2. Select **Release** and **Win32**.
3. Build → Build Solution.

Or, from a *Developer Command Prompt for VS 2019*:

```
msbuild winshim.sln /p:Configuration=Release /p:Platform=Win32
```

Output:

```
winshim\out\msvc\Win32\Release\winshim.dll
```

(Debug builds land in `...\Win32\Debug\`.)

If `..\..\AsDll\WinsockVB6.dll` exists, a post-build step copies it next to the
shim so the pair can be tested immediately.

## Project settings that matter

* **C code, C89-compatible.** `winshim.c` is written in the common subset of C89
  and is compiled with `CompileAsC`; the project does not enable `/std:c11`
  because the source does not need it.
* **`WIN32_LEAN_AND_MEAN`.** Keeps the SDK from pulling in the OLE headers, whose
  `IUnknown` / `IOleObject` / `IDispatch` definitions would collide with the
  shim's hand-rolled vtables (the shim defines only the slots it implements).
* **`WINSHIM_NO_DLLEXPORT`.** Compiles the four entry points without
  `__declspec(dllexport)`, so the `.def` file is the only export source and the
  DLL has no decorated aliases. The TCC build leaves the macro undefined and is
  unaffected.
* **`WINSHIM_NO_LOG` (Release only).** Compiles the trace log out of the DLL:
  no `winshim.log`, no formatting calls, no log strings in the image. Debug
  builds leave it undefined and log exactly like the TCC build. The TCC script
  offers the same switch as `build.ps1 -NoLog`.
* **Static CRT** (`/MT`, `/MTd`). No `vcruntime*.dll` / `msvcp*.dll` dependency,
  which keeps the deployment identical to the TCC build.
* **No extra import libraries.** The shim touches only `kernel32`
  (`LoadLibraryA`, `GetProcAddress`, registry stubs via `advapi32` loaded at
  runtime) and `user32` (`wsprintfA`). `oleaut32!SysAllocString` is resolved with
  `GetProcAddress`, so there is no import for it.

## Why there is a .def file

MSVC decorates `__stdcall` exports (`_DllGetClassObject@12`). COM and the host
programs expect the undecorated names. `winshim.def` maps:

```
DllGetClassObject=_DllGetClassObject@12
DllCanUnloadNow=_DllCanUnloadNow@0
DllRegisterServer=_DllRegisterServer@0
DllUnregisterServer=_DllUnregisterServer@0
```

`winshim.c` marks these four functions `__declspec(dllexport)` only when
`WINSHIM_NO_DLLEXPORT` is **not** defined. The project defines it, so the `.def`
file is the only source of exports: the DLL exposes exactly the four plain
names, with no `_DllGetClassObject@12`-style aliases. The TCC build does not
define the macro, so `build.ps1` still sees the decorated names it patches.

Do not remove the `.def`: without it nothing is exported at all in this build.

Do not build this project for x64: the decoration and the aliases do not apply.

## Verify the export table

```
dumpbin /exports winshim\out\msvc\Win32\Release\winshim.dll
```

Only the four plain names should be listed — `DllGetClassObject`,
`DllCanUnloadNow`, `DllRegisterServer`, `DllUnregisterServer`. A name such as
`_DllGetClassObject@12` means the source was compiled without
`WINSHIM_NO_DLLEXPORT`.

## Verification status

* The C source compiles cleanly as C89 (`-std=c89 -pedantic`,
  `WIN32_LEAN_AND_MEAN`) with only three "macro redefined" notes for
  `CLASS_E_NOAGGREGATION`, `CLASS_E_CLASSNOTAVAILABLE` and
  `REGDB_E_CLASSNOTREG`, which the SDK also defines with the same values. With
  `WINSHIM_NO_DLLEXPORT` the declarations keep `__stdcall` and lose only the
  export attribute.
* The TCC build was re-run after the change and produced a **byte-identical**
  `winshim.dll` (same SHA-256, 11,264 bytes) with the four plain exports, so the
  reference build is unaffected.
* With `WINSHIM_NO_LOG` the same source compiles to 8,192 bytes with no
  `winshim.log`, no `IOleObject::`/`IClassFactory2::` format strings, and the
  four exports unchanged. A `WINSHIM_NO_LOG` DLL that still writes
  `winshim.log` means the macro did not reach the compiler.
* **This VS2019 project has not been compile-verified in the authoring
  environment** (no MSVC installation there). It is a standard v142 Win32 DLL
  project.
* Runtime behaviour is covered by the project tests and is identical to the TCC
  build: same source, same exports, same runtime layout.

## Runtime layout

Keep the same layout as the TCC build: `winshim.dll` and `WinsockVB6.dll` next
to the host EXE, plus the per-application manifest that redirects the
`MSWINSCK.OCX` CLSID.
