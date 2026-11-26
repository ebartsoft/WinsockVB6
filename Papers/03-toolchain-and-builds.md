# 03 — Toolchain and builds

## Tools actually used

| Tool | Version | Path | Role |
| --- | --- | --- | --- |
| Visual Basic 6.0 | 6.00.8176 | `<VB6>\Vb6.exe` | Builds all VB6 projects (`/make`), including the DLL and test EXEs |
| Tiny C Compiler | 0.9.27 (i386) | `<TOOLS>\tcc\tcc.exe` | Builds `winshim.dll`, `probe.exe`, `testhost.exe` |
| Microsoft Resource Compiler | 5.00.1641.1 | `<TOOLS>\vb98\Wizards\RC.EXE` | Builds `winshimtest.res` from the manifest for the embedded-manifest test EXE |
| MIDL | 6.0.6001.16529 | `<TOOLS>\midl\midl.exe` | Compiled `TypeLib\MSWinsockLib.idl` as a check against the extracted TLB |
| Git for Windows (MinGit) | 2.56.1 | `<TOOLS>\mingit\cmd\git.exe` | Version control; not on `PATH` |
| PowerShell | 5.1 | system | Build/test scripting. Note its native-argument quoting bug (below) |

The machine has no Visual Studio/C++ toolchain in use for the shim; TCC is small
and sufficient. There is no `.def` support, which is why exports are patched.
A VS2019 project (`winshim\vs2019\`) is provided for users who prefer MSVC; it is
not part of this machine's build loop.

## Build map

`Win32API.tlb` was replaced with a newer, larger EBArtSoft build (814 KB) in `7705b45`; the type declarations the module consumes are unchanged, but the checked-in dump/documentation predates it.

| Artefact | Project / script | Output |
| --- | --- | --- |
| Replacement DLL | `AsDll\WinsockVB6.vbp` | `AsDll\WinsockVB6.dll` (~32 KB) |
| Shim + probes | `winshim\build.ps1` | `winshim\out\{winshim.dll,probe.exe,testhost.exe}` |
| Shim (MSVC, optional) | `winshim\vs2019\winshim.sln` (v142, Win32) | `winshim\out\msvc\Win32\{Debug,Release}\winshim.dll` |
| Form host | `winshim\testform\winshimtest.vbp` | `winshim\testform\winshimtest.exe` |
| Feature suite | `FeatureTests\FeatureTests.vbp` | `FeatureTests\FeatureTests.exe` |
| Manifest resource | `winshim\testform\winshimtest.rc` | `winshimtest.res` |

### VB6 `/make` — the quoting trap

Windows PowerShell 5.1 does not quote native arguments. Passing a path that
contains a space (`...\Default Project\...`) unquoted makes VB6 receive a
truncated path and open `<parent>\Default.vbp` instead. The reliable form is:

```powershell
$psi = New-Object System.Diagnostics.ProcessStartInfo
$psi.FileName = '<VB6>\Vb6.exe'
$psi.Arguments = '/make "' + $projectPath + '"'
$psi.UseShellExecute = $false
$p = [System.Diagnostics.Process]::Start($psi)
$p.WaitForExit(60000)
```

Using `& $vb6 /make $path` works only when the path has no spaces; do not rely on
it. `Start-Process -ArgumentList @('/make', $path)` also fails: the array is
joined without quotes.

`/make` may show a modal dialog. For the OleDll build this is normally
**"Error accessing the system registry"** (the DLL is written, then registering
its typelib fails without elevation). The build is still valid; dismiss the
dialog and check the exit code. A real compile error also appears as a dialog
("Compile Error in File ..., Line N: ...") and the process exits non-zero.

```powershell
# generic dialog dismisser used during development
Add-Type ... 'user32' ...   # EnumChildWindows + SendMessage(BM_CLICK)
```

### TCC export patching

TCC decorates `__stdcall` exports (`_DllGetClassObject@12`) and cannot take a
`.def` file. `build.ps1` therefore rewrites the four export-name strings in the
finished DLL in place. The decorated names are longer than the plain ones, so the
replacement is NUL-padded; every RVA in the export directory stays valid. The
script then re-opens the DLL and verifies that `DllGetClassObject`,
`DllCanUnloadNow`, `DllRegisterServer` and `DllUnregisterServer` are exported
under their plain names.

### Visual Studio 2019 project

`winshim\vs2019\winshim.sln` builds the same `winshim.c` with the v142 toolset,
Win32 only, static CRT (`/MT`), no extra import libraries. Two details make it
work where the TCC script patches names:

* `winshim.def` aliases the decorated symbols to the plain COM names
  (`DllGetClassObject=_DllGetClassObject@12`, and `@0` for the three no-argument
  exports).
* `WINSHIM_NO_DLLEXPORT` is defined, so the source does not use
  `__declspec(dllexport)` and the `.def` is the only source of exports. MSVC then
  exports exactly the four plain names — no `_Name@N` aliases. The TCC build
  leaves the macro undefined, keeps `__declspec(dllexport)`, and `build.ps1`
  still finds and patches the decorated names; a rebuild after the change was
  byte-identical to the previous DLL.
* `WIN32_LEAN_AND_MEAN` is defined so the SDK does not pull in `ole2.h`; its
  `IUnknown`/`IOleObject`/`IDispatch` definitions would collide with the shim's
  own partial vtables.

The trace log is switched at compile time: `WINSHIM_NO_LOG` turns `trace_append`,
`loghr` and `logiid` into no-op macros and removes the string-building blocks.
The Release configuration defines it, so a Release `winshim.dll` never creates
`winshim.log`; Debug leaves it undefined and logs. The TCC script exposes the
same switch as `build.ps1 -NoLog` (8 KB silent DLL, verified by the script); the
default TCC build keeps the log and remains byte-identical.

The C source is C89-clean (`gcc -std=c89 -pedantic` with `WIN32_LEAN_AND_MEAN`
reports only identical macro redefinitions), so VS2019's default C mode compiles
it without `/std:c11`. The project has not been compiled on this machine — no
MSVC is installed — but it does not alter the TCC reference build.

### Resource compiler and the embedded manifest

The form test EXE embeds its manifest instead of relying on an external
`winshimtest.exe.manifest`:

```powershell
# winshimtest.rc
1 24 "winshimtest.exe.manifest"
```

```powershell
& '<TOOLS>\vb98\Wizards\RC.EXE' `
    /fo winshimtest.res winshimtest.rc
```

`winshimtest.vbp` links it with `ResFile32="winshimtest.res"`.

Two SxS traps, both hit during development:

1. **Length must be a multiple of four.** RC pads the resource with NUL bytes to
   align it. The SxS parser reads the whole resource and rejects trailing NULs as
   "Invalid Xml syntax" (event log, `SideBySide`, line 6/7). Pad the manifest text
   with spaces so `len % 4 == 0` before running RC.
2. **`<comClass>` must be on one line.** The SxS embedded-manifest parser rejected
   a two-line element with the same error.

### CRLF in `.frm`/`.vbp`

VB6 requires CRLF in form/project files. With LF-only endings the `Attribute`
lines are not parsed, the form class is not named `Form1`, and the compiler
reports `Variable not defined` in `modMain` at `Load Form1`. All `.frm`/`.vbp`
files in the repository are CRLF.

## Manifests and registration

### Per application (registry-free)

```
WebServer.exe
WebServer.exe.manifest
winshim.dll
WinsockVB6.dll
```

`app.exe.manifest` (copy and rename to `<host>.exe.manifest`) is:

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<assembly xmlns="urn:schemas-microsoft-com:asm.v1" manifestVersion="1.0">
  <assemblyIdentity type="win32" name="WinsockVB6.Winshim" version="1.0.0.0"/>
  <file name="winshim.dll">
    <comClass clsid="{248DD896-BB45-11CF-9ABC-0080C7E7B78D}" threadingModel="Apartment" progid="MSWinsock.Winsock.1" description="Winsock control (replacement)"/>
  </file>
</assembly>
```

Notes:

* Keep the XML element order and put every element on one line (embedded case) or
  at least keep `<comClass>` on one line.
* An external manifest is not always honoured after a recompile: Windows may keep
  using the previous binding for that path, and the app silently loads the real
  OCX. Creating a fresh path or embedding the manifest avoids it.
* The same reason makes `winshim.log` the acceptance criterion: if the file is
  missing next to the host, the shim did not serve the process.

### Machine wide

```bat
regsvr32 <path>\winshim.dll
regsvr32 /u %WINDIR%\SysWOW64\mswinsck.ocx   :: undo, re-register the real control
```

`DllRegisterServer` writes HKLM (elevation required):

* `HKLM\Software\Classes\CLSID\{248DD896-...}\InprocServer32` = shim path,
  `ThreadingModel = Apartment`
* `HKLM\Software\Classes\MSWinsock.Winsock\CLSID` and `.1\CLSID` = the CLSID
* `MSWinsock.Winsock\CurVer = MSWinsock.Winsock.1`

It deliberately does **not** touch `Control`, `Implemented Categories` or
`TypeLib`: those belong to the OLE control surface, which v1 does not implement,
and the real typelib registration must stay so VB6 can compile projects that
reference `mswinsck.ocx`.

For development tests, a temporary per-user registration was used instead
(`HKCU\Software\Classes\CLSID\{...}\InprocServer32`), written through the .NET
registry API. PowerShell's `Set-ItemProperty -Name '(default)'` creates a value
*literally named* `(default)`, which COM ignores; use
`$key.SetValue('', $path)`.

## Build outputs and ignore rules

`.gitignore` excludes: `FeatureTests\results\`, `FeatureTests\*.exe`,
`FeatureTests\*.vbw`, `*.log`, `*.obj`, `winshim\out\`, form-test binaries and
the user-supplied `testform2\`. The committed binaries are
`AsDll\WinsockVB6.{dll,exp,lib}` (the deliverable) and the test sources/manifest.

## Rebuilding the DLL from source

`AsDll\WinsockVB6.vbp` has `CompatibleMode=1` and `ReplCLSID=1` in
`[Compile Options]`; the project reuses its typelib/binary compatibility data
from the existing DLL. Rebuilding overwrites `AsDll\WinsockVB6.dll`,
`.exp` and `.lib`; the registry dialog appears at the end and can be dismissed.
