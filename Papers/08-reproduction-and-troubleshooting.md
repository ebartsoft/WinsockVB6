# 08 — Reproduction and troubleshooting

All commands assume the project root:

```powershell
$w = '<repo>'
$vb6 = '<VB6>\Vb6.exe'
```

## Build order

1. `AsDll\WinsockVB6.dll` (the implementation).
2. `winshim\out\winshim.dll` (identity + OLE layer) and copy the DLL to `out\`.
3. `winshim\testform\winshimtest.exe` (form host).
4. `FeatureTests\FeatureTests.exe` (suite).
5. Optional: `winshim\testform\winshimtest.res` (only if the manifest changed).

### 1. Replacement DLL

```powershell
$psi = New-Object System.Diagnostics.ProcessStartInfo
$psi.FileName = $vb6
$psi.Arguments = '/make "' + (Join-Path $w 'AsDll\WinsockVB6.vbp') + '"'
$psi.UseShellExecute = $false
$p = [System.Diagnostics.Process]::Start($psi); $p.WaitForExit(60000)
"exit=$($p.ExitCode)"
```

* `AsDll\WinsockVB6.vbp` contains `Common.bas` (renamed from `modDllMain.bas`
  in `7705b45`), `modWinsock2.bas` and `Dummy.cls`; the export lives in
  `Common.bas`.

### Building without Win32API.tlb (parked)

`modWinsock2.bas` reserves a `#If NO_TYPELIB` area (`PUT DATA HERE`) for local
Win32 declarations and the four supporting UDTs. It is empty; the work is parked
because seven COM types (`Winsock`, `DMSWinsockControlEvents`, `ITypeInfo`,
`IConnectionPoint*`) cannot be declared in a `.bas`. See
`06-state-of-development.md` P1 before attempting it.

A dialog **"Error accessing the system registry"** at the end is expected (the
typ elib registration needs elevation); the DLL is written. A compile error
appears as "Compile Error in File …, Line N" and exits non-zero.

### 2. Shim

```powershell
powershell -ExecutionPolicy Bypass -File (Join-Path $w 'winshim\build.ps1')
```

It builds `winshim.dll`, `probe.exe`, `testhost.exe`, patches the export names,
verifies them, and copies `..\AsDll\WinsockVB6.dll` into `winshim\out\`.

Quick check:

```powershell
& (Join-Path $w 'winshim\out\testhost.exe')
# DllGetClassObject ... hr=0
# CreateInstance ... hr=0
# get_State hr=0 state=0
# get_Protocol hr=0 protocol=0
# QI IConnectionPointContainer hr=0
```

### 3. Form test

Rebuild the resource only when the manifest changes (and pad to 4 bytes):

```powershell
$res = '<TOOLS>\vb98\Wizards\RC.EXE'
& $res /fo (Join-Path $w 'winshim\testform\winshimtest.res') (Join-Path $w 'winshim\testform\winshimtest.rc')
```

Compile and run:

```powershell
$t = Join-Path $w 'winshim\testform'
$psi = New-Object System.Diagnostics.ProcessStartInfo
$psi.FileName = $vb6
$psi.Arguments = '/make "' + (Join-Path $t 'winshimtest.vbp') + '"'
$psi.UseShellExecute = $false
$p = [System.Diagnostics.Process]::Start($psi); $p.WaitForExit(60000)

Copy-Item (Join-Path $w 'winshim\out\winshim.dll') $t -Force
Copy-Item (Join-Path $w 'winshim\out\WinsockVB6.dll') $t -Force
Remove-Item (Join-Path $t 'winshim.log'),(Join-Path $t 'form.log') -ErrorAction SilentlyContinue
$app = Start-Process (Join-Path $t 'winshimtest.exe') -WorkingDirectory $t -PassThru
$app.WaitForExit(20000)
Get-Content (Join-Path $t 'form.log')
```

Expected: `Form_Load …`, `listen state=2`, `server connreq 1`,
`accepted state=7`, `server got 'hello'`, `client got 'world'`, `closed 0/0` —
and `winshim.log` must exist (proof the shim served the process).

### 4. Feature suite

```powershell
$f = Join-Path $w 'FeatureTests'
$psi = New-Object System.Diagnostics.ProcessStartInfo
$psi.FileName = $vb6
$psi.Arguments = '/make "' + (Join-Path $f 'FeatureTests.vbp') + '"'
$psi.UseShellExecute = $false
$p = [System.Diagnostics.Process]::Start($psi); $p.WaitForExit(60000)

$p = Start-Process (Join-Path $f 'FeatureTests.exe') -ArgumentList '/autorun' -WorkingDirectory $f -PassThru
$p.WaitForExit(180000)
Get-Content (Join-Path $f 'results\latest.md') | Select-Object -First 6
```

Expected: `PASS=32 DIFF=0 FAIL=0 EXPECTED/SKIP=1 Scenarios=33`.

### 5. Matrix (deliberate, ~9 minutes)

```bat
FeatureTests.exe /matrix
FeatureTests.exe /matrix:send
FeatureTests.exe /cell:send:emptyArr /impl:w2 /out:<out>\w2.txt
```

Do not run it casually: each cell starts a child process and the full sweep takes
several minutes.

## Verifying the real application without touching `IRL-Test`

Work on a copy:

```powershell
$src = Join-Path $w 'IRL-Test'
$dst = Join-Path $env:TEMP 'irl-repro'
Remove-Item $dst -Recurse -Force -ErrorAction SilentlyContinue
New-Item -ItemType Directory $dst | Out-Null
Copy-Item (Join-Path $src 'WebServer.exe'),(Join-Path $src 'WebServer.config'),(Join-Path $src 'WinsockVB6.dll') $dst
Copy-Item (Join-Path $src 'www') $dst -Recurse
Copy-Item (Join-Path $w 'winshim\out\winshim.dll') $dst
Copy-Item (Join-Path $w 'AsDll\WinsockVB6.dll') $dst
```

Set `USE_MSWINSOCK = True` / `USE_VBWINSOCK = False` in the copy's
`WebServer.config` (the COM path is the one that exercises the shim).

Then either place `WebServer.exe.manifest` next to it, or use a temporary
per-user registration:

```powershell
$clsid = '{248DD896-BB45-11CF-9ABC-0080C7E7B78D}'
$k = [Microsoft.Win32.Registry]::CurrentUser.CreateSubKey("Software\Classes\CLSID\$clsid\InprocServer32")
$k.SetValue('', (Join-Path $dst 'winshim.dll')); $k.SetValue('ThreadingModel','Apartment'); $k.Close()
# run, then always clean up:
[Microsoft.Win32.Registry]::CurrentUser.DeleteSubKeyTree("Software\Classes\CLSID\$clsid", $false)
```

Run and observe:

```powershell
$p = Start-Process (Join-Path $dst 'WebServer.exe') -WorkingDirectory $dst -PassThru
Start-Sleep 4
if ($p.HasExited) { "exited $($p.ExitCode)" } else {
  Get-NetTCPConnection -OwningProcess $p.Id
  $p.Kill()
}
Get-Content (Join-Path $dst 'winshim.log')
```

Expected: still running, `0.0.0.0:80 Listen`, and a log that ends with
`QI-unhandled` for `IPersistStreamInit`/`IPersistPropertyBag` (code-created
object: persistence refused by design).

## Troubleshooting

| Symptom | Likely cause | Action |
| --- | --- | --- |
| App starts then exits immediately, no dialog | Persistence handshake accepted for a code-created object | Check `winshim.log` tail: it should end at the persistence QIs with `QI-unhandled`. If `GetSizeMax` is last, the site gate is missing/broken |
| `winshim.log` is not created next to the host | Manifest ignored (or invalid `winshim.dll`), so the real OCX served the process | Verify the manifest name matches the EXE, DLL exports are patched, and the folder contains `winshim.dll`; try an embedded manifest or a fresh path |
| Form fails with **372** "Failed to load control" | `IPersistStreamInit` missing (QI refused) | Persistence must be offered once a client site exists; check `c->site` gate |
| **80010108** Automation error / object disconnected | A VB method was called through C and raised | Keep dispatch in VB; never wrap VB methods in C |
| **440** Automation error | EXCEPINFO returned without the VB HRESULT convention | Return `&H800A0000 Or number` from `W2_Invoke`'s `Failed` handler |
| SxS startup error or "Invalid Xml syntax" in the event log | Embedded manifest length not a multiple of 4, or `<comClass>` split over lines | Pad the manifest; one line per element; rebuild the `.res` |
| VB6 opens `<parent>\Default.vbp` | Unquoted native argument (PowerShell 5.1) | Use `ProcessStartInfo.Arguments` with quotes |
| `Variable not defined` at `Load Form1` | `.frm`/`.vbp` saved with LF endings | Convert to CRLF and rebuild |
| "Error accessing the system registry" after `/make` | Normal for the OleDll build without elevation | Dismiss; the DLL is written |
| Form still shows the old control | Manifest not applied (or EXE started from a different folder) | Check `winshim.log`; keep the EXE name stable; avoid new names (UAC/firewall) |
| Port 80 already in use | A previous test server is still running | `Get-NetTCPConnection -LocalPort 80`, stop the stray process |
| `ORACLE_NO_OUTPUT` matrix cells | Known harness limitation (`ps`/`accept`/`listen`, `impl=ocx`) | Not a replacement bug; fix the harness before judging those cells |

## Cleanup checklist

* Delete temporary per-user CLSID keys (both `Software\Classes\CLSID` and
  `Software\Classes\WOW6432Node\CLSID` paths were used during experiments).
* Remove `%TEMP%\irl-repro` and other copies.
* Do not leave `regsvr32` registration behind on a shared machine.
* Generated logs (`winshim.log`, `form.log`, `w2bind.log`) are git-ignored but
  should not be copied into the repository.
