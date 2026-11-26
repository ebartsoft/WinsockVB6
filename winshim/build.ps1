# Builds winshim.dll, probe.exe and testhost.exe with Tiny C Compiler.
#
#   powershell -ExecutionPolicy Bypass -File build.ps1            # log on
#   powershell -ExecutionPolicy Bypass -File build.ps1 -NoLog     # log off
#
# -NoLog compiles the trace log out (WINSHIM_NO_LOG), like the VS2019 Release
# configuration. Without it the shim appends to winshim.log next to the host.
#
# TCC decorates stdcall exports (_DllGetClassObject@12) and has no .def support,
# so the script patches the export name strings in the finished DLL. The names
# are only ever shortened, padded with NULs in place, which keeps every RVAs in
# the export directory valid. This is verified at the end of the build.

param(
    [switch] $NoLog
)

$ErrorActionPreference = 'Stop'

$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$repo = Split-Path -Parent $here
$tcc  = if ($env:TCC) { $env:TCC } else { 'C:\Users\Noemie\Documents\Tools\tcc\tcc.exe' }
$out  = Join-Path $here 'out'
$vbdll = Join-Path $repo 'AsDll\WinsockVB6.dll'

if (-not (Test-Path $tcc)) { throw "TCC not found at $tcc (set the TCC environment variable)" }
New-Item -ItemType Directory -Force -Path $out | Out-Null

function Invoke-Tcc {
    param([string[]] $TccArgs)
    $prev = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    & $tcc @TccArgs 2>&1 | ForEach-Object { Write-Host "  $_" }
    $ErrorActionPreference = $prev
    if ($LASTEXITCODE -ne 0) { throw "tcc failed: $($TccArgs -join ' ')" }
}

Write-Host "== building shim (log: $(if ($NoLog) { 'off' } else { 'on' }))"
$shimArgs = @('-shared')
if ($NoLog) { $shimArgs += '-DWINSHIM_NO_LOG' }
$shimArgs += @('-o', (Join-Path $out 'winshim.dll'), (Join-Path $here 'winshim.c'))
Invoke-Tcc $shimArgs

Write-Host "== building tools"
Invoke-Tcc @('-o', (Join-Path $out 'probe.exe'),    (Join-Path $here 'probe.c'))
Invoke-Tcc @('-o', (Join-Path $out 'testhost.exe'), (Join-Path $here 'testhost.c'))

Write-Host "== patching stdcall export names"
$dll = Join-Path $out 'winshim.dll'
$bytes = [IO.File]::ReadAllBytes($dll)
$map = @(
    @('_DllGetClassObject@12', 'DllGetClassObject'),
    @('_DllCanUnloadNow@0',    'DllCanUnloadNow'),
    @('_DllRegisterServer@0',  'DllRegisterServer'),
    @('_DllUnregisterServer@0','DllUnregisterServer')
)
foreach ($pair in $map) {
    $old = [Text.Encoding]::ASCII.GetBytes($pair[0])
    $new = [Text.Encoding]::ASCII.GetBytes($pair[1])
    $done = $false
    for ($i = 0; $i -le $bytes.Length - $old.Length; $i++) {
        $match = $true
        for ($j = 0; $j -lt $old.Length; $j++) {
            if ($bytes[$i + $j] -ne $old[$j]) { $match = $false; break }
        }
        if ($match) {
            for ($j = 0; $j -lt $old.Length; $j++) { $bytes[$i + $j] = 0 }
            for ($j = 0; $j -lt $new.Length; $j++)  { $bytes[$i + $j] = $new[$j] }
            $done = $true
            break
        }
    }
    if (-not $done) { Write-Warning "export string $($pair[0]) not found" }
}
[IO.File]::WriteAllBytes($dll, $bytes)

Write-Host "== verifying exports"
$b = [IO.File]::ReadAllBytes($dll)
foreach ($pair in $map) {
    $name = [Text.Encoding]::ASCII.GetBytes($pair[1] + [char]0)
    $found = $false
    for ($i = 0; $i -le $b.Length - $name.Length; $i++) {
        $match = $true
        for ($j = 0; $j -lt $name.Length; $j++) {
            if ($b[$i + $j] -ne $name[$j]) { $match = $false; break }
        }
        if ($match) { $found = $true; break }
    }
    if ($found) { Write-Host "  ok   $($pair[1])" } else { throw "export $($pair[1]) missing" }
}

Write-Host "== verifying log mode"
$hasLog = [Text.Encoding]::ASCII.GetString($b).Contains('winshim.log')
if ($NoLog -and $hasLog) { throw 'winshim.log string still present in a -NoLog build' }
if (-not $NoLog -and -not $hasLog) { throw 'winshim.log string missing from a logging build' }
if ($NoLog) { Write-Host '  ok   log compiled out' } else { Write-Host '  ok   log enabled' }

if (Test-Path $vbdll) {
    Copy-Item $vbdll (Join-Path $out 'WinsockVB6.dll') -Force
    Write-Host "== copied WinsockVB6.dll next to the shim for testing"
} else {
    Write-Warning "AsDll\WinsockVB6.dll not found; copy it into .\out before testing"
}

Write-Host ""
Write-Host "Next:"
Write-Host "  .\out\probe.exe        # confirm the VB6 export ABI"
Write-Host "  .\out\testhost.exe     # confirm the shim creates a live object"
