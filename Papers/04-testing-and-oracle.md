# 04 — Testing method and the oracle

## Philosophy

The replacement is judged **differentially against the real control** on the same
machine. There is no hand-written expectation table: each scenario runs twice —
once with `New_Winsock2` (W2) and once with `New Winsock` (the OCX) — and the two
records are compared. A behaviour is correct when it matches the OCX, or when the
OCX itself is demonstrably broken and the difference is recorded as such.

Two consequences matter:

* The OCX is a **behavioural oracle, not a reliability target**. Where it crashes
  or drops data, the harness records the observation and does not "fix" the
  replacement to match the bug.
* Anything not exercised by a scenario is not proven. The suite covers the
  documented surface plus integrity, concurrency, stress and adversarial cases;
  design-time hosting is explicitly outside it.

## Harness architecture

`FeatureTests\FeatureTests.vbp` is a standalone VB6 project that links
`..\modWinsock2.bas` (it does not include `modWinsock.bas`/`modUtils.bas`, which
also proves the module is self-contained).

| File | Role |
| --- | --- |
| `frmFeatureTests.frm` | UI + `/autorun` entry point |
| `modTestEngine.bas` | Scenarios, comparison, performance rows, report writers |
| `modTestUtil.bas` | QPC timing, message pump, diff/HTML/CSV helpers |
| `clsEndpoint.cls` | `WithEvents` wrapper that drives either implementation |
| `clsRun.cls`, `clsResult.cls`, `clsPerfRow.cls`, `clsCell.cls` | Result records |
| `modMatrix.bas`, `modMatrixCells*.bas` | The 199-cell differential matrix |

Each functional scenario produces, per implementation:

* an **event trace** (every event and its arguments),
* a **property snapshot** (`Protocol`, `State`, `SocketHandle`, `LocalPort`,
  `LocalIP`, `LocalHostName`, `RemotePort`, `RemoteHost`, `RemoteHostIP`,
  `BytesReceived`),
* **assertions** on payloads, byte counts and error numbers.

Traces are compared per endpoint; `DataArrival` byte counts are ignored in the
comparison (they depend on buffering) and payload totals are asserted separately.

## Verdicts

| Verdict | Meaning |
| --- | --- |
| `PASS` | Both implementations pass and the traces are identical |
| `DIFF` | Both pass but traces differ (first divergence shown) |
| `FAIL` | At least one assertion failed |
| `EXPECTED` | Known/deliberate gap or an oracle limitation, not counted as failure |
| `W2_ONLY` | Oracle unavailable (OCX not registered) |

## Scenario coverage (33 scenarios)

**Core API** — defaults, property round-trip, TCP client string echo, byte arrays
with `PeekData`, TCP server with multiple clients, close-and-reuse, 1 MB
transfer, UDP datagrams, DNS by name, error paths (closed/refused/port in use),
create/release lifetime, `CallByName`.

**Reliability tiers** (all compare bytes that actually arrived):

| Group | What it proves |
| --- | --- |
| Integrity | 4 MiB mixed chunks; slice/peek reads; UDP datagram bounds (1…65507); edge byte patterns; ANSI bytes; 1-byte/64 KB interleave |
| Concurrency | 32 clients without crosstalk; 3 listeners × 2 clients; TCP+UDP together; 200 connect/close cycles with zero leaked objects |
| Stress | 8 MB stream; 20k round trips; 4k datagram burst |
| State machine | wrong-order calls; `Close` inside `DataArrival`; peer reset mid-transfer; 4 KB window with 1 MB send |
| Interop | OCX server ↔ W2 client and the reverse, including 4 MB integrity |

`AboutBox` is modal and is not invoked automatically.

## Differential matrix (`/matrix`)

A second, much larger sweep compares one behaviour per cell. **Every cell runs in
its own child process**, so an OCX crash, hang or modal dialog is recorded
instead of aborting the run.

```bat
FeatureTests.exe /matrix            :: 199 cells, about 9 minutes
FeatureTests.exe /matrix:pg         :: filter by id substring
FeatureTests.exe /cell:send:emptyArr /impl:ocx /out:<out>\ocx.txt
```

Cell ids are `<family>:<variant>[:<arg>]`. Families: `pg`, `ps`, `conn`, `listen`,
`accept`, `bind`, `send`, `get`, `close`, `dns`, `evt`, `abort`, `conc`.

Verdicts: `MATCH`, `DIVERGENCE`, `ORACLE_EXIT`/`DIALOG`/`TIMEOUT` (recorded, not
judged), `ORACLE_NO_OUTPUT`, and `W2_BROKEN`.

Reference run (2026-11-24, 199 cells): **MATCH 101, DIVERGENCE 42,
ORACLE-LIMITED 56, W2_BROKEN 0**. The CSV/Markdown are run artefacts under
`results\` and are not committed; the divergence list is summarised in
`06-state-of-development.md`.

Known first-pass defects of the matrix:

* `impl=ocx` children of the `ps`, `accept` and `listen` families sometimes exit
  without writing their record (`ORACLE_NO_OUTPUT`), so those 56 cells carry no
  signal yet.
* The `get` family currently measures the "no data" path rather than the
  `MaxLen`/type matrix.

## Performance suite

Measured with `QueryPerformanceCounter`; the report prints the detected clock.
Rows: object create/release (500), TCP round-trip 64 B (200, one pump wait per
echo), TCP throughput 1 MB (twenty sends averaged), UDP round-trip 64 B (200).
The OCX is native C++/COM, the replacement is P-code VB6, so ratios are
indicative.

Reference run (2026-11-25 01:55, same host, after the source reformat):

| Operation | W2 | OCX | ratio |
| --- | ---: | ---: | ---: |
| Create/release 500 | 5.15 ms | 52.67 ms | 0.10× |
| TCP round-trip 64 B ×200 | 29.59 ms | 30.06 ms | 0.98× |
| TCP throughput 1 MB | 4.72 ms | 5.15 ms | 0.91× |
| UDP round-trip 64 B ×200 | 39.03 ms | 29.92 ms | 1.30× |

## Report artefacts

`/autorun` (and "Save Reports") writes to `results\`:

| File | Content |
| --- | --- |
| `report_<stamp>.html` | Self-contained report: summary, results, diffs, performance, notes |
| `report_<stamp>.md` | Same in Markdown |
| `report_<stamp>.csv` | Scenario and performance rows |
| `rawlog_<stamp>.log` | Full narrative: assertions, both traces, snapshots, diagnostics |
| `latest.*` | Copies of the most recent run |

## Oracle defects the harness must avoid

`MSWINSCK.OCX 6.01.9782` raises an access violation (`0xC0000005`) in three
situations the replacement handles:

* reading `BytesReceived` after `Close`;
* `GetData` when no connection exists;
* `Bind` to a port already in use.

Those probes run only against W2; the OCX run records them as skipped, and the
snapshot omits `BytesReceived` once the socket is closed. The same applies to
wrong-order calls (`Connect`/`Listen` while connected), which crash the OCX.

## Interpretation rules

* A `DIFF` is a finding, not necessarily a bug: check whether the OCX behaviour is
  documented and sane.
* `ORACLE_*` means "no signal"; never count it as a pass.
* Payload integrity is asserted independently of traces, so a trace `PASS` with a
  corrupt byte stream is still a `FAIL`.
* Performance numbers are only comparable within one run (same host, same load).
