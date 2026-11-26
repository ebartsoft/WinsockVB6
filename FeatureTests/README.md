# modWinsock2 feature & performance comparison

This folder is a standalone VB6 project that exercises **every member of
`IMSWinsockControl`** through `..\modWinsock2.bas` and through the real
`MSWINSCK.OCX`, then reports where the two implementations behave differently
and how fast each one is.

The production module is still `..\modWinsock2.bas`; this project only links to
it. `modWinsock.bas` and `modUtils.bas` are **not** included, so the build also
proves that the new module is self-contained.

The harness does not modify the module. The module itself now carries the dispatch layer and two parity fixes found by these tests; see `..\Papers\05-discoveries.md`.

## Files

| File | Purpose |
| --- | --- |
| `FeatureTests.vbp` | Project file; references `..\Win32API.tlb` and `stdole2.tlb` |
| `frmFeatureTests.frm` | UI: summary, result list, detail pane, run/save buttons, `/autorun` |
| `modTestEngine.bas` | Scenarios, comparison logic, performance suite, report writers |
| `modTestUtil.bas` | QPC timing, message pump, diff/HTML/CSV helpers, file output |
| `clsEndpoint.cls` | Test-only `WithEvents` wrapper; drives either `New Winsock` or `New_Winsock2` |
| `clsRun.cls` | Result of one implementation for one scenario (trace, snapshot, failure) |
| `clsResult.cls` | Paired result: W2 run, OCX run, status, first divergence |
| `clsPerfRow.cls` | One performance comparison row |
| `results/` | Generated reports (HTML, Markdown, CSV, raw log) |

The wrapper exists only for event handling in the harness. The production
implementation uses no class module for the Winsock object lifecycle.

## Build and run

```bat
Vb6.exe /make FeatureTests.vbp /out build.log
FeatureTests.exe /autorun
```

`/autorun` runs everything, writes the reports to `results\`, and exits. Without
it the form stays open and the buttons drive the same engine.

| Switch | Effect |
| --- | --- |
| `/autorun` | Run, save reports, exit |
| `/func` | Functional scenarios only |
| `/perf` | Performance suite only |
| `/only:<id>` | Single scenario, e.g. `/only:tcp_str` |

Scenario ids: `defaults`, `props`, `tcp_str`, `tcp_bytes`, `tcp_multi`,
`tcp_large`, `udp`, `dns`, `errors`, `lifecycle`, `latebinding`, `interop`.

If `MSWINSCK.OCX` is not registered, `New Winsock` fails, the oracle is marked
unavailable, and the scenarios still run against the replacement (reported as
`W2_ONLY`).

The `stdole2.tlb` reference uses the 32-bit path
`%WINDIR%\SysWOW64\stdole2.tlb`; adjust it if your VB6 installation keeps the
type library elsewhere.

## How the comparison works

Each scenario runs twice, once per implementation, and records:

* an **event trace** - every event with its arguments (`Connect`, `DataArrival`,
  `SendProgress`, `SendComplete`, `Close`, `Error`, `ConnectionRequest`);
* a **property snapshot** - `Protocol`, `State`, `SocketHandle`, `LocalPort`,
  `LocalIP`, `LocalHostName`, `RemotePort`, `RemoteHost`, `RemoteHostIP`,
  `BytesReceived`;
* pass/fail assertions on the payloads and on the reported error numbers.

Results are classified as:

| Status | Meaning |
| --- | --- |
| `PASS` | Both pass and the event traces are identical |
| `DIFF` | Both pass, but the traces differ (first divergence is shown) |
| `FAIL` | At least one implementation failed an assertion |
| `EXPECTED` | Known v1 gap, deliberately not counted as a failure |
| `W2_ONLY` | Oracle unavailable |

## Scenarios

| Group | Scenario | What it covers |
| --- | --- | --- |
| Defaults | Initial properties | Default `Protocol`, `State`, `SocketHandle`, `LocalPort` |
| Properties | Protocol/remote round-trip | `Protocol`, `RemoteHost`, `RemotePort`, `LocalHostName`, `LocalIP`, out-of-range setter behaviour |
| TCP Client | String echo | `Connect`, `SendData` string, `GetData` string, `PeekData` string, second message |
| TCP Client | Byte array + PeekData | `SendData`/`GetData`/`PeekData` with `Byte()` including a 250 byte |
| TCP Server | Multiple clients | `Listen`, two `ConnectionRequest`s, two `Accept`s, per-connection echo |
| TCP Client | Close and reuse control | `Close` leaves `State = sckClosed` and the same control can connect again |
| TCP Data | 1 MB transfer | Large `SendData` and byte counting through repeated `GetData` |
| UDP | Two datagrams + echo | `Bind`, `SendData`, `DataArrival`, `RemoteHostIP`/`RemotePort`, reply to the learned peer |
| DNS | localhost resolution | `Connect` by host name, `State` right after the call |
| Errors | Closed/refused/bind in use | `GetData`/`SendData` when closed, connection refused, port already in use |
| Lifetime | Create/release cycles | 50 create/release cycles |
| Late binding | CallByName/IDispatch | `CallByName(sock, "State", VbGet)` matches the OCX |
| Interop | OCX server + W2 client | Real OCX listener serving the replacement |
| Interop | W2 server + OCX client | Replacement listener serving the real OCX |

### Reliability tiers

Beyond the API-compatibility scenarios above, the suite drives both
implementations through payload integrity, concurrency, stress and adversarial
inputs. Every check compares the bytes that actually arrived, not just the event
sequence.

| Group | Scenario | What it proves |
| --- | --- | --- |
| Integrity | 4 MB mixed chunks | 4 MiB pattern sent in 1/3/17/1024/65535/65537-byte chunks arrives byte-exact |
| Integrity | Slice/peek reads | `GetData` with `MaxLen` slicing plus `PeekData`; no byte lost or duplicated, peek does not consume |
| Integrity | UDP datagram bounds | 1, 2, 1472, 8192 and 65507-byte datagrams: one `GetData` returns exactly one datagram, content intact |
| Integrity | Edge byte patterns | all `0x00`, all `0xFF`, NUL/CR/LF mixes and a counter pattern |
| Integrity | ANSI string bytes | `SendData` with characters above 127; byte count and content reported per implementation |
| Integrity | 1-byte/64 KB interleave | worst-case chunk interleaving against the read/write buffer compaction paths |
| Concurrency | 32 clients, no crosstalk | 32 live sockets on one listener, each echoes a unique payload; catches handle/slot mis-routing |
| Concurrency | 3 listeners x 2 clients | six sockets across three listeners, correct pairing |
| Concurrency | TCP and UDP together | 4 TCP + 4 UDP sockets live at once |
| Concurrency | 200 connect/close cycles | sequential churn, then asserts the module reports zero objects, sockets and pending accepts |
| Stress | 8 MB stream | integrity plus MB/s |
| Stress | 20k message round trips | throughput and per-message latency distribution |
| Stress | 4k datagram burst | paced burst, delivery count and loss |
| State machine | Wrong-order calls | `Connect` with no host, invalid protocol, double `Close`, bogus `Accept`, `Connect`/`Listen` while connected |
| State machine | Close inside DataArrival | closing from inside an event must not crash or corrupt state |
| State machine | Peer reset mid-transfer | `SO_LINGER 0` peer kills a live transfer; the socket must surface error/close and stay reusable |
| State machine | 4 KB window, 1 MB send | forced partial sends, all bytes must still arrive intact |
| Interop | 4 MB integrity both directions | OCX server to W2 client and the reverse, byte-exact |

`AboutBox` is modal, so it is not invoked by the automated run.

## Performance suite

Measured with `QueryPerformanceCounter` (the report prints the detected clock):

| Operation | What is timed |
| --- | --- |
| Create/release object | 500 x create + init + release |
| TCP round-trip 64 B | 100 echo round trips over one connection, same W2 server for both clients |
| TCP throughput 1 MB | Five 1 MB sends averaged, reported per transfer |
| UDP round-trip 64 B | 100 datagram round trips to the same W2 echo receiver |

The OCX is native C++/COM and the replacement is pure VB6, so the numbers are
indicative rather than a fair fight. Each round-trip row includes one
message-pump wait per echo, which is why TCP round trips land close together.

The project builds as VB6 P-code (the default, and what the other projects in
this repository use). A native-code build was measured as well and was not
faster for these scenarios on this machine, so the table above is P-code.

## Current results (reference run)

```
PASS=32  DIFF=0  FAIL=0  EXPECTED/SKIP=1  Scenarios=33  (about 17 s)
```

| Operation | modWinsock2 | OCX | W2/OCX |
| --- | ---: | ---: | ---: |
| Create/release object (500) | 5.4 ms | 51.8 ms | 0.10x |
| TCP round-trip 64 B (200) | 30.0 ms | 31.5 ms | 0.95x |
| TCP throughput 1 MB (20 averaged) | 4.7 ms | 4.4 ms | 1.08x |
| UDP round-trip 64 B (200) | 37.9 ms | 31.0 ms | 1.22x |

Reliability numbers from the same run: 4 MiB and 8 MiB streams byte-exact
(203 MB/s), 20,000 round trips at ~6,800/s, 4,000 datagrams with 0% loss, 200
connect/close cycles in 0.16 s with zero leaked objects, sockets or pending
accepts.

### What the new tiers already found

* **A real data-loss bug in `SendData`:** the byte count came from `Len(s)`, so
  when the ANSI conversion produced fewer bytes than characters the send read
  past the converted array. `n` now comes from the converted array. The ANSI
  scenario also shows the two implementations disagree on *how* unmappable
  characters are converted (recorded in the report, not scored).
* **MSWINSCK cannot be driven out of order:** `Connect` while connected,
  `Listen` while connected and similar calls crash it, so the wrong-order
  scenario runs against `modWinsock2` only and is classified `EXPECTED`.
* **MSWINSCK drops most of a 4,000-datagram burst** (around 96% loss on this
  machine) while `modWinsock2` delivers all of them. Recorded, not judged — the
  oracle is not the reliability target.
* **Kernel handles grow on both sides** during connect/close churn (TIME_WAIT),
  so the harness asserts on the module's own counters (`W2_Diagnostics`)
  instead of process handle counts.

`modWinsock2` exposes a small read-only hook for exactly this,
`W2_Diagnostics()`, which reports live objects, used state slots, open sockets,
pending accepts, queued datagrams and pending DNS requests.

## Differences the harness currently reports

The suite reports no `DIFF` rows. What remains is one classified gap and two
observations the harness records but does not score:

1. **`Error` Source** - the OCX reports its own module path
   (`<OCX path>`), the replacement reports `MSWinsock`. `Number`,
   `Description` and `Scode` all match.
2. **UDP read timing** - the OCX discards a datagram that is not read inside the
   `DataArrival` handler; the replacement keeps it buffered. The harness reads
   inside the event, which is the documented pattern for both.

### OCX defects that the harness must avoid

`MSWINSCK.OCX 6.1.97.82` raises an access violation (`0xC0000005`) in three
situations that the replacement handles gracefully:

* reading `BytesReceived` on a socket that has already been `Close`d;
* calling `GetData` when no connection exists;
* calling `Bind` on a port that is already in use.

Those probes run only against `modWinsock2`; the OCX run records them as skipped
in the result notes, and the property snapshot omits `BytesReceived` once the
socket is closed.

## Differential matrix (`/matrix`)

A second, much larger sweep compares individual behaviours one cell at a time.
Each cell runs in **its own process**, so an MSWINSCK crash, hang or modal
dialog is recorded as an observation instead of aborting the run.

```bat
FeatureTests.exe /matrix              :: all 199 cells, about 9 minutes
FeatureTests.exe /matrix:pg           :: only cells whose id contains "pg"
FeatureTests.exe /matrix:abort        :: only the abort family
```

Output goes to `results\matrix_latest.md` and `results\matrix_latest.csv`
(plus timestamped copies). Each row is one cell: verdict, the oracle's process
outcome, and the first differences between the two records.

Verdicts: `MATCH`, `DIVERGENCE`, `ORACLE_EXIT:/DIALOG:/TIMEOUT` (the oracle
crashed, hung or showed a dialog — recorded, not judged), `W2_BROKEN`.

Cell ids are `<family>:<variant>[:<arg>]`, for example
`pg:connected:LocalPort`, `send:emptyArr`, `abort:rstThenSend`,
`dns:concurrent8`. Families: `pg`, `ps`, `conn`, `listen`, `accept`, `bind`,
`send`, `get`, `close`, `dns`, `evt`, `abort`, `conc`.

Individual cells can be run directly, which is the fastest way to inspect one
behaviour:

```bat
FeatureTests.exe /cell:send:emptyArr /impl:w2  /out:<out>\w2.txt
FeatureTests.exe /cell:send:emptyArr /impl:ocx /out:<out>\ocx.txt
```

Two known limitations in this first pass: `impl=ocx` children of the `ps`,
`accept` and `listen` families sometimes exit without writing their record
(those rows show `ORACLE_NO_OUTPUT`), and the `get` family currently measures
the "no data" path rather than the MaxLen/type matrix.

## Reports

`Save Reports` (and `/autorun`) writes to `results\`: 

| File | Content |
| --- | --- |
| `report_<stamp>.html` | Self-contained HTML: summary, results, diffs, performance, notes |
| `report_<stamp>.md` | Same in Markdown |
| `report_<stamp>.csv` | Scenario and performance rows for spreadsheets |
| `rawlog_<stamp>.log` | Full narrative: assertions, both event traces, snapshots, diagnostics |
| `latest.*` | Copies of the most recent run |
