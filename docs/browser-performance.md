# Fast browser testing

The browser suite keeps Playwright's assertions, isolation, retries and diagnostic tools.
Most of the avoidable cost was repeated builds, trace recording, repeated parity comparisons
and fixed sleeps. Changing the harness language does not remove those costs.

## Local runs

Build the runtime packages after changing product code:

```powershell
bun run build
bun run --cwd src/ui build
Set-Location viewers/pptx
bun run build:packages
```

Then prepare production demos and run a selected spec:

```powershell
# Choose an offset that is unused by other worktrees.
$env:PPTX_E2E_PORT_OFFSET = '300'
bun run e2e:fast -- table-styling.spec.ts --project=react
```

`e2e:fast` regenerates deterministic fixtures and builds only stale demo bundles. Its cache
hash includes actual source, fixture and built dependency bytes, including uncommitted edits.
Preview runs reject stale builds. Cache stamps live in private `node_modules/.cache` folders.
`bun run e2e` still supports development servers; their dependency cache is retained unless
`PPTX_E2E_VITE_FORCE=1` explicitly requests a rebuild.

For repeated debugging, keep the preview servers warm:

```powershell
$env:PPTX_E2E_SERVE = 'preview'
bun e2e-servers.ts start
bun run e2e -- table-styling.spec.ts --project=react
# Stop the servers when finished. Restart them after rebuilding dependencies.
bun e2e-servers.ts stop
```

React's parity tests drive every framework. To exercise each framework's ordinary tests too,
pass all five projects or set `PPTX_E2E_ALL_BINDINGS=1`.

`OOXML_E2E_WORKERS` overrides the shared worker budget with a positive integer. The default
is half the available logical CPUs, capped at eight locally and two on Linux CI runners.
PowerPoint also retains the existing `PPTX_E2E_WORKERS` override. Teams keeps one worker
because its suites share server state. Use `--workers=N` for a one-off comparison.

Local tracing defaults to off. Use `--trace=on` when diagnosing a failure. CI records the
first retry with `on-first-retry`, retaining diagnostic traces without recording every pass.

`bun run verify --browser` preserves the CI planner's selected spec files and includes the
PowerPoint headless host projects. Browser steps run exclusively, after other running checks
drain, so their workers do not compete with typechecks or unit tests. `CI=1` still enforces
strict checks locally, while `OOXML_LOCAL_VERIFY=1` preserves local GPU and worker settings.
On Windows, verification uses Git Bash rather than WSL; `OOXML_BASH` can select its executable.

## CI scheduling and coverage

- Ordinary PowerPoint browser demos are built once and distributed to all shards. Fixture
  generation is checked once before reuse. The packaged smoke retains its own shipping build
  without the suite's injected test handle.
- Core package checks and PowerPoint type/package checks run in separate required jobs, so
  browsers can start when their builds are ready. `ci-success` still requires those checks.
- `@cross-binding` marks scenarios that already replay through `acrossFrameworks`. They run
  in React when React is selected. A standalone non-React run still runs its comparisons.
  The contract checker rejects the tag on tests using the ordinary page fixture or creating
  independent browser pages. Ordinary single-framework scenarios remain enabled everywhere.
- `PPTX_E2E_SHARD=1/14` assigns files before Playwright imports them. All shards receive the
  same immutable timing baseline. Longest-first assignment uses per-framework durations and
  combines ordinary/headless host costs. New files get a median estimate; missing history
  falls back to equal file weights. Timing history affects ordering, never whether a file runs.
- Timing reports are merged with previous history after a successful matrix. Failed runs
  retain their evidence. The existing retry policy and hosted Chromium stability flags remain.

`PPTX_E2E_SPECS` accepts the planner's space-separated spec filenames. Do not combine the custom
shard variable with Playwright's `--shard`: applying both would omit tests. A missing selected
file fails explicitly. To audit parity redundancy, set `PPTX_E2E_REFERENCE_COVERS_PARITY=0`
and compare `playwright test --list --reporter=json` output with the normal selection.

## Measurements on 28 October 2026

Machine: Ryzen 7 9800X3D, 16 logical CPUs, approximately 64 GB RAM, Windows, Bun 1.4.2.
These are single-run observations with other workstation processes present, not a statistical
benchmark or a promise for full-suite/CI duration. Every table comparison used the same 20
selected tests across all five frameworks and warm production preview servers.

| Configuration                                         | Elapsed |
| ----------------------------------------------------- | ------: |
| Original table spec, retained traces, 4 workers       |  48.8 s |
| Updated waits/parallelism, retained traces, 4 workers |  34.5 s |
| Updated waits/parallelism, traces off, 4 workers      |  20.6 s |
| Updated waits/parallelism, traces off, 8 workers      |  15.6 s |
| Updated waits/parallelism, traces off, 12 workers     |  13.4 s |

Eight workers remain the conservative automatic cap; twelve helped this small sample but is
not established as optimal for heavy video, canvas or collaboration tests. A cached five-demo
preparation took 10.4 seconds including fixture generation. The Teams reminder test passed in
4.3 seconds after replacing real timer waits with Playwright's clock.

The all-framework collection audit reduced 5,955 executions to 5,431: all 524 removed
executions had the cross-binding tag. Shard coverage matched every collected file exactly
once and retained each headless collaboration host run. Full affected table, selection and
SmartArt files passed: 240 tests in 1.8 minutes at eight workers. Comment lifecycle coverage
also passed in all five frameworks. CI wall-clock improvement needs a hosted run; existing
shard counts remain unchanged, so runner queue time is still a separate constraint.

## Rust assessment

[cargo-nextest](https://nexte.st/) runs Rust tests; it does not execute these TypeScript browser
specs. [headless_chrome](https://github.com/rust-headless-chrome/rust-headless-chrome) drives
Chrome through CDP, but replacing Playwright with it would require rewriting the suite and
its assertion, fixture, retry and reporting integration. It would still pay browser rendering,
application startup and timer costs. No Rust harness is introduced without evidence that
JavaScript orchestration is the bottleneck.

The next measurements should use repeated warm runs of representative light and heavy specs,
then compare hosted shard durations and queue time. Use the timing artifacts to identify the
next slow files before replacing more waits or increasing parallelism. Timer-driven tests can
use [Playwright Clock](https://playwright.dev/docs/clock); real media playback tests must keep
real elapsed-time assertions.
