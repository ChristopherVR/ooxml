# Security regression corpus

Every security defect fixed in this repository keeps a case here, so it stays fixed. A case is the smallest generated input that shows the defect's class, fed to the public entry point the defect was reached through, with the observation a safe build returns.

`security-corpus.test.ts` runs in the core unit suite. It bundles each module under `cases/` once, then runs every case in a child process of its own:

- the process heap is capped at the case's `limits.heapMb`, so an input that exhausts memory ends that process and fails as `over-heap`, never the test runner;
- `run` must finish within `limits.ms`, or the process is killed and the case fails as `over-time`;
- `fetch`, `WebSocket`, `EventSource`, sockets, HTTP and DNS are replaced before the case loads, so any network attempt fails the case;
- the observation must equal `expected` exactly, which is where a refusal code, an inert value or "no script-capable markup" is asserted.

Run it from `src/core`:

```sh
bunx vitest run test-support/security-corpus
SECURITY_CORPUS_REPORT=1 bunx vitest run test-support/security-corpus --silent=false
```

The second form prints each case's time and heap, for choosing its limits.

## Adding a case

Add the case in the commit that fixes the defect, to the module of its area (or a new module: every `cases/*.ts` exporting `cases` is picked up).

- **Generate the input in `run`.** No third-party samples and no binary fixtures.
- **Use the smallest input that shows the class.** For a resource defect, pick a size the unfixed code takes at least three times the limit to process, and the fixed code at most a third of it. Measure both with the report flag before you set the limits.
- **Name the safe behaviour, not the attack.** The test name describes the defect class ("markup in a part name never reaches element ids"). Keep payloads inert: an empty `javascript:` URL, a `<p>` in a data URL, a loopback or `.test` host.
- **Only fixed defects go here.** An unpatched finding stays out of the public tree until its fix lands.
- **Check that it fails before the fix.** Run the case against the commit before the fix, and say in the commit message that you did.

A case that reaches `pptx/editor/render` lives in `cases/pptx-collaboration.ts`, which the relaxed pptx project typechecks; everything else here is strict.
