# Performance budgets

CI fails a change that makes the packages or the suite download more, or makes an engine do
more work, than its budget allows. This keeps the speed won in 2.0 from slipping back one
small change at a time. Every budget here is a byte count or a work count, so it measures
the same on a laptop and on a shared CI runner. Wall-clock timings are not budgets (see
[Timings](#timings)).

## Byte budgets

`node scripts/perf-budgets.mjs check` measures these and compares them with the limits in
`scripts/perf-budgets.json`:

| Budget                 | What it measures                                                                                                                                                                                                                                                                                 |
| ---------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| `package/<entry>`      | Each published ESM entry point of `ooxml-core` and `ooxml-ui`: the entry file and every file it imports statically, gzipped, from `dist`. This is the measurement the `bundle-size` comment shows on a pull request.                                                                             |
| `suite/all javascript` | Every JavaScript file of the built suite site except the docs pages, gzipped. This is the number the 2.0 target of at most 9 MB applies to.                                                                                                                                                      |
| `suite/route/<screen>` | What one screen of the suite downloads: the gzipped JavaScript files Chromium requested from opening the page until no script was requested for 1.5 seconds. The screens are the start chooser, home, the four product landing pages, chat, New in Word, Excel, PowerPoint and Visio, and Teams. |

A screen is measured from the requests the browser makes, not from the chunk graph, because
the regressions that matter most are lazy imports: New loading the automation bundle, or a
ribbon loading the SmartArt definitions while it is built. A static graph cannot tell which
lazy import a screen takes. Every page except the start chooser skips the chooser, the way a
returning user does, and a page that throws fails the check.

The script builds the suite into a temporary directory itself, so it needs the core and
`ooxml-ui` builds (`bun run build`, then `bun run --cwd src/ui build`) and Playwright's Chromium
(`bun x playwright install chromium` in `e2e/a11y`). It also refuses build output that contains
the absolute path of the checkout, because such output differs from machine to machine and
would publish that path.

### Limits

A limit is the measurement plus 3%, at least 1 KiB more, rounded up to a whole KiB. The check
fails when:

- a measurement is over its limit (a regression);
- a limit is more than 10% above what `update` would write for the current measurement, so a
  regression of that size would pass unseen (lower it to keep the win);
- a measurement has no limit (a new entry point or screen), or a limit has no measurement (one
  was removed).

### When CI runs them

The `perf-budgets` job runs whenever the plan's `budgets` flag is set (`scripts/ci-plan.mjs`):
a change to shipped code in `src/core` or `src/ui`, to `apps/office-suite`, to the PowerPoint
viewer packages the suite compiles in, to the suite shell in `site/` (not its docs), or to the
budget and suite build scripts. Tests and docs do not run it. The nightly full run always does.
`bun run verify --browser` runs it locally.

### Changing a budget on purpose

1. Build the core and `ooxml-ui`, then run `node scripts/perf-budgets.mjs measure --files` to
   see each number and the files every screen requested. Compare the lists with a run on the
   base commit to find what a screen started loading.
2. If the growth is a regression, fix it in the code.
3. If it is deliberate (a feature that has to load on that screen), or the code got smaller,
   run `node scripts/perf-budgets.mjs update` and commit `scripts/perf-budgets.json` with the
   change. Say in the commit message which budgets moved, by how much, and why.

`update` rewrites every limit from one measurement, so review the whole JSON diff: only the
budgets the change explains should move. A change that lowers a number by more than 10% has to
lower its limit in the same commit, which locks the win in.

## Work budgets

Engine wins whose regression is a count of work, not a size, are guarded by unit tests that
assert the count. They run with the unit suite of their area whenever its code changes.

| Win                                                                     | Test                                                          | Counts                                    |
| ----------------------------------------------------------------------- | ------------------------------------------------------------- | ----------------------------------------- |
| Word load copies its model once; saves copy it never                    | `src/core/docx/load-work.test.ts`                             | `structuredClone` calls on a whole model  |
| PowerPoint reads package relationships once for all charts              | `src/core/pptx/__tests__/integration/chart-load-work.test.ts` | `structuredClone` calls across a load     |
| Word save moves only changed blocks                                     | `src/core/docx/write-slots.test.ts`                           | DOM inserts and removals                  |
| Excel lookups over one range share an index                             | `src/core/xlsx/formula/lookup-index.test.ts`                  | cells read against a plain scan           |
| Excel undo records what an edit changed                                 | `src/core/xlsx/edit/history-perf.test.ts`                     | cells held by an undo step                |
| Excel row refit recalculates only dependents                            | `src/core/xlsx/edit/session-rows.test.ts`                     | full recalculations and rebuilds          |
| PowerPoint rail reads scrolling once per frame and fills rows in budget | `src/ui/src/pptx/render/virtualized-list.test.ts`             | reads, window changes and fills per frame |
| The tab strip scrolls only when its selection or tabs change            | `src/ui/src/navigation.test.ts`                               | `scrollIntoView` calls                    |

## Timings

There are no wall-clock budgets. A timing budget needs its run-to-run spread on a shared
hosted runner to be small next to the regression it should catch, and that spread has not
been measured: runners share cores with other jobs, and the unit suite runs files in parallel.
The engine tests that assert a time (`src/core/xlsx/formula/engine-perf.test.ts`,
`src/core/xlsx/cells.test.ts`, `src/core/xlsx/ui/commands/util.test.ts`,
`src/core/pptx/core/core/runtime/slide-fingerprint.test.ts`,
`src/core/docx/style-inheritance.test.ts`) guard against orders of magnitude, not percentages:
their limits on CI are about three times the local time, or a ratio between two input sizes.
A timing budget belongs in this page only with the spread measured over CI runs and the
headroom stated next to it.
