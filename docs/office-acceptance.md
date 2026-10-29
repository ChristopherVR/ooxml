# Native Office acceptance

Our writers claim that Word, Excel, PowerPoint and Visio open what we save without repair. Unit tests cannot show that, so a set of PowerShell scripts drives the real desktop apps through COM. Hosted CI runners have no Office, so these run on a self-hosted Windows runner by `.github/workflows/office-acceptance.yml`, weekly and on demand, and never from the main CI.

A Stable label for a feature (see [stability](stability.md)) requires a recent passing run of the product's acceptance scripts on this workflow.

## What runs

`scripts/office-acceptance.mjs <product>` runs the scripts of one product, records each script's result and writes `summary.json`. A script passes when `pwsh -File` exits 0 (they throw or `exit 1` on failure).

| Product      | Scripts                                                                                                                                                                                                           | Also needs                                             |
| ------------ | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------ |
| `word`       | `src/core/scripts/docx-word-acceptance.ps1`, `docx-word-chart-acceptance.ps1`, `shared-picture-native-acceptance.ps1`                                                                                             | Excel (charts, picture workbook)                       |
| `excel`      | `src/core/scripts/xlsx-excel-acceptance.ps1`; `scripts/xlsx-pivot-acceptance.ps1`, `xlsx-slicer-acceptance.ps1`, `xlsx-sparkline-acceptance.ps1`, `xlsx-print-acceptance.ps1`, `xlsx-print-format-acceptance.ps1` | print checks: `pdfinfo`, Python with pypdf and PyMuPDF |
| `powerpoint` | `src/core/scripts/pptx-classic-chart-acceptance.ps1`                                                                                                                                                              | Excel (Edit Data workbook)                             |
| `visio`      | `scripts/record-visio-new-drawing.ps1`, `record-visio-size-position.ps1` (fixtures from the matching `create-visio-*-fixtures.ts`)                                                                                | Visio                                                  |

Not wired:

- The other `scripts/record-visio-*.ps1` and `record-xlsx-*.ps1` scripts record what the real app does, to capture reference values by hand; they have no pass or fail result. The `xlsx-pivot-*-oracle.ps1` scripts likewise only emit oracle data.
- `e2e/pptx/native/*.ps1` (chart title and label checks, animation renders, clipboard) need an artifact folder produced by the PowerPoint browser suite or a manual step, with no generator that fits this runner yet.
- `scripts/check-word-*-exports.ps1` compare exports against a reference folder that must be supplied.

A step is skipped, with a notice, when a program it needs (`pdfinfo`, `python`, another Office app) is missing. A whole product is skipped when its COM class is absent (typically Visio). Pass `--require` to make a missing app a failure.

## Registering the runner

1. A Windows 10 or 11 machine (a dedicated VM is best) with the desktop Office apps installed: Word, Excel and PowerPoint (Microsoft 365 Apps or Office 2019 and later) and optionally Visio. Open each app once, accept the licence prompts and sign in so no first-run dialog blocks automation. Office needs a licence for the machine or user; use one your organisation holds, and do not share the account outside the runner.
2. Install Git for Windows (the setup action uses its `bash`), PowerShell 7 (`pwsh`), Node 24 and Bun. For the print checks add Python with `pypdf` and `pymupdf`, and Poppler's `pdfinfo` on `PATH`.
3. Register a runner for the repository (Settings, Actions, Runners, New self-hosted runner) with the labels `self-hosted`, `windows` and `office`.
4. Run it in an **interactive session**, not as a service in session 0. Office COM automation fails or hangs without a desktop: sign in as a dedicated user, enable automatic logon and start `run.cmd` at logon (a scheduled task "at log on" works). Disable the screen lock and sleep.
5. Keep the machine free of other Office windows while a run is going; the scripts refuse to start when a Word, Excel or PowerPoint process already exists. The workflow runs one job at a time (`max-parallel: 1` and a `concurrency` group).
6. Only this repository's workflows should use the runner. Do not enable it for pull requests from forks; the workflow runs only on schedule and manual dispatch.

## Running locally

```sh
node scripts/office-acceptance.mjs excel                 # one product
node scripts/office-acceptance.mjs all --out ./results   # every product the machine can run
```

Each script can also be run alone, for example `pwsh -File src/core/scripts/xlsx-excel-acceptance.ps1`; its header lists its options. The scripts run TypeScript through `bun`, so no build is needed. Close open Word, Excel and PowerPoint windows first.

## Results

The workflow uploads one artifact per product, `office-acceptance-<product>`, kept for 90 days, and writes a table to the job summary. It contains:

- `summary.json`: product, date, commit, runner, the installed Office apps with `version` and `build`, an overall `status` (`passed`, `failed` or `skipped`), `totals` (`passed`, `failed`, `skipped`) and one entry per script (`id`, `script`, `status`, `exitCode`, `durationMs`, `passed`, `failed`). Counts come from the script's "All N documents" or "FAILED:" line; a script without one counts as one pass or one failure.
- `run.log`: the combined output of the product's scripts.
- one folder per script with the files it generated (documents, workbooks, manifests) for inspection. The pivot, slicer, sparkline and print scripts write to their own temporary folder instead.

A failed run is a failed job. The scheduled run only reads the repository; it never publishes, tags or pushes.

## Recorded runs

Local runs of `node scripts/office-acceptance.mjs <product>` on a Windows workstation with Office 16.0 build 20527 (Word, Excel, PowerPoint and Visio), before the self-hosted runner is registered. The scheduled workflow replaces these once it runs.

| Date       | Product    | Commit      | Status | Passed | Failed | Skipped                                   |
| ---------- | ---------- | ----------- | ------ | ------ | ------ | ----------------------------------------- |
| 2026-10-29 | Word       | `ab769a453` | passed | 22     | 0      | 0                                         |
| 2026-10-29 | Excel      | `09976a614` | passed | 124    | 0      | 1 (`xlsx-print`: `pdfinfo` not installed) |
| 2026-10-29 | PowerPoint | `09976a614` | passed | 1      | 0      | 0                                         |
| 2026-10-29 | Visio      | `ab769a453` | passed | 2      | 0      | 0                                         |
