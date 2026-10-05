<!--
Thanks for contributing. Please read CONTRIBUTING.md if you have not already:
https://github.com/ChristopherVR/ooxml/blob/main/CONTRIBUTING.md

Delete any section that genuinely does not apply.
-->

## What does this change?

<!-- One or two sentences. Link the issue if there is one: Fixes #123 -->

## Needed by

<!-- If a viewer repository is waiting for this, link its PR or issue:
Needed by: ChristopherVR/pptx-viewer#123. Land this first, release it, then the viewer; see
docs/linked-changes.md. Delete if none. -->

## Type of change

- [ ] Bug fix (non-breaking)
- [ ] New feature (non-breaking)
- [ ] Breaking change (viewers adopt releases automatically; say what breaks)
- [ ] Docs / tooling / CI only

## Area

<!-- e.g. xml, opc, docx, pptx, xlsx, visio, chart, ooxml-ui/xlsx. -->

- [ ] The change is made once in the area that owns it, not copied per format
- [ ] DOM-free logic stays in `src/<area>/`; custom elements stay in `packages/ui`
- [ ] Unsupported behaviour is reported honestly (no claim of Office parity without evidence)
- [ ] Extraction provenance is recorded in `PROVENANCE.md` for any module moved here

## Testing

- [ ] Parsing, preservation, round-trip or editing tests added or updated next to the code
- [ ] Regression test added that fails without this change
- [ ] `bun run test:package` passes if I added or changed an entry point

## Checks run locally

- [ ] `bun run lint`
- [ ] `bun run fmt:check`
- [ ] `bun run typecheck`
- [ ] `bun run test`

## Conventional Commits

- [ ] My commit messages follow
      [Conventional Commits](https://www.conventionalcommits.org)

Commit types drive the **published version bump**, and the paths a commit touches decide whether
it releases at all (only published files under `src/`, the bundler configs, the manifest's
shipping fields and the licence files do). `feat` bumps minor, `fix`/`perf`/`refactor`/`chore`
bump patch, `!` or a `BREAKING CHANGE:` footer bumps major.

## Screenshots / recordings

<!-- For UI changes in packages/ui, a screenshot helps. -->
