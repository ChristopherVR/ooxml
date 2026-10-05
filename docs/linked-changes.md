# Changes that span `ooxml` and a viewer

A feature or fix often needs the model or logic here first and the viewer (or its binding) second.
The viewers depend on the **published** `ooxml-core` and `ooxml-ui`, so the order matters.

## The order

1. **Land the `ooxml` change first**, with its own tests, as a conventional commit whose type and
   paths release (`feat` or `fix` under `src/` or `src/ui/src`). A pull request here is fine
   for outside contributors; maintainers commit straight to `main`.
2. **Release it.** Run `gh workflow run release.yml -R ChristopherVR/ooxml` (or wait for the
   schedule). Do not start the viewer change by assuming the release exists.
3. **Land the viewer change on the new version.** The viewers adopt new releases through the
   scheduled `sync-ooxml` workflow, or by hand with `node scripts/sync-ooxml-deps.mjs` from the
   viewer.

Until step 2 is done, develop the viewer change against a local checkout:
`node scripts/link-local.mjs <viewer dir>` here, then `--restore` before committing (see
`AGENTS.md`). A viewer change may use something unreleased only if it also works on the last
release.

## Linking the two

- Open the viewer pull request as a **draft** while it depends on an unreleased `ooxml` change.
- Put `Depends on: ChristopherVR/ooxml#<number>` (or the commit) at the top of the viewer pull
  request description, and `Needed by: ChristopherVR/<viewer>#<number>` in the `ooxml` one, so the
  pair can be found from either side.
- Mark the viewer pull request ready only once the `ooxml` change is released and the viewer's
  range points at that release (no `file:` ranges; `check:published` rejects them).
- A single pull request must not change both repositories' behaviour: split it, even if the two
  halves are written together.

## Merging

- Merge the `ooxml` side first and confirm the release (`ooxml-core@<version>` tag, then npm).
- If the viewer's CI is red only because the new version is not adopted yet, fix it by adopting
  the release, not by loosening checks.
- When an automated adoption fails (`Adopting ... fails verification` issue), the issue lists the
  failing step and the last lines of its output. A failure caused by core belongs in an `ooxml`
  fix; a failure caused by the viewer belongs in the viewer. The issue closes itself when a later
  bump verifies.
