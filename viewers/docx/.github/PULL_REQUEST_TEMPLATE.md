<!--
Thanks for contributing. Please read CONTRIBUTING.md if you have not already:
https://github.com/ChristopherVR/docx-viewer/blob/main/CONTRIBUTING.md

Delete any section that genuinely does not apply.
-->

## What does this change?

<!-- One or two sentences. Link the issue if there is one: Fixes #123 -->

## Depends on

<!-- If this needs an unreleased ooxml change, open this PR as a draft and link it:
Depends on: ChristopherVR/ooxml#123. Land and release the ooxml side first; see
https://github.com/ChristopherVR/ooxml/blob/main/docs/linked-changes.md. Delete if none. -->

## Type of change

- [ ] Bug fix (non-breaking)
- [ ] New feature (non-breaking)
- [ ] Breaking change
- [ ] Docs / tooling / CI only

---

## Where does the change belong?

This repository holds only the framework bindings, demos, browser tests and docs. Logic lives in
`ooxml-core/docx` and the `<docx-editor>` element in `ooxml-ui/docx` (both in the
[`ooxml`](https://github.com/ChristopherVR/ooxml) repository). If your change is really about
parsing, the document model, editing behaviour, or the element itself, it belongs there.

- [ ] This change only touches bindings, demos, tests or docs in this repository
- [ ] The logic or element change it needs is already released in `ooxml` (version: )

## Cross-binding parity

**Skip this section only if your change touches no binding code at all** (docs, CI, tooling).

All bindings wrap the same element, so a user on Svelte is entitled to the feature set a user on
React gets. Parity is a merge requirement.

### Which bindings did you check, and what did you find?

<!--
Be honest here. "Checked" means you actually ran the demo or a test against it,
not that you assumed. Reviewers care more about an accurate map than a clean one.
-->

| Binding                      | Affected?    | Fixed / implemented here? | Notes |
| ---------------------------- | ------------ | ------------------------- | ----- |
| React (`packages/react`)     | yes / no / ? |                           |       |
| Vue (`packages/vue`)         | yes / no / ? |                           |       |
| Angular (`packages/angular`) | yes / no / ? |                           |       |
| Svelte (`packages/svelte`)   | yes / no / ? |                           |       |
| Solid (`packages/solid`)     | yes / no / ? |                           |       |
| Vanilla (`packages/vanilla`) | yes / no / ? |                           |       |

If any affected binding is **not** fixed here, say which and why, and link the
tracking issue:

<!-- e.g. "Angular needs a change-detection refactor first, tracked in #123" -->

---

## Testing

- [ ] Unit tests added or updated for each binding I changed
- [ ] Regression test added that fails without this change
- [ ] Browser (Playwright) test added or updated when behaviour in the browser changed
- [ ] `bun run test:browser` passes locally, or I have named what I ran

## Checks run locally

- [ ] `bun run lint`
- [ ] `bun run fmt:check`
- [ ] `bun run typecheck`
- [ ] `bun run test`

## Conventional Commits

- [ ] My commit messages follow
      [Conventional Commits](https://www.conventionalcommits.org)

Commit types drive **published package version bumps**, and non-conforming
commits are dropped from the changelog entirely, so this is load-bearing.
`feat` bumps minor, `fix`/`perf`/`refactor`/`chore` bump patch, `!` or a
`BREAKING CHANGE:` footer bumps major. Which package gets bumped is decided by
the **paths** the commit touches, so keep commits within one package where you
can.

## Screenshots / recordings

<!--
For UI changes, a screenshot per affected binding is worth a lot. Even two
(the React reference plus the binding you were least sure about) helps.
-->
