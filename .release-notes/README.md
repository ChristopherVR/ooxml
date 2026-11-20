# Release notes for a major release

A folder here holds the hand-written notes of one coordinated major release, such as `2.0/`. The release name is not a package version: every package moves to its own next major. When the release workflow releases a package with a major bump, `scripts/major-notes.mjs` writes that package's share of the notes into its `CHANGELOG.md`, under the new version's heading and above the git-cliff commit list. The GitHub release body gets the same block. See [docs/releasing.md](../docs/releasing.md#hand-written-notes-for-a-major-release).

## The files

- `overview.md` is a few paragraphs of plain prose that open the notes of every package.
- Every other `.md` file is a fragment from one piece of work. Its name is never shown to readers.

## Write a fragment

A fragment has `## <section>` headings and `- ` bullets, and nothing else. The sections are `Breaking`, `Migration`, `Security`, `Added`, `Changed`, `Fixed`, `Performance` and `Accessibility`, and a package's notes list them in that order.

Start every bullet with the packages it affects, then a colon:

```md
## Breaking

- `ooxml-core/teams`, `openteams-server` and OpenTeams viewers: a new 1:1 chat id is `dm-` and 32 hex digits.
```

A name in the lead-in is one of these:

- an npm name from `scripts/release-plan.mjs`, in backticks, such as `` `pptx-viewer-core` ``;
- a subpath of one, in backticks, such as `` `ooxml-core/opc` ``, which counts for its package;
- a group, without backticks: `All packages`, `PowerPoint viewers`, `Word viewers`, `Excel viewers`, `Visio viewers`, `OpenTeams viewers` or `MCP packages`;
- `Office suite`, for the hosted app and the site, which are not on npm. These bullets reach no package changelog.

Separate names with commas and `and`. The group of a viewer product holds its framework packages. `PowerPoint viewers` also holds `@christophervr/pptx-viewer`, which re-exports the React viewer.

Write for package users, in plain sentences. Name no internal work ids, write no em dashes, and describe a security defect by its class, never by how to exploit it.

## Check the notes

```sh
node scripts/major-notes.mjs --check
node scripts/major-notes.mjs --render pptx-react 2.0   # one package's notes, by planner key
```

`--check` fails on a bullet that names no package, an unknown name, a line that is neither a heading nor a bullet, and a published package that no bullet names. `bun run test:scripts` runs the same check, and so does the release workflow before it tags anything.
