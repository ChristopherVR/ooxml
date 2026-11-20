import assert from 'node:assert/strict';
import { mkdirSync, mkdtempSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join } from 'node:path';
import { test } from 'node:test';
import { fileURLToPath } from 'node:url';

import {
	insertIntoChangelog,
	loadNotes,
	marker,
	noteNames,
	parseFragment,
	pendingNotes,
	releaseBody,
	renderPackageNotes,
} from './major-notes.mjs';

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..');

/** A planner table with a core, a UI package, two PowerPoint bindings and a server. */
const TABLE = {
	core: { dir: 'src/core', npm: 'ooxml-core' },
	ui: { dir: 'src/ui', npm: 'ooxml-ui' },
	'pptx-react': { dir: 'viewers/pptx/packages/react', npm: 'pptx-react-viewer' },
	'pptx-vue': { dir: 'viewers/pptx/packages/vue', npm: 'pptx-vue-viewer' },
	'teams-server': { dir: 'viewers/teams/server', npm: 'openteams-server' },
};

/** A repository under a temporary directory holding the manifests of TABLE and `files`. */
function repo(files) {
	const root = mkdtempSync(join(tmpdir(), 'major-notes-'));
	for (const meta of Object.values(TABLE)) {
		mkdirSync(join(root, meta.dir), { recursive: true });
		writeFileSync(join(root, meta.dir, 'package.json'), '{"version":"1.0.0"}');
	}
	for (const [path, text] of Object.entries(files)) {
		mkdirSync(dirname(join(root, path)), { recursive: true });
		writeFileSync(join(root, path), text);
	}
	return root;
}

const FRAGMENT = [
	'## Fixed',
	'',
	'- PowerPoint viewers: slides paint.',
	'## Breaking',
	'',
	'- `ooxml-core/opc` and `openteams-server`: limits apply.',
	'  The second line of that entry.',
	'- Office suite: keys stay in the tab.',
	'## Migration',
	'',
	'- All packages: upgrade together.',
	'',
].join('\n');

test('a bullet takes the packages its lead-in names, by npm name, subpath or group', () => {
	const { entries, problems } = parseFragment(FRAGMENT, 'a.md', TABLE);
	assert.deepEqual(problems, []);
	assert.deepEqual(
		entries.map((e) => [e.kind, e.packages]),
		[
			['Fixed', ['pptx-react-viewer', 'pptx-vue-viewer']],
			['Breaking', ['ooxml-core', 'openteams-server']],
			['Breaking', []],
			['Migration', Object.values(TABLE).map((m) => m.npm)],
		],
	);
	assert.equal(
		entries[1].text,
		'`ooxml-core/opc` and `openteams-server`: limits apply.\n  The second line of that entry.',
	);
});

test('a bullet with no lead-in, an unknown name or a stray line is a problem with its place', () => {
	const text = [
		'# Title',
		'- `ooxml-core`: before any section.',
		'## Fixed',
		'- Saving works again.',
		'- PowerPoint charts: blanks stay blank.',
		'- `ooxml-cor`, `ooxml-ui`: a typo.',
		'## Notes',
		'',
	].join('\n');
	const { entries, problems } = parseFragment(text, 'b.md', TABLE);
	assert.equal(entries.length, 0);
	assert.equal(problems.length, 6);
	assert.match(problems[0], /^b\.md:1: expected a "## <section>" heading/u);
	assert.match(problems[1], /^b\.md:2: bullet before any/u);
	assert.match(problems[2], /^b\.md:4: names no package/u);
	assert.match(problems[3], /^b\.md:5: names no package/u);
	assert.match(problems[4], /^b\.md:6: unknown package `ooxml-cor`/u);
	assert.match(problems[5], /^b\.md:7: unknown section "Notes"/u);
});

test('every published package needs an entry, and Office suite entries reach none', () => {
	const root = repo({
		'.release-notes/2.0/overview.md': 'One release.\n',
		'.release-notes/2.0/A.md':
			'## Fixed\n\n- `ooxml-core`, PowerPoint viewers: one.\n- Office suite: two.\n',
	});
	assert.deepEqual(loadNotes(root, '2.0', TABLE).problems, [
		'.release-notes/2.0: no entry names `ooxml-ui`; every published package needs one',
		'.release-notes/2.0: no entry names `openteams-server`; every published package needs one',
	]);
	const empty = repo({ '.release-notes/2.0/A.md': '## Fixed\n\n- All packages: one.\n' });
	assert.deepEqual(loadNotes(empty, '2.0', TABLE).problems, [
		'.release-notes/2.0/overview.md: missing or empty',
	]);
});

test("a package's notes hold the overview and its own entries, breaking changes first", () => {
	const notes = {
		name: '2.0',
		overview: 'One release.',
		entries: parseFragment(FRAGMENT, 'a.md', TABLE).entries,
	};
	assert.equal(
		renderPackageNotes(notes, 'pptx-vue-viewer'),
		[
			marker('2.0'),
			'',
			'One release.',
			'',
			'### Migration',
			'',
			'- All packages: upgrade together.',
			'',
			'### Fixed',
			'',
			'- PowerPoint viewers: slides paint.',
		].join('\n'),
	);
	assert.match(
		renderPackageNotes(notes, 'ooxml-core'),
		/### Breaking changes\n\n- `ooxml-core\/opc`[^\n]*\n {2}The second line[^#]*### Migration/u,
	);
});

const CHANGELOG = [
	'# Changelog',
	'',
	'## [5.0.0](https://example.test/5.0.0) - 2026-11-30',
	'',
	'### Features',
	'',
	'- **pptx:** One ([abc](https://example.test/abc))',
	'',
	'## [4.81.0](https://example.test/4.81.0) - 2026-11-18',
	'',
	'### Bug Fixes',
	'',
	'- **pptx:** Two',
	'',
].join('\n');

test('the notes go under the new version heading, above its commits, once', () => {
	const block = `${marker('2.0')}\n\nOne release.`;
	const once = insertIntoChangelog(CHANGELOG, '5.0.0', ['2.0'], [block]);
	assert.equal(
		once,
		[
			'# Changelog',
			'',
			'## [5.0.0](https://example.test/5.0.0) - 2026-11-30',
			'',
			marker('2.0'),
			'',
			'One release.',
			'',
			'### Commits',
			'',
			'#### Features',
			'',
			'- **pptx:** One ([abc](https://example.test/abc))',
			'',
			'## [4.81.0](https://example.test/4.81.0) - 2026-11-18',
			'',
			'### Bug Fixes',
			'',
			'- **pptx:** Two',
			'',
		].join('\n'),
	);
	assert.equal(insertIntoChangelog(once, '5.0.0', ['2.0'], [block]), once);
	assert.throws(
		() => insertIntoChangelog(CHANGELOG, '6.0.0', ['2.0'], [block]),
		/no "## \[6\.0\.0\]" section/u,
	);
});

test('a release body keeps its heading and puts the notes above the commits', () => {
	const body = '## [5.0.0](https://example.test/5.0.0) - 2026-11-30\n\n### Features\n\n- One';
	assert.equal(
		releaseBody(body, ['Notes.']),
		'## [5.0.0](https://example.test/5.0.0) - 2026-11-30\n\nNotes.\n\n### Commits\n\n#### Features\n\n- One\n',
	);
	assert.equal(
		releaseBody('### Changes\n\n- One', ['Notes.']),
		'Notes.\n\n### Commits\n\n#### Changes\n\n- One\n',
	);
	assert.equal(releaseBody(body, []), body);
});

test('only a major release carries the notes, and only until a changelog section has them', () => {
	const root = repo({ '.release-notes/2.0/overview.md': 'One release.\n' });
	const pkg = { release: true, bump: 'major', version: '5.0.0' };
	assert.deepEqual(pendingNotes(root, pkg, CHANGELOG), ['2.0']);
	assert.deepEqual(pendingNotes(root, { ...pkg, bump: 'minor', version: '5.1.0' }, CHANGELOG), []);
	assert.deepEqual(pendingNotes(root, { ...pkg, release: false }, CHANGELOG), []);
	const written = insertIntoChangelog(CHANGELOG, '5.0.0', ['2.0'], [marker('2.0')]);
	// The release-notes step runs after the changelog step: the notes are still this release's.
	assert.deepEqual(pendingNotes(root, pkg, written), ['2.0']);
	// A later major finds them in an older section and leaves them out.
	assert.deepEqual(pendingNotes(root, { ...pkg, version: '6.0.0' }, written), []);
});

test('the release notes in this repository name only real packages and cover every one', () => {
	const problems = noteNames(ROOT).flatMap((name) => loadNotes(ROOT, name).problems);
	assert.deepEqual(problems, [], problems.join('\n'));
});
