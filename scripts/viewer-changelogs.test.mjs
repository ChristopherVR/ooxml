import assert from 'node:assert/strict';
import { mkdirSync, mkdtempSync, readFileSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { test } from 'node:test';

import { checkChangelog } from './check-changelog-sections.mjs';
import { insertIntoChangelog, marker } from './major-notes.mjs';
import {
	HEADER,
	mergeGroups,
	missingDates,
	parsePackageChangelog,
	planDates,
	renderDaySection,
	updateViewerChangelogs,
	upsertDaySection,
} from './viewer-changelogs.mjs';

const pkg = (npm, sections) =>
	[
		'# Changelog',
		'',
		'Header text.',
		'',
		...sections.flatMap(([version, date, groups]) => [
			`## [${version}](https://example.test/${npm}@${version}) - ${date}`,
			'',
			...Object.entries(groups).flatMap(([heading, entries]) => [
				`### ${heading}`,
				'',
				...entries,
				'',
			]),
		]),
	].join('\n');

test('parses package sections with their groups and wrapped entries', () => {
	const sections = parsePackageChangelog(
		pkg('a', [['1.1.0', '2026-11-13', { Features: ['- **x:** One', '- Two\n  continued'] }]]),
	);
	assert.equal(sections.length, 1);
	assert.deepEqual(sections[0].groups, [
		{ heading: 'Features', entries: ['- **x:** One', '- Two\n  continued'] },
	]);
	assert.equal(sections[0].date, '2026-11-13');
});

test('merges groups in cliff order and lists each entry once', () => {
	const merged = mergeGroups([
		{
			groups: [
				{ heading: 'Documentation', entries: ['- d'] },
				{ heading: 'Features', entries: ['- f'] },
			],
		},
		{
			groups: [
				{ heading: 'Features', entries: ['- f', '- g'] },
				{ heading: 'Bug Fixes', entries: [] },
			],
		},
	]);
	assert.deepEqual(merged, [
		{ heading: 'Features', entries: ['- f', '- g'] },
		{ heading: 'Documentation', entries: ['- d'] },
	]);
});

test('a day with release notes lists them in their own order and the commits last', () => {
	const notes = insertIntoChangelog(
		pkg('a', [
			['2.0.0', '2026-11-30', { Features: ['- **x:** One'], Performance: ['- **x:** Two'] }],
		]),
		'2.0.0',
		['2.0'],
		[
			`${marker('2.0')}\n\nOverview.\n\n### Breaking changes\n\n- \`a\`: gone.\n\n### Performance\n\n- \`a\`: faster.`,
		],
	);
	const sections = parsePackageChangelog(notes);
	assert.deepEqual(
		mergeGroups([...sections, { groups: [{ heading: 'Fixed', entries: ['- `b`: fixed.'] }] }]),
		[
			{ heading: 'Breaking changes', entries: ['- `a`: gone.'] },
			{ heading: 'Fixed', entries: ['- `b`: fixed.'] },
			{ heading: 'Performance', entries: ['- `a`: faster.'] },
			{ heading: 'Commits', entries: ['- **x:** One', '- **x:** Two'] },
		],
	);
});

test('a day section names its releases and passes the stub check', () => {
	const section = renderDaySection('2026-11-15', [
		{ npm: 'a', version: '1.0.1', groups: [{ heading: 'Bug Fixes', entries: ['- fix'] }] },
		{ npm: 'b', version: '2.0.0', groups: [] },
	]);
	assert.match(
		section,
		/^## 2026-11-15\n\n_Releases: a@1\.0\.1, b@2\.0\.0_\n\n### Bug Fixes\n\n- fix\n/u,
	);
	assert.deepEqual(checkChangelog(section, { requireReleases: true }), []);
});

test('upserts by date: replaces the same day, inserts a new day in order, keeps old ones', () => {
	const old = 'Old header\n\n## 2026-10-05\n\n_Releases: a@0.1.0_\n\nold body\n';
	const added = upsertDaySection(old, '2026-11-13', '## 2026-11-13\n\n_Releases: a@1.0.0_\n');
	assert.ok(added.startsWith(HEADER));
	assert.ok(added.indexOf('## 2026-11-13') < added.indexOf('## 2026-10-05'));
	assert.ok(added.includes('old body'));
	const replaced = upsertDaySection(added, '2026-11-13', '## 2026-11-13\n\n_Releases: a@1.0.1_\n');
	assert.equal(replaced.match(/## 2026-11-13/gu).length, 1);
	assert.ok(replaced.includes('a@1.0.1') && !replaced.includes('a@1.0.0'));
	assert.equal(
		upsertDaySection(replaced, '2026-11-13', '## 2026-11-13\n\n_Releases: a@1.0.1_\n'),
		replaced,
	);
});

function fixtureRepo() {
	const root = mkdtempSync(join(tmpdir(), 'viewer-changelogs-'));
	const write = (path, text) => {
		mkdirSync(join(root, path, '..'), { recursive: true });
		writeFileSync(join(root, path), text);
	};
	write(
		'viewers/demo/packages/core/CHANGELOG.md',
		pkg('demo-core', [
			['1.1.0', '2026-11-13', { Features: ['- core feature'] }],
			['1.0.0', '2026-10-20', { 'Bug Fixes': ['- core fix'] }],
		]),
	);
	write(
		'viewers/demo/packages/react/CHANGELOG.md',
		pkg('demo-react', [['2.0.0', '2026-11-13', { Features: ['- react feature'] }]]),
	);
	write(
		'viewers/demo/CHANGELOG.md',
		'Old header\n\n## 2026-10-05\n\n_Releases: demo-core@0.9.0_\n',
	);
	const table = {
		'demo-core': { dir: 'viewers/demo/packages/core', npm: 'demo-core' },
		'demo-react': { dir: 'viewers/demo/packages/react', npm: 'demo-react' },
	};
	return { root, table, read: () => readFileSync(join(root, 'viewers/demo/CHANGELOG.md'), 'utf8') };
}

test('backfill adds every missing release day, newest first, once', () => {
	const { root, table, read } = fixtureRepo();
	assert.deepEqual(updateViewerChangelogs(root, missingDates, table), [
		'viewers/demo/CHANGELOG.md',
	]);
	const text = read();
	assert.ok(text.indexOf('## 2026-11-13') < text.indexOf('## 2026-10-20'));
	assert.ok(text.indexOf('## 2026-10-20') < text.indexOf('## 2026-10-05'));
	assert.match(text, /_Releases: demo-core@1\.1\.0, demo-react@2\.0\.0_/u);
	assert.deepEqual(updateViewerChangelogs(root, missingDates, table), []);
});

test('release mode writes only the days of the packages the plan released', () => {
	const { root, table, read } = fixtureRepo();
	const plan = {
		packages: {
			'demo-react': {
				npm: 'demo-react',
				dir: 'viewers/demo/packages/react',
				release: true,
				version: '2.0.0',
			},
			'demo-core': {
				npm: 'demo-core',
				dir: 'viewers/demo/packages/core',
				release: false,
				version: '1.1.0',
			},
		},
	};
	updateViewerChangelogs(root, planDates(plan), table);
	const text = read();
	assert.ok(text.includes('## 2026-11-13'));
	assert.ok(!text.includes('## 2026-10-20'));
	assert.match(text, /demo-core@1\.1\.0, demo-react@2\.0\.0/u);
	const none = { packages: { x: { npm: 'x', dir: 'src/core', release: true, version: '1.0.0' } } };
	assert.deepEqual(updateViewerChangelogs(root, planDates(none), table), []);
});

/** A viewer changelog whose top section is hand-written and waits for its date. */
const CRAFTED_VIEWER = [
	'# Changelog',
	'',
	'A hand-written preamble.',
	'',
	'## Unreleased',
	'',
	'<!-- release-notes: 2.0 -->',
	'',
	'_Releases: demo-react@2.0.0_',
	'',
	'The product summary.',
	'',
	'## 2026-10-05',
	'',
	'_Releases: demo-core@0.9.0_',
	'',
].join('\n');

const releasePlan = (versions) => ({
	packages: Object.fromEntries(
		Object.entries(versions).map(([npm, version]) => [
			npm,
			{ npm, dir: `viewers/demo/packages/${npm.slice(5)}`, release: true, version },
		]),
	),
});

test('release mode dates a hand-written section instead of rebuilding it', () => {
	const { root, table, read } = fixtureRepo();
	writeFileSync(join(root, 'viewers/demo/CHANGELOG.md'), CRAFTED_VIEWER);
	// The package section as crafted-changelogs.mjs --date left it, earlier in the run.
	writeFileSync(
		join(root, 'viewers/demo/packages/react/CHANGELOG.md'),
		'# Changelog\n\n## [2.0.0](x) - 2026-11-30\n\n<!-- release-notes: 2.0 -->\n\n### Added\n\n- a\n',
	);
	const plan = releasePlan({ 'demo-react': '2.0.0' });
	const options = { plan, date: '2026-11-30' };
	assert.deepEqual(updateViewerChangelogs(root, planDates(plan), table, options), [
		'viewers/demo/CHANGELOG.md',
	]);
	assert.equal(read(), CRAFTED_VIEWER.replace('## Unreleased', '## 2026-11-30'));
	assert.deepEqual(checkChangelog(read(), { requireReleases: true }), []);
	// A second run the same day changes nothing.
	assert.deepEqual(updateViewerChangelogs(root, planDates(plan), table, options), []);
});

test('a release the hand-written section does not name gets its own section that day', () => {
	const { root, table, read } = fixtureRepo();
	writeFileSync(join(root, 'viewers/demo/CHANGELOG.md'), CRAFTED_VIEWER);
	const plan = releasePlan({ 'demo-react': '2.0.0', 'demo-core': '1.1.0' });
	updateViewerChangelogs(root, planDates(plan), table, { plan, date: '2026-11-13' });
	const text = read();
	assert.equal(text.match(/^## 2026-11-13$/gmu).length, 2);
	assert.ok(text.includes('The product summary.'));
	assert.match(text, /^_Releases: demo-core@1\.1\.0_$/mu);
	assert.ok(text.indexOf('The product summary.') < text.indexOf('_Releases: demo-core@1.1.0_'));
	assert.deepEqual(checkChangelog(text, { requireReleases: true }), []);
});

test('a hand-written section stays pending while the plan releases nothing it names', () => {
	const { root, table, read } = fixtureRepo();
	writeFileSync(join(root, 'viewers/demo/CHANGELOG.md'), CRAFTED_VIEWER);
	const plan = releasePlan({ 'demo-core': '1.1.0' });
	updateViewerChangelogs(root, planDates(plan), table, { plan, date: '2026-11-13' });
	const text = read();
	assert.ok(text.startsWith(`${HEADER}\n## Unreleased\n`), text);
	assert.match(text, /^## 2026-11-13\n\n_Releases: demo-core@1\.1\.0_$/mu);
	assert.deepEqual(checkChangelog(text, { requireReleases: true }), []);
});

test('backfill keeps a pending hand-written section at the top', () => {
	const { root, table, read } = fixtureRepo();
	writeFileSync(join(root, 'viewers/demo/CHANGELOG.md'), CRAFTED_VIEWER);
	updateViewerChangelogs(root, missingDates, table);
	const text = read();
	assert.ok(text.startsWith(`${HEADER}\n## Unreleased\n`), text);
	assert.ok(text.indexOf('## 2026-11-13') < text.indexOf('## 2026-10-20'));
	assert.deepEqual(checkChangelog(text, { requireReleases: true }), []);
});
