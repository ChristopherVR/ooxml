import assert from 'node:assert/strict';
import { mkdirSync, mkdtempSync, readFileSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join } from 'node:path';
import { test } from 'node:test';

import {
	craftedNotes,
	craftedSectionFor,
	datedCraftedSection,
	datePackages,
	datePending,
	notesBody,
	packageProblems,
	pendingKeys,
	pendingSection,
	planProblems,
	releaseUrl,
	viewerProblems,
} from './crafted-changelogs.mjs';
import { pendingNotes } from './major-notes.mjs';

const PREAMBLE = ['# Changelog', '', 'All notable changes to this package, newest first.', ''];
const OLD = [
	'## [1.4.0](https://github.com/ChristopherVR/ooxml/releases/tag/demo-core@1.4.0) - 2026-11-01',
	'',
	'### Features',
	'',
	'- **demo:** an old feature',
	'',
];

/** A package changelog with a hand-written section for `version` of `npm` at the top. */
const crafted = (npm, version, link = releaseUrl(npm, version)) =>
	[
		...PREAMBLE,
		`## [${version}](${link}) - Unreleased`,
		'',
		'<!-- release-notes: 2.0 -->',
		'',
		'The lead paragraph.',
		'',
		'### Breaking changes',
		'',
		'- `demo` now does it differently.',
		'',
		...OLD,
	].join('\n');

const pkg = (npm, version, dir, extra = {}) => ({
	npm,
	version,
	dir,
	release: true,
	bump: 'major',
	tag: `${npm}@${version}`,
	changelog: `${dir}/CHANGELOG.md`,
	...extra,
});

test('reads the pending section at the top and strips its marker for publishing', () => {
	const pending = pendingSection(crafted('demo-core', '2.0.0'));
	assert.equal(pending.version, '2.0.0');
	assert.equal(pending.url, releaseUrl('demo-core', '2.0.0'));
	assert.equal(
		notesBody(pending.body),
		'The lead paragraph.\n\n### Breaking changes\n\n- `demo` now does it differently.',
	);
	assert.equal(pendingSection([...PREAMBLE, ...OLD].join('\n')), null);
});

test('only the top section counts as pending', () => {
	const text = [...PREAMBLE, ...OLD, '## [2.0.0](x) - Unreleased', '', 'body'].join('\n');
	assert.equal(pendingSection(text), null);
	assert.equal(datePending(text, '2026-11-30'), text);
});

test('a pending section must match the planned version and tag', () => {
	const core = pkg('demo-core', '2.0.0', 'viewers/demo/packages/core');
	assert.deepEqual(packageProblems(core, crafted('demo-core', '2.0.0')), []);
	const [version] = packageProblems({ ...core, version: '2.0.1' }, crafted('demo-core', '2.0.0'));
	assert.match(
		version,
		/hand-written section is for 2\.0\.0, but this release plans demo-core@2\.0\.1/u,
	);
	const [link] = packageProblems(core, crafted('demo-core', '2.0.0', releaseUrl('other', '2.0.0')));
	assert.match(link, /links .*other@2\.0\.0, expected .*demo-core@2\.0\.0/u);
	const empty = [
		...PREAMBLE,
		`## [2.0.0](${releaseUrl('demo-core', '2.0.0')}) - Unreleased`,
		'',
		'<!-- release-notes: 2.0 -->',
		'',
	].join('\n');
	assert.match(packageProblems(core, empty)[0], /is empty/u);
	assert.deepEqual(packageProblems(core, [...PREAMBLE, ...OLD].join('\n')), []);
});

test('dates the package and the viewer pending forms, and nothing else', () => {
	const dated = datePending(crafted('demo-core', '2.0.0'), '2026-11-30');
	assert.ok(
		dated.includes(`## [2.0.0](${releaseUrl('demo-core', '2.0.0')}) - 2026-11-30\n`),
		dated,
	);
	assert.ok(!dated.includes('Unreleased'));
	assert.ok(dated.includes('## [1.4.0]') && dated.includes('<!-- release-notes: 2.0 -->'));
	const viewer = '# Changelog\n\n## Unreleased\n\n_Releases: a@2.0.0_\n\n## 2026-11-01\n';
	assert.equal(
		datePending(viewer, '2026-11-30'),
		'# Changelog\n\n## 2026-11-30\n\n_Releases: a@2.0.0_\n\n## 2026-11-01\n',
	);
	const plain = [...PREAMBLE, ...OLD].join('\n');
	assert.equal(datePending(plain, '2026-11-30'), plain);
});

test('a viewer section must name the releases of the plan, once the plan releases one of them', () => {
	const text = (list) => `# Changelog\n\n## Unreleased\n\n<!-- release-notes: 2.0 -->\n\n${list}\n`;
	const released = new Map([
		['a', '2.0.0'],
		['b', '2.0.0'],
	]);
	const path = 'viewers/demo/CHANGELOG.md';
	assert.deepEqual(viewerProblems(released, text('_Releases: a@2.0.0, b@2.0.0_'), path), []);
	assert.match(
		viewerProblems(released, text('_Releases: a@2.0.0, b@2.1.0_'), path)[0],
		/names b@2\.1\.0, but this release plans b@2\.0\.0/u,
	);
	assert.match(
		viewerProblems(released, text('_Releases: a@2.0.0, c@2.0.0_'), path)[0],
		/names c@2\.0\.0, which this release does not release/u,
	);
	assert.match(viewerProblems(released, text(''), path)[0], /has no `_Releases/u);
	// Written ahead of a release this run does not make: left pending.
	assert.deepEqual(viewerProblems(new Map(), text('_Releases: a@2.0.0_'), path), []);
});

test('the release body is the section body, found before or after dating', () => {
	const core = pkg('demo-core', '2.0.0', 'viewers/demo/packages/core');
	const pending = crafted('demo-core', '2.0.0');
	const expected =
		'The lead paragraph.\n\n### Breaking changes\n\n- `demo` now does it differently.';
	assert.equal(craftedNotes(core, [pending]), expected);
	assert.equal(craftedNotes(core, [datePending(pending, '2026-11-30'), pending]), expected);
	assert.equal(craftedNotes({ ...core, version: '2.0.1' }, [pending]), null);
	assert.equal(craftedNotes(core, ['', [...PREAMBLE, ...OLD].join('\n')]), null);
});

test('a pending section needs no major-notes insertion', () => {
	const root = mkdtempSync(join(tmpdir(), 'crafted-notes-'));
	mkdirSync(join(root, '.release-notes', '2.0'), { recursive: true });
	const core = pkg('demo-core', '2.0.0', 'viewers/demo/packages/core');
	assert.deepEqual(pendingNotes(root, core, crafted('demo-core', '2.0.0')), []);
	const generated = [
		...PREAMBLE,
		`## [2.0.0](${releaseUrl('demo-core', '2.0.0')}) - 2026-11-30`,
		'',
		...OLD,
	].join('\n');
	assert.deepEqual(pendingNotes(root, core, generated), ['2.0']);
});

function fixtureRepo(files) {
	const root = mkdtempSync(join(tmpdir(), 'crafted-changelogs-'));
	for (const [path, text] of Object.entries(files)) {
		mkdirSync(dirname(join(root, path)), { recursive: true });
		writeFileSync(join(root, path), text);
	}
	return { root, read: (path) => readFileSync(join(root, path), 'utf8') };
}

const PLAN = {
	order: ['demo-core', 'demo-react', 'demo-vue'],
	packages: {
		'demo-core': pkg('demo-core', '2.0.0', 'viewers/demo/packages/core'),
		'demo-react': pkg('demo-react', '3.0.0', 'viewers/demo/packages/react'),
		'demo-vue': pkg('demo-vue', '1.0.0', 'viewers/demo/packages/vue', { release: false }),
	},
};

test('the plan dates the hand-written sections of released packages only', () => {
	const plain = [...PREAMBLE, ...OLD].join('\n');
	const { root, read } = fixtureRepo({
		'viewers/demo/packages/core/CHANGELOG.md': crafted('demo-core', '2.0.0'),
		'viewers/demo/packages/react/CHANGELOG.md': plain,
		'viewers/demo/packages/vue/CHANGELOG.md': crafted('demo-vue', '2.0.0'),
	});
	assert.deepEqual(planProblems(root, PLAN), []);
	assert.deepEqual(pendingKeys(root, PLAN), ['demo-core']);
	assert.deepEqual(datePackages(root, PLAN, '2026-11-30'), [
		'viewers/demo/packages/core/CHANGELOG.md',
	]);
	assert.match(read('viewers/demo/packages/core/CHANGELOG.md'), /@2\.0\.0\) - 2026-11-30$/mu);
	assert.equal(read('viewers/demo/packages/react/CHANGELOG.md'), plain);
	assert.match(read('viewers/demo/packages/vue/CHANGELOG.md'), /- Unreleased$/mu);
	// Dated, the section still publishes this release as written, so git-cliff still skips it.
	assert.deepEqual(pendingKeys(root, PLAN), ['demo-core']);
	assert.deepEqual(datePackages(root, PLAN, '2026-12-01'), []);
});

/** `crafted` with its date already written. */
const dated = (npm, version, date = '2026-11-21') =>
	crafted(npm, version).replace(/ - Unreleased$/mu, ` - ${date}`);

test('a section that already carries its date is published as written', () => {
	const text = dated('demo-core', '2.0.0');
	const { root, read } = fixtureRepo({ 'viewers/demo/packages/core/CHANGELOG.md': text });
	assert.deepEqual(
		planProblems(root, PLAN, () => false),
		[],
	);
	assert.deepEqual(pendingKeys(root, PLAN), ['demo-core']);
	assert.deepEqual(
		datePackages(root, PLAN, '2026-11-30', () => false),
		[],
	);
	assert.equal(read('viewers/demo/packages/core/CHANGELOG.md'), text);
	const core = PLAN.packages['demo-core'];
	assert.match(craftedNotes(core, [text]), /^The lead paragraph\./u);
	assert.equal(craftedSectionFor(text, '2.0.0').date, '2026-11-21');
});

test('a dated section without the marker is an ordinary generated section', () => {
	const text = dated('demo-core', '2.0.0').replace('<!-- release-notes: 2.0 -->\n', '');
	assert.equal(datedCraftedSection(text), null);
	assert.equal(craftedSectionFor(text, '2.0.0'), null);
});

test('a released dated section is history; an unreleased one for another version fails', () => {
	const text = dated('demo-core', '2.0.0');
	const next = { ...PLAN.packages['demo-core'], version: '2.0.1' };
	assert.equal(craftedSectionFor(text, '2.0.1'), null);
	assert.deepEqual(
		packageProblems(next, text, 'CHANGELOG.md', () => true),
		[],
	);
	assert.match(
		packageProblems(next, text, 'CHANGELOG.md', () => false).join('\n'),
		/is for 2\.0\.0, which is not released, but this release plans demo-core@2\.0\.1/u,
	);
});

test('a mismatched section fails the dating step before any file is written', () => {
	const text = crafted('demo-core', '1.9.0');
	const { root, read } = fixtureRepo({
		'viewers/demo/packages/core/CHANGELOG.md': text,
		'viewers/demo/CHANGELOG.md': '# Changelog\n\n## Unreleased\n\n_Releases: demo-react@3.1.0_\n',
	});
	const problems = planProblems(root, PLAN);
	assert.equal(problems.length, 2);
	assert.match(problems[0], /viewers\/demo\/packages\/core\/CHANGELOG\.md: .* is for 1\.9\.0/u);
	assert.match(problems[1], /viewers\/demo\/CHANGELOG\.md: .* demo-react@3\.1\.0/u);
	assert.throws(() => datePackages(root, PLAN, '2026-11-30'), /is for 1\.9\.0/u);
	assert.equal(read('viewers/demo/packages/core/CHANGELOG.md'), text);
});
