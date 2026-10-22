import assert from 'node:assert/strict';
import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join } from 'node:path';
import { after, test } from 'node:test';
import {
	anchorsOf,
	checkCodePaths,
	checkLinks,
	checkSidebars,
	checkTerms,
	checkTranslations,
	globToRegExp,
	renderReport,
	runChecks,
} from './check-docs.mjs';

const dirs = [];
function fixture(files) {
	const root = mkdtempSync(join(tmpdir(), 'check-docs-'));
	dirs.push(root);
	for (const [path, text] of Object.entries(files)) {
		mkdirSync(dirname(join(root, path)), { recursive: true });
		writeFileSync(join(root, path), text);
	}
	return root;
}
after(() => {
	for (const d of dirs) rmSync(d, { recursive: true, force: true });
});

test('relative links and images must resolve; code spans and fences are ignored', () => {
	const root = fixture({ 'docs/a.md': '# A\n', 'docs/img/x.png': '' });
	const text = [
		'[ok](a.md) ![ok](img/x.png) [dead](missing.md) [web](https://example.com/x.md)',
		'`[not](a-link.md)`',
		'```',
		'[fenced](nope.md)',
		'```',
	].join('\n');
	const { errors } = checkLinks(root, 'docs/b.md', text, new Map());
	assert.deepEqual(
		errors.map((e) => e.message),
		['Dead link `missing.md`'],
	);
});

test('VitePress absolute links resolve against the docs root, with public and index pages', () => {
	const root = fixture({
		'viewers/p/docs/guide/index.md': '# G\n',
		'viewers/p/docs/guide/foo.md': '# Foo\n',
		'viewers/p/docs/public/logo.svg': '',
	});
	const text = '[a](/guide/) [b](/guide/foo) [c](/guide/foo.html) [d](/logo.svg) [e](/guide/gone)';
	const { errors } = checkLinks(root, 'viewers/p/docs/x.md', text, new Map());
	assert.deepEqual(
		errors.map((e) => e.message),
		['Dead link `/guide/gone`'],
	);
});

test('generated routes are not reported', () => {
	const root = fixture({ 'viewers/p/docs/x.md': '' });
	const text = '[demo](/demo-react/)';
	assert.equal(checkLinks(root, 'viewers/p/docs/x.md', text, new Map()).errors.length, 1);
	const generated = [/^\/demo(?:-[a-z]+)?\//];
	assert.equal(
		checkLinks(root, 'viewers/p/docs/x.md', text, new Map(), generated).errors.length,
		0,
	);
});

test('anchors follow GitHub and VitePress slugs and are warnings', () => {
	const root = fixture({
		'a.md': '# Big Title\n## Calls over NAT (TURN)\n## Über uns\n## Dup\n## Dup\n',
	});
	const set = anchorsOf('# Big Title\n## Calls over NAT (TURN)\n## Über uns\n## Dup\n## Dup\n');
	for (const id of ['big-title', 'calls-over-nat-turn', 'über-uns', 'uber-uns', 'dup', 'dup-1']) {
		assert.ok(set.has(id), id);
	}
	const { errors, warnings } = checkLinks(
		root,
		'b.md',
		'[ok](a.md#big-title) [bad](a.md#nope)',
		new Map(),
	);
	assert.equal(errors.length, 0);
	assert.equal(warnings.length, 1);
	assert.equal(warnings[0].check, 'anchor');
});

test('inline-code repository paths must exist; globs, placeholders and build output pass', () => {
	const root = fixture({
		'src/core/xml/index.ts': '',
		'scripts/a.mjs': '',
		'viewers/p/packages/core/src/x.ts': '',
	});
	const text = [
		'`src/core/xml/index.ts` `src/core/*/index.ts` `scripts/<name>.mjs` `src/core/{xml,gone}/index.ts`',
		'`src/core/missing.ts` `scripts/gone.mjs:12` `src/core/dist/x.js`',
		'`not/a/known/dir.ts` `npm run src/foo bar`',
		'```',
		'src/core/fenced-missing.ts',
		'```',
	].join('\n');
	const found = checkCodePaths(root, 'README.md', text).map((e) => e.message);
	assert.deepEqual(found, [
		'Path `src/core/{xml,gone}/index.ts` does not exist',
		'Path `src/core/missing.ts` does not exist',
		'Path `scripts/gone.mjs` does not exist',
	]);
	// A user's own project file in a doc folder is not a repository path.
	assert.equal(checkCodePaths(root, 'docs/a.md', '`src/main.ts`').length, 0);
	// A viewer doc resolves packages/ against the viewer.
	const v = checkCodePaths(
		root,
		'viewers/p/README.md',
		'`packages/core/src/x.ts` `packages/gone/x.ts`',
	);
	assert.equal(v.length, 1);
	// Files that record history are skipped.
	assert.equal(
		checkCodePaths(root, 'PROVENANCE.md', '`src/gone.ts/x`', { files: ['PROVENANCE.md'] }).length,
		0,
	);
});

test('stale terms come from rules, honour allow-lists and the intended marker', () => {
	const rules = [
		{
			id: 'scope',
			pattern: '@old/core',
			message: 'old',
			allow: ['AGENTS.md', 'viewers/*/docs/**'],
		},
		{ id: 'dash', pattern: '\\u2014', message: 'dash' },
	];
	const dash = String.fromCharCode(0x2014);
	assert.equal(checkTerms('docs/a.md', 'use @old/core', rules).length, 1);
	assert.equal(checkTerms('AGENTS.md', 'use @old/core', rules).length, 0);
	assert.equal(checkTerms('viewers/x/docs/deep/a.md', 'use @old/core', rules).length, 0);
	assert.equal(checkTerms('a.md', `a ${dash} b`, rules)[0].line, 1);
	assert.equal(checkTerms('a.md', `a ${dash} b em-dash: intended`, rules).length, 0);
	assert.deepEqual(globToRegExp('a/**/b.md').test('a/x/y/b.md'), true);
});

test('sidebar links must exist; English pages in no sidebar are warnings', () => {
	const root = fixture({
		'viewers/p/docs/index.md': '',
		'viewers/p/docs/guide/index.md': '',
		'viewers/p/docs/guide/in.md': '',
		'viewers/p/docs/guide/orphan.md': '',
		'viewers/p/docs/zh/guide/orphan.md': '',
		'viewers/p/docs/.vitepress/config.ts': [
			"nav: [{ text: 'G', link: '/guide/' }, { text: 'Ext', link: 'https://example.com' }]",
			"sidebar: [{ text: 'In', link: '/guide/in' }, { text: 'Gone', link: '/guide/gone' }]",
		].join('\n'),
	});
	const files = [
		'viewers/p/docs/index.md',
		'viewers/p/docs/guide/index.md',
		'viewers/p/docs/guide/in.md',
		'viewers/p/docs/guide/orphan.md',
		'viewers/p/docs/zh/guide/orphan.md',
		'viewers/p/docs/.vitepress/config.ts',
	];
	const { errors, warnings } = checkSidebars(root, 'viewers/p/docs', files);
	assert.deepEqual(
		errors.map((e) => e.message),
		['Dead sidebar/nav link `/guide/gone`'],
	);
	assert.deepEqual(
		warnings.map((w) => w.file),
		['viewers/p/docs/guide/orphan.md'],
	);
});

test('translation lag counts missing and older locale pages', () => {
	const files = [
		'viewers/p/docs/a.md',
		'viewers/p/docs/b.md',
		'viewers/p/docs/c.md',
		'viewers/p/docs/zh/a.md',
		'viewers/p/docs/zh/b.md',
	];
	const times = new Map([
		['viewers/p/docs/a.md', 200],
		['viewers/p/docs/zh/a.md', 100],
		['viewers/p/docs/b.md', 100],
		['viewers/p/docs/zh/b.md', 200],
	]);
	const [zh] = checkTranslations('.', 'viewers/p/docs', files, times);
	assert.equal(zh.locale, 'zh');
	assert.equal(zh.total, 3);
	assert.deepEqual(zh.missing, ['c.md']);
	assert.deepEqual(zh.stale, ['a.md']);
});

test('runChecks combines the checks and the report separates errors from warnings', () => {
	const root = fixture({
		'README.md': '[dead](gone.md) `src/gone.ts` and @old/core\n',
		'CHANGELOG.md': '[dead](gone.md)\n',
	});
	const config = { rules: [{ id: 'scope', pattern: '@old/core', message: 'old' }] };
	const result = runChecks({ root, files: ['README.md', 'CHANGELOG.md'], config });
	assert.deepEqual(result.errors.map((e) => e.check).sort(), ['link', 'path', 'term:scope']);
	const report = renderReport(result);
	assert.match(report, /3 error\(s\), 0 warning\(s\)/);
	assert.match(report, /### Dead links \(1\)/);
});
