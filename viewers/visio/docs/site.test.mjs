// Guards the VitePress documentation site: it must build, every page must render with one heading,
// internal links must resolve, the capability ledger must keep its inventory, and the honesty
// statements about the beta must not be lost. The demos are checked by the browser tests.
import { before, describe, it, after } from 'node:test';
import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { existsSync, mkdtempSync, readFileSync, readdirSync, rmSync, statSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { JSDOM } from 'jsdom';
import { FRAMEWORK_DEMOS } from '../scripts/framework-demos.mjs';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const BASE = '/visio-viewer/';
const out = mkdtempSync(join(tmpdir(), 'visio-docs-'));
const read = (path) => readFileSync(resolve(root, path), 'utf8');
const frameworks = ['react', 'vue', 'angular', 'svelte', 'solid', 'vanilla'];
const pages = [
	'index',
	'getting-started',
	'demos',
	'architecture',
	'bindings',
	'api',
	'theming',
	'collaboration',
	'parity',
	'verification',
	'corpus',
	'suite-provenance',
	'releasing',
	...frameworks.map((id) => `frameworks/${id}`),
];
const html = (page) => readFileSync(join(out, `${page}.html`), 'utf8');
const text = (page) => new JSDOM(html(page)).window.document.body.textContent.replace(/\s+/g, ' ');

before(() => {
	const cli = resolve(root, 'docs/node_modules/vitepress/bin/vitepress.js');
	assert.ok(existsSync(cli), 'Install the docs dependencies with `npm ci --prefix docs`.');
	const result = spawnSync(process.execPath, [cli, 'build', 'docs'], {
		cwd: root,
		env: { ...process.env, DOCS_OUT: out },
		encoding: 'utf8',
	});
	assert.equal(result.status, 0, `VitePress build failed:\n${result.stdout}\n${result.stderr}`);
});
after(() => rmSync(out, { recursive: true, force: true }));

describe('documentation site build', () => {
	for (const page of pages) {
		it(`${page}: renders one heading and unique ids`, () => {
			const { document } = new JSDOM(html(page)).window;
			assert.equal(document.documentElement.lang, 'en-US');
			if (page !== 'index') assert.equal(document.querySelectorAll('.vp-doc h1').length, 1);
			const ids = [...document.querySelectorAll('[id]')].map((node) => node.id);
			assert.equal(ids.length, new Set(ids).size, 'IDs must be unique');
			for (const image of document.images) assert.ok(image.hasAttribute('alt'));
		});
		it(`${page}: internal links resolve`, () => {
			const { document } = new JSDOM(html(page)).window;
			for (const element of document.querySelectorAll('a[href], iframe[src]')) {
				const value = element.getAttribute('href') ?? element.getAttribute('src');
				if (!value.startsWith(BASE)) continue;
				const path = decodeURIComponent(value.slice(BASE.length).split(/[?#]/)[0]);
				// Demos are built after VitePress; their routes are checked by the list test below.
				if (/^demo(-[a-z]+)?\//.test(path)) continue;
				const candidates = [path, `${path}.html`, join(path, 'index.html')];
				assert.ok(
					candidates.some((c) => existsSync(join(out, c)) && statSync(join(out, c)).isFile()),
					`${page}: missing target ${value}`,
				);
			}
		});
	}

	it('keeps the public beta and fidelity limitations explicit', () => {
		const home = text('index');
		assert.match(home, /public beta/i);
		assert.match(home, /not Microsoft Visio parity/i);
		assert.match(home, /npm install visio-react-viewer/);
		const start = text('getting-started');
		assert.match(start, /General drawing/);
		assert.match(
			start,
			/Native Visio reopening and lossless round-trip fidelity remain unverified/,
		);
		assert.doesNotMatch(home + start, /private and unpublished|MIT license/i);
		assert.match(text('parity'), /parity is a target, not a current claim/i);
		assert.match(text('collaboration'), /same-browser only/i);
		assert.match(text('collaboration'), /Sharing across devices/);
	});

	it('lists every demo route and every framework package', () => {
		const demos = read('docs/demos.md');
		const routes = ['/demo/', '/demo-vanilla/', ...FRAMEWORK_DEMOS.map((d) => `/demo-${d.id}/`)];
		for (const route of routes) assert.ok(demos.includes(route), route);
		const start = read('docs/getting-started.md');
		for (const id of frameworks) {
			assert.ok(start.includes(`npm install visio-${id}-viewer`), `install command for ${id}`);
			assert.ok(existsSync(resolve(root, `docs/frameworks/${id}.md`)), `framework guide ${id}`);
		}
	});

	it('has landing code samples for every framework', () => {
		const samples = read('docs/.vitepress/theme/landing/code/samples.ts');
		for (const id of frameworks)
			assert.ok(samples.includes(`entry: 'visio-${id}-viewer'`), `sample for ${id}`);
	});

	it('loads no third-party scripts and keeps em dashes out of the sources', () => {
		const { document } = new JSDOM(html('index')).window;
		for (const script of document.querySelectorAll('script[src]'))
			assert.ok(!/^(https?:)?\/\//.test(script.getAttribute('src')), 'No remote scripts');
		for (const file of readdirSync(resolve(root, 'docs')).filter((f) => f.endsWith('.md')))
			assert.ok(
				!read(`docs/${file}`).includes(String.fromCharCode(0x2014)),
				`${file} contains an em dash`,
			);
	});
});

describe('capability ledger', () => {
	const lines = read('docs/parity.md')
		.split('\n')
		.filter((line) => line.startsWith('|'));
	const cells = (line) =>
		line
			.slice(1, -1)
			.split(/(?<!\\)\|/)
			.map((cell) => cell.trim());
	const rows = lines.filter((line) => !/^\|[\s:|-]+\|$/.test(line)).map(cells);
	it('keeps the five-column inventory', () => {
		for (const row of rows) assert.equal(row.length, 5, `Row has the wrong width: ${row[0]}`);
		assert.ok(rows.length - 1 >= 29, 'The ledger must retain the established capability inventory');
	});
	it('has unique capability names and the required rows', () => {
		const names = rows.slice(1).map((row) => row[0]);
		assert.equal(names.length, new Set(names).size, 'Capability names must be unique');
		for (const required of ['ShapeSheet formulas', 'Corner rounding', 'Visio visual parity'])
			assert.ok(names.includes(required), `Missing capability: ${required}`);
	});
	it('names the evidence and what is missing for every row', () => {
		for (const row of rows.slice(1))
			assert.ok(row[2].length > 0 && row[3].length > 0 && row[4].length > 0, row[0]);
	});
});
