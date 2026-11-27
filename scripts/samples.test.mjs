/**
 * The sample documents (`samples/`): the generators are deterministic, `manifest.json` describes
 * exactly the committed files, the gallery page is built from it, and the suite only ever opens a
 * file the manifest lists.
 */
import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { existsSync, mkdtempSync, readdirSync, readFileSync, rmSync, statSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { test } from 'node:test';
import { fileURLToPath } from 'node:url';
import {
	downloadSample,
	findSample,
	takeSampleLink,
	untouchedCopy,
} from '../apps/office-suite/src/samples.js';
import {
	PRODUCTS,
	formatSize,
	readManifest,
	renderSamplesPage,
	writeSamplesPage,
} from './build-samples-page.mjs';

const ROOT = resolve(fileURLToPath(new URL('..', import.meta.url)));
const SAMPLES = join(ROOT, 'samples');
const EXTENSION = { docx: '.docx', xlsx: '.xlsx', pptx: '.pptx', visio: '.vsdx' };
/** The size limits the gallery promises: no file a visitor would wait long for. */
const FILE_LIMIT = 1.5 * 1024 * 1024;
const TOTAL_LIMIT = 15 * 1024 * 1024;
/** U+2014, which the repository's style bans; built from its code so this file holds none. */
const EM_DASH = String.fromCharCode(0x2014);

const manifest = readManifest();

test('every generator writes the same bytes twice', { timeout: 600_000 }, () => {
	// One command line through the shell, which is how Windows finds bun's launcher.
	const run = spawnSync('bun samples/generators/build.ts --determinism', {
		cwd: ROOT,
		encoding: 'utf8',
		shell: true,
	});
	assert.equal(run.status, 0, `${run.stdout}\n${run.stderr}`);
	// One line per sample: a generator that was skipped proves nothing.
	for (const sample of manifest.samples)
		assert.match(run.stdout, new RegExp(`generated twice ${sample.file.replace('.', '\\.')} `));
});

test('the manifest describes exactly the committed files', () => {
	assert.ok(manifest.samples.length >= 12, 'three samples per product');
	const ids = new Set();
	let total = 0;
	for (const sample of manifest.samples) {
		assert.ok(!ids.has(sample.id), `${sample.id} is listed twice`);
		ids.add(sample.id);
		assert.equal(sample.file, `${sample.product}/${sample.id}${EXTENSION[sample.product]}`);
		assert.equal(sample.thumbnail, `thumbnails/${sample.id}.webp`);
		const bytes = readFileSync(join(SAMPLES, sample.file));
		assert.equal(bytes.length, sample.size, `${sample.file}: size`);
		assert.equal(
			createHash('sha256').update(bytes).digest('hex'),
			sample.sha256,
			`${sample.file}: run bun samples/generators/build.ts`,
		);
		assert.ok(bytes.length <= FILE_LIMIT, `${sample.file} is over 1.5 MB`);
		assert.equal(bytes.subarray(0, 2).toString('latin1'), 'PK', `${sample.file} is a ZIP package`);
		total += bytes.length;
		const thumbnail = readFileSync(join(SAMPLES, sample.thumbnail));
		assert.equal(thumbnail.subarray(8, 12).toString('latin1'), 'WEBP', sample.thumbnail);
		assert.ok(thumbnail.length <= 150 * 1024, `${sample.thumbnail} is over 150 KB`);
		assert.ok(sample.features.length >= 3, `${sample.id} lists its features`);
		const copy = [sample.title, sample.description, ...sample.features].join(' ');
		assert.ok(!copy.includes(EM_DASH), `${sample.id}: no em-dashes`);
	}
	assert.ok(total <= TOTAL_LIMIT, 'the samples total more than 15 MB');
	for (const product of PRODUCTS) {
		const listed = manifest.samples
			.filter((sample) => sample.product === product.id)
			.map((sample) => sample.file.split('/')[1])
			.sort();
		assert.ok(listed.length >= 3, `${product.app} has three samples`);
		assert.deepEqual(readdirSync(join(SAMPLES, product.id)).sort(), listed, product.id);
	}
	assert.deepEqual(
		readdirSync(join(SAMPLES, 'thumbnails')).sort(),
		manifest.samples.map((sample) => `${sample.id}.webp`).sort(),
	);
});

test('the gallery page has a card, an open link and a download for every sample', () => {
	const template = readFileSync(join(ROOT, 'site/samples/template.html'), 'utf8');
	const html = renderSamplesPage(template, manifest);
	assert.doesNotMatch(html, /<!-- samples:/);
	for (const sample of manifest.samples) {
		const app = PRODUCTS.find((product) => product.id === sample.product).app;
		const card = html.slice(html.indexOf(`id="sample-${sample.id}"`));
		assert.ok(card.length > 0 && card.length < html.length, `${sample.id} has a card`);
		const own = card.slice(0, card.indexOf('</article>'));
		assert.match(own, new RegExp(`href="\\.\\./\\?sample=${sample.id}">Open in ${app}<`));
		assert.ok(own.includes(`href="${sample.file}" download`), `${sample.id} downloads`);
		assert.ok(own.includes(`src="${sample.thumbnail}"`), `${sample.id} shows its thumbnail`);
		assert.ok(own.includes(formatSize(sample.size)), `${sample.id} shows its size`);
	}
	for (const product of PRODUCTS)
		assert.ok(html.includes(`<section id="${product.app.toLowerCase()}"`), product.app);
	// The page is static: the hub's script is the only one, so the suite's budgets do not move.
	assert.deepEqual(
		[...html.matchAll(/<script[^>]*src="([^"]+)"/g)].map((match) => match[1]),
		['../appearance-init.js', '../docs/docs.js'],
	);
	assert.equal(formatSize(900), '1 KB');
	assert.equal(formatSize(421_000), '411 KB');
	assert.equal(formatSize(1_300_000), '1.2 MB');
});

test('the build writes the page, the manifest and every file under samples/', () => {
	const out = mkdtempSync(join(tmpdir(), 'ooxml-samples-page-'));
	try {
		const ids = writeSamplesPage(out);
		assert.deepEqual(
			ids,
			manifest.samples.map((sample) => sample.id),
		);
		for (const name of ['index.html', 'samples.css', 'manifest.json'])
			assert.ok(existsSync(join(out, 'samples', name)), name);
		assert.ok(!existsSync(join(out, 'samples', 'template.html')));
		for (const sample of manifest.samples) {
			assert.equal(statSync(join(out, 'samples', sample.file)).size, sample.size);
			assert.ok(existsSync(join(out, 'samples', sample.thumbnail)));
		}
	} finally {
		rmSync(out, { recursive: true, force: true });
	}
});

test('a sample link names a manifest entry and is taken out of the address', () => {
	const base = 'https://example.invalid/ooxml/';
	assert.equal(takeSampleLink(`${base}?suite#/`), null);
	assert.deepEqual(takeSampleLink(`${base}?sample=technical-report&suite#/chat`), {
		id: 'technical-report',
		href: `${base}?suite=#/chat`,
	});
	for (const bad of ['../secret', 'A', 'a/b', 'a.docx', '', 'https://example.invalid/x'])
		assert.equal(takeSampleLink(`${base}?sample=${encodeURIComponent(bad)}`).id, null, bad);
});

test('the suite fetches only the manifest and the file it lists, from its own site', async () => {
	const base = 'https://example.invalid/ooxml/apps/word/../../';
	const sample = manifest.samples[0];
	const requested = [];
	const listing = (samples) => async (url) => {
		requested.push(String(url));
		if (String(url).endsWith('/manifest.json'))
			return { ok: true, json: async () => ({ samples }) };
		return { ok: true, arrayBuffer: async () => new Uint8Array([80, 75, 3, 4]).buffer };
	};
	const fetch = listing(manifest.samples);
	const found = await findSample(sample.id, { base, fetch });
	assert.deepEqual(found, {
		name: sample.file.split('/')[1],
		url: `https://example.invalid/ooxml/samples/${sample.file}`,
		size: sample.size,
	});
	assert.deepEqual([...(await downloadSample(found, { fetch }))], [80, 75, 3, 4]);
	assert.deepEqual(requested, [
		'https://example.invalid/ooxml/samples/manifest.json',
		`https://example.invalid/ooxml/samples/${sample.file}`,
	]);
	await assert.rejects(findSample('no-such-sample', { base, fetch }), /not available/);
	// A manifest that points outside samples/ (a tampered deployment) is refused, not followed.
	for (const file of [
		'../index.html',
		'https://example.invalid/x.docx',
		'docx/../../x.docx',
		'x.js',
	])
		await assert.rejects(
			findSample('x', { base, fetch: listing([{ id: 'x', file }]) }),
			/not available/,
			file,
		);
	await assert.rejects(
		findSample(sample.id, { base, fetch: async () => ({ ok: false }) }),
		/could not be loaded/,
	);
	await assert.rejects(
		findSample(sample.id, {
			base,
			fetch: async () => {
				throw new TypeError('offline');
			},
		}),
		/could not be loaded/,
	);
	await assert.rejects(
		downloadSample(found, { fetch: async () => ({ ok: false }) }),
		/not available/,
	);
});

test('an unedited earlier copy of a sample is reused', () => {
	const sample = { url: 'https://example.invalid/ooxml/samples/docx/a.docx', size: 3 };
	const bytes = new Uint8Array(3);
	const copy = { id: '1', sourceUrl: sample.url, revision: 1, bytes };
	assert.equal(untouchedCopy([copy], sample), copy);
	assert.equal(untouchedCopy([{ ...copy, revision: 2 }], sample), undefined);
	assert.equal(untouchedCopy([{ ...copy, bytes: new Uint8Array(4) }], sample), undefined);
	assert.equal(
		untouchedCopy([{ ...copy, sourceUrl: 'https://example.invalid/b' }], sample),
		undefined,
	);
	assert.equal(untouchedCopy([{ ...copy, cloud: {} }], sample), undefined);
});
