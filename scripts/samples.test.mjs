/**
 * The sample documents (`samples/`): the generators typecheck and are deterministic,
 * `manifest.json` describes exactly the committed files, the gallery page is built from it, and
 * the suite only ever opens a file the manifest lists, with the size and hash it records.
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
	SAMPLE_ID,
	downloadSample,
	findSample,
	sha256Hex,
	showSampleNotice,
	takeSampleLink,
	untouchedCopy,
} from '../apps/office-suite/src/samples.js';
import {
	PRODUCTS,
	checkManifest,
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

test('the generators typecheck against the core sources they import', { timeout: 600_000 }, () => {
	// samples/tsconfig.json: the core's strict flags over the generators. A model or writer change
	// that breaks a generator fails here (the CI plan runs this for the areas they import).
	const run = spawnSync('bun x tsc -p samples/tsconfig.json', {
		cwd: ROOT,
		encoding: 'utf8',
		shell: true,
	});
	assert.equal(run.status, 0, `${run.stdout}\n${run.stderr}`);
});

test('a manifest entry the page or the suite could not trust is refused', () => {
	const good = manifest.samples[0];
	assert.equal(checkManifest({ samples: [good] }).samples.length, 1);
	const refused = (change, why) =>
		assert.throws(() => checkManifest({ samples: [{ ...good, ...change }] }), why);
	// The id becomes an element id and a link in the page.
	for (const id of ['A', 'a b', 'a"onclick="x', '../a', '', 'a--b', '-a'])
		refused({ id }, /no id|not a lower-case slug/);
	refused({ features: 'Charts' }, /list of feature strings/);
	refused({ features: [] }, /list of feature strings/);
	refused({ features: ['Charts', 3] }, /list of feature strings/);
	for (const size of [0, -1, 1.5, '1024', Number.NaN, undefined])
		refused({ size }, /size in bytes/);
	refused({ sha256: 'abc' }, /sha256/);
	refused({ sha256: undefined }, /no sha256/);
	refused({ file: '../x.docx' }, /unexpected file path/);
	refused({ thumbnail: 'https://example.invalid/x.webp' }, /unexpected thumbnail path/);
	refused({ product: 'onenote' }, /names the product/);
	assert.throws(() => checkManifest({ samples: [good, good] }), /listed twice/);
	assert.throws(() => checkManifest({}), /lists no samples/);
});

test('the desktop app is built without the sample files', () => {
	const tauri = JSON.parse(readFileSync(join(ROOT, 'desktop/src-tauri/tauri.conf.json'), 'utf8'));
	assert.match(tauri.build.beforeBuildCommand, /build:suite:desktop$/);
	const scripts = JSON.parse(readFileSync(join(ROOT, 'package.json'), 'utf8')).scripts;
	assert.match(scripts['build:suite:desktop'], /build-suite\.mjs --no-samples$/);
	// The other builds keep the gallery.
	assert.doesNotMatch(scripts['build:suite'], /--no-samples/);
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
	for (const sample of manifest.samples) assert.match(sample.id, SAMPLE_ID);
});

test('the suite fetches only the manifest and the file it lists, from its own site', async () => {
	const base = 'https://example.invalid/ooxml/apps/word/../../';
	const sample = manifest.samples[0];
	const bytes = readFileSync(join(SAMPLES, sample.file));
	const requested = [];
	const listing =
		(samples, body = bytes) =>
		async (url) => {
			requested.push(String(url));
			if (String(url).endsWith('/manifest.json'))
				return { ok: true, json: async () => ({ samples }) };
			return { ok: true, arrayBuffer: async () => new Uint8Array(body).buffer };
		};
	const fetch = listing(manifest.samples);
	const found = await findSample(sample.id, { base, fetch });
	assert.deepEqual(found, {
		name: sample.file.split('/')[1],
		url: `https://example.invalid/ooxml/samples/${sample.file}`,
		size: sample.size,
		sha256: sample.sha256,
	});
	assert.equal((await downloadSample(found, { fetch })).length, sample.size);
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
			findSample('x', { base, fetch: listing([{ ...sample, id: 'x', file }]) }),
			/not available/,
			file,
		);
	// So is an entry without a usable size or hash: the file could not be checked.
	for (const change of [{ size: 0 }, { size: '12' }, { sha256: 'abc' }, { sha256: undefined }])
		await assert.rejects(
			findSample(sample.id, { base, fetch: listing([{ ...sample, ...change }]) }),
			/not available/,
			JSON.stringify(change),
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

test('a downloaded file that is not the one the manifest describes is refused', async () => {
	const sample = manifest.samples[0];
	const bytes = readFileSync(join(SAMPLES, sample.file));
	const found = {
		url: 'https://example.invalid/s.docx',
		size: sample.size,
		sha256: sample.sha256,
	};
	const serving = (body) => async () => ({
		ok: true,
		arrayBuffer: async () => new Uint8Array(body).buffer,
	});
	assert.equal(await sha256Hex(bytes), sample.sha256);
	// A stale or truncated copy has another length.
	await assert.rejects(
		downloadSample(found, { fetch: serving(bytes.subarray(1)) }),
		/not available/,
	);
	// One flipped byte keeps the length and changes the hash.
	const flipped = Uint8Array.from(bytes);
	flipped[flipped.length - 1] ^= 1;
	await assert.rejects(downloadSample(found, { fetch: serving(flipped) }), /not available/);
	// Where the page has no Web Crypto (plain http), the size is all that can be compared.
	assert.equal(await sha256Hex(bytes, null), undefined);
	assert.equal(
		(await downloadSample(found, { fetch: serving(flipped), subtle: null })).length,
		sample.size,
	);
});

test('an unedited earlier copy with the published bytes is reused', async () => {
	const sample = manifest.samples[0];
	const bytes = new Uint8Array(readFileSync(join(SAMPLES, sample.file)));
	const found = {
		url: 'https://example.invalid/ooxml/samples/docx/a.docx',
		size: sample.size,
		sha256: sample.sha256,
	};
	const copy = { id: '1', sourceUrl: found.url, revision: 1, bytes };
	assert.equal(await untouchedCopy([copy], found), copy);
	// The bytes decide, not where the copy came from: the same file opened by hand is that sample.
	const imported = { id: '2', revision: 1, bytes };
	assert.equal(await untouchedCopy([imported], found), imported);
	assert.equal(await untouchedCopy([{ ...copy, revision: 2 }], found), undefined);
	assert.equal(await untouchedCopy([{ ...copy, cloud: {} }], found), undefined);
	assert.equal(await untouchedCopy([{ ...copy, parent: { id: 'p' } }], found), undefined);
	assert.equal(await untouchedCopy([{ ...copy, bytes: bytes.subarray(1) }], found), undefined);
	// The same address and size with other content (an older build of the sample) is not it.
	const older = Uint8Array.from(bytes);
	older[0] ^= 1;
	assert.equal(await untouchedCopy([{ ...copy, bytes: older }], found), undefined);
	assert.equal(await untouchedCopy([{ ...copy, bytes: older }, copy], found), copy);
	// Without Web Crypto, only a copy fetched from the sample's own address is taken for it.
	assert.equal(await untouchedCopy([imported], found, { subtle: null }), undefined);
	assert.equal(await untouchedCopy([copy], found, { subtle: null }), copy);
});

test('a sample that cannot be opened is said in a notice with a link to the gallery', () => {
	const element = (tag) => ({
		tag,
		children: [],
		append(...nodes) {
			this.children.push(...nodes);
		},
		prepend(node) {
			this.children.unshift(node);
		},
		remove() {
			this.removed = true;
		},
	});
	let previous;
	const document = {
		createElement: element,
		getElementById: (id) => (id === 'sample-notice' ? previous : undefined),
	};
	const host = element('div');
	const base = 'https://example.invalid/ooxml/apps/word/../../';
	const notice = showSampleNotice('That sample is not available.', { base, host, document });
	assert.equal(host.children[0], notice);
	assert.equal(notice.id, 'sample-notice');
	const [text, browse, close] = notice.children;
	assert.equal(text.textContent, 'That sample is not available.');
	assert.equal(browse.tag, 'a');
	assert.equal(browse.textContent, 'Browse samples');
	assert.equal(browse.href, 'https://example.invalid/ooxml/samples/');
	// The status line announces the message; the notice is not a second live region.
	assert.equal(notice.role, undefined);
	close.onclick();
	assert.equal(notice.removed, true);
	// A second failure replaces the notice instead of stacking another.
	previous = notice;
	notice.removed = false;
	showSampleNotice('The samples could not be loaded.', { base, host, document });
	assert.equal(notice.removed, true);
});
