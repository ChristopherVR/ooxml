import assert from 'node:assert/strict';
import test, { before } from 'node:test';
import { createRequire } from 'node:module';
import { readFileSync } from 'node:fs';
import { pathToFileURL } from 'node:url';
import { createHash } from 'node:crypto';
import { emf, record, plus, comment, syntheticCases } from './fixtures.mjs';
import { LIMITS, preflight } from './preflight.mjs';
import { sanitizeVectorTree } from './sanitize.mjs';
import { runIsolated } from './isolated.mjs';

const cases = syntheticCases();

before(() => {
	const core = createRequire(new URL('../../../ooxml/package.json', import.meta.url));
	const entry = core.resolve('emf-converter');
	const metadata = JSON.parse(
		readFileSync(new URL('../package.json', pathToFileURL(entry)), 'utf8'),
	);
	assert.equal(
		metadata.version,
		'3.5.1',
		'Characterization is pinned to 3.5.1; re-audit another version.',
	);
	assert.equal(
		createHash('sha256').update(readFileSync(entry)).digest('hex'),
		'2127569d3337d19f98e44344fb51af914a9dbb830f6ded8170983d779d9b8527',
		'Installed bundle differs from the audited artifact.',
	);
});

test('strict preflight admits the bounded rectangle subset', () => {
	assert.equal(preflight(cases['plain-rectangle']).eligible, true);
});

test('preflight rejects unknown records and every EMF+ stream explicitly', () => {
	const unsupported = preflight(emf([record(0x12345678)]));
	assert.equal(unsupported.structurallyValid, true);
	assert.equal(unsupported.diagnostics[0].code, 'emf.unsupported-record');
	assert.equal(preflight(cases['driver-even-count']).eligible, false);
	assert.ok(
		preflight(cases['driver-even-count']).diagnostics.some(
			(d) => d.code === 'emf.known-fidelity-risk',
		),
	);
});

test('preflight rejects bad signatures, versions, counts, sizes and trailing bytes', () => {
	for (const [offset, value] of [
		[40, 0],
		[44, 0],
		[48, 1],
		[52, 100],
		[112, 0],
		[112, 9],
		[112, 0xffffffff],
	]) {
		const bytes = Buffer.from(cases['plain-rectangle']);
		bytes.writeUInt32LE(value, offset);
		assert.equal(preflight(bytes).structurallyValid, false, `offset ${offset}`);
	}
	assert.equal(preflight(cases['plain-rectangle'].subarray(0, -20)).structurallyValid, false);
	assert.equal(
		preflight(Buffer.concat([cases['plain-rectangle'], Buffer.alloc(4)])).structurallyValid,
		false,
	);
});

test('preflight checks comment and EMF+ record-local boundaries', () => {
	const bytes = emf([comment([plus(0x4002)])]);
	bytes.writeUInt32LE(400, 116);
	assert.ok(preflight(bytes).diagnostics.some((d) => d.code === 'emf.comment-range'));
	const child = emf([comment([plus(0x4002)])]);
	child.writeUInt32LE(64, 132);
	assert.ok(preflight(child).diagnostics.some((d) => d.code === 'emf.plus-record-size'));
});

test('preflight has aggregate record, input and state-depth budgets', () => {
	assert.equal(preflight(new Uint8Array(LIMITS.inputBytes + 1)).eligible, false);
	assert.equal(preflight(cases['plain-rectangle'], { ...LIMITS, records: 4 }).eligible, false);
	const deep = emf(Array.from({ length: 65 }, () => record(33)));
	assert.ok(preflight(deep).diagnostics.some((d) => d.code === 'emf.state-depth'));
	const many = emf([comment(Array.from({ length: 10 }, () => plus(0x4002)))]);
	assert.ok(
		preflight(many, { ...LIMITS, records: 6 }).diagnostics.some(
			(d) => d.code === 'emf.plus-record-limit',
		),
	);
});

test('preflight stays bounded for deterministic truncations and header mutations', () => {
	const source = cases['driver-odd-padded-data'];
	for (let end = 0; end <= source.length; end++)
		assert.doesNotThrow(() => preflight(source.subarray(0, end)));
	let seed = 0x3451;
	for (let i = 0; i < 500; i++) {
		seed = (Math.imul(seed, 1664525) + 1013904223) >>> 0;
		const bytes = Buffer.from(source);
		bytes.writeUInt32LE(seed, (seed % Math.floor(bytes.length / 4)) * 4);
		assert.doesNotThrow(() => preflight(bytes));
	}
});

test('safe path returns neutral commands and refuses unsupported content', async () => {
	const result = await runIsolated(cases['plain-rectangle'], 'safe');
	assert.equal(result.status, 'ok');
	assert.equal(result.vector.kind, 'vector');
	assert.equal(result.vector.items[0].kind, 'path');
	assert.equal(result.vector.items[0].paint.fill, '#ff0000');
	assert.equal((await runIsolated(cases['driver-even-count'], 'safe')).status, 'unsupported');
});

const root = (child) => ({
	tag: 'svg',
	attrs: { xmlns: 'http://www.w3.org/2000/svg', width: 10, height: 10, viewBox: '0 0 10 10' },
	children: [child],
});
const path = (attrs = {}) => ({
	tag: 'path',
	attrs: { d: 'M0 0L1 1z', fill: '#112233', ...attrs },
});

test('sanitizer rejects scripts, links, images, styles, CSS URLs, text and event attributes', () => {
	for (const tag of ['script', 'foreignObject', 'image', 'use', 'a', 'text', 'style', 'animate'])
		assert.throws(() => sanitizeVectorTree(root({ tag, attrs: {} })));
	for (const attrs of [
		{ onclick: 'alert(1)' },
		{ href: 'https://example.invalid' },
		{ style: 'fill:red' },
		{ fill: 'url(#paint)' },
		{ stroke: 'url(https://example.invalid)' },
		{ d: 'M0 0LInfinity 0' },
		{ d: 'M0 0L1e300 0' },
		{ 'stroke-width': NaN },
	])
		assert.throws(() => sanitizeVectorTree(root(path(attrs))));
});

test('sanitizer enforces node, depth, output and cyclic-tree limits', () => {
	assert.throws(() => sanitizeVectorTree(root(path()), { ...LIMITS, outputNodes: 1 }));
	assert.throws(() => sanitizeVectorTree(root(path()), { ...LIMITS, outputChars: 1 }));
	assert.throws(() => sanitizeVectorTree(root(path()), { ...LIMITS, depth: 0 }));
	const cycle = root(path());
	cycle.children = [cycle];
	assert.throws(() => sanitizeVectorTree(cycle));
});

test('watchdog terminates a synchronous infinite loop', async () => {
	await assert.rejects(runIsolated(new Uint8Array(), 'watchdog-test', 100), /emf.timeout/);
});

test('sanitizer rejects oversized and sparse child arrays before allocating output', () => {
	for (const children of [new Array(2 ** 30), new Array(2), {}, null]) {
		const top = root(path());
		top.children = children;
		assert.throws(() => sanitizeVectorTree(top));
		assert.throws(() => sanitizeVectorTree(root({ tag: 'g', attrs: {}, children })));
	}
	assert.throws(() => sanitizeVectorTree(root({ ...path(), children: {} })));
});

// Characterization tests: these PASS when the installed package still has the defects.
test('3.5.1 characterization: MM_TEXT incorrectly applies extents', async () => {
	const normal = await runIsolated(cases['plain-rectangle']);
	const mapped = await runIsolated(cases['mm-text-extents']);
	assert.ok(normal.metrics.redBox.width > 290);
	assert.equal(mapped.metrics.redBox.width, 99);
	assert.equal(mapped.warningCount, 0);
});

test('3.5.1 characterization: odd driver glyphs vanish, even glyphs and matrix variants match', async () => {
	const control = await runIsolated(cases['draw-string-control']);
	const odd = await runIsolated(cases['driver-odd-count']);
	const padded = await runIsolated(cases['driver-odd-padded-data']);
	const even = await runIsolated(cases['driver-even-count']);
	const matrix = await runIsolated(cases['driver-even-matrix']);
	assert.ok(control.metrics.darkRight > 0);
	assert.equal(odd.metrics.darkRight, 0);
	assert.equal(padded.metrics.darkRight, 0);
	assert.ok(even.metrics.darkRight > 0);
	assert.deepEqual(even.textNodes, matrix.textNodes);
	assert.deepEqual(even.metrics, matrix.metrics);
	assert.equal(odd.warningCount, 0);
});

test('3.5.1 characterization: a leaf region is rejected and leaves the old clip', async () => {
	const replaced = await runIsolated(cases['clip-replace-leaf']);
	const control = await runIsolated(cases['clip-leaf-control']);
	const wrongCount = await runIsolated(cases['clip-replace-wrong-count']);
	assert.equal(replaced.metrics.redRight, 0);
	assert.ok(control.metrics.redRight > 40_000);
	assert.equal(wrongCount.metrics.redRight, control.metrics.redRight);
	assert.equal(replaced.warningCount, 0);
});

test('3.5.1 characterization: header signature and EOF are not enforced', async () => {
	const bytes = Buffer.from(cases['plain-rectangle']);
	bytes.writeUInt32LE(0, 40);
	const signature = await runIsolated(bytes);
	const eof = await runIsolated(cases['plain-rectangle'].subarray(0, -20));
	for (const result of [signature, eof]) {
		assert.equal(result.status, 'converted');
		assert.equal(result.warningCount, 0);
		assert.equal(result.scan.structurallyValid, false);
	}
});
