import assert from 'node:assert/strict';
import { test } from 'node:test';
import {
	encodeDrivePath,
	entryId,
	fmtBytes,
	fmtMs,
	interpret,
	median,
	parseArgs,
	percentOver,
	renderTable,
	stats,
	throughput,
} from './onedrive-speed-lib.mjs';
import { startMockGraph } from './onedrive-speed-mock.mjs';

test('median and stats ignore non-finite values', () => {
	assert.equal(median([3, 1, 2]), 2);
	assert.equal(median([4, 1, 2, 3]), 2.5);
	assert.deepEqual(stats([5, NaN, 1, 3]), { n: 3, median: 3, min: 1, max: 5, mean: 3 });
	assert.equal(stats([]).n, 0);
	assert.ok(Number.isNaN(stats([NaN]).median));
});

test('formatting and arithmetic', () => {
	assert.equal(fmtMs(12.34), '12.3 ms');
	assert.equal(fmtMs(2500), '2.50 s');
	assert.equal(fmtMs(NaN), '-');
	assert.equal(fmtBytes(1048576 * 2), '2.00 MB');
	assert.equal(fmtBytes(2048), '2.0 KB');
	assert.equal(throughput(10 * 1048576, 2000), 5);
	assert.equal(percentOver(150, 100), 50);
	assert.ok(Number.isNaN(percentOver(1, 0)));
});

test('drive paths are encoded per segment', () => {
	assert.equal(encodeDrivePath('Documents/My deck #1.pptx'), 'Documents/My%20deck%20%231.pptx');
	assert.equal(encodeDrivePath('/a//b/'), 'a/b');
	assert.throws(() => encodeDrivePath('/'));
	assert.equal(entryId({ driveId: 'b!x', itemId: 'I1' }), 'b%21x!I1');
});

test('renderTable pads columns', () => {
	const text = renderTable([
		['a', 'bb'],
		['long', '1'],
	]);
	assert.deepEqual(text.split('\n'), ['a     bb', '----  --', 'long   1']);
});

test('arguments are validated', () => {
	assert.equal(parseArgs([]).runs, 5);
	assert.deepEqual(
		{ ...parseArgs(['--runs=3', '--json', '--mock', '--size=2', '--latency=0']) },
		{ ...parseArgs([]), runs: 3, json: true, mock: true, size: 2, latency: 0 },
	);
	for (const bad of [
		'--runs=0',
		'--runs=1.5',
		'--runs=',
		'--size=-1',
		'--bogus',
		'--json=1',
		'--runs=500',
	])
		assert.throws(() => parseArgs([bad]), /Invalid value|--help/u);
});

test('interpretation names the dominant cost and library overhead', () => {
	const text = interpret({
		rawMeta: 100,
		rawDownload: 900,
		rawTtfb: 50,
		library: 1010,
		provider: 1020,
		content: 1200,
		bytes: 10 * 1048576,
	}).join('\n');
	assert.match(text, /dominated by the download host/u);
	assert.match(text, /not the bottleneck/u);
	assert.match(text, /faster|as fast/u);
});

test('the mock serves metadata, a 302 and sized, throttled downloads', async () => {
	const mock = await startMockGraph({ size: 0.25, latency: 0, bandwidth: 0 });
	try {
		const meta = await mock.fetch('https://graph.microsoft.com/v1.0/shares/u!x/driveItem');
		const item = await meta.json();
		assert.equal(item.size, 262144);
		const redirect = await mock.fetch(
			'https://graph.microsoft.com/v1.0/shares/u!x/driveItem/content',
			{ redirect: 'manual' },
		);
		assert.equal(redirect.status, 302);
		const body = await mock.fetch(item['@microsoft.graph.downloadUrl']);
		assert.equal((await body.arrayBuffer()).byteLength, 262144);
		await assert.rejects(async () => mock.fetch('https://example.com/'));
	} finally {
		await mock.close();
	}
});
