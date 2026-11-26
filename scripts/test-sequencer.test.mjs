import assert from 'node:assert/strict';
import { mkdtempSync, rmdirSync, unlinkSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { resolve } from 'node:path';
import test from 'node:test';
import { pathToFileURL } from 'node:url';
import { balanceTestFiles, durationSequencer } from './test-sequencer.mjs';
import { vitestTimings } from './test-timings.mjs';
import { testWorkers } from './test-workers.mjs';

test('longest files go to the lightest shard with stable ties and a weight for new files', () => {
	const files = ['new/z', 'light', 'medium', 'new/a', 'heavy'];
	const timings = { light: 1, medium: 6, heavy: 10, stale: 90 };
	const expected = [
		{ files: ['heavy', 'new/z'], duration: 15 },
		{ files: ['medium', 'new/a', 'light'], duration: 12 },
	];
	assert.deepEqual(balanceTestFiles(files, timings, 2), expected);
	assert.deepEqual(balanceTestFiles(files.toReversed(), timings, 2), expected);
	assert.deepEqual(balanceTestFiles(['new/a', 'new/z'], {}, 3), [
		{ files: ['new/a'], duration: 5 },
		{ files: ['new/z'], duration: 5 },
		{ files: [], duration: 0 },
	]);
});

test('the sequencer partitions Vitest specifications, including a file in multiple projects', async (t) => {
	class BaseSequencer {
		constructor(ctx) {
			this.ctx = ctx;
		}
	}
	const root = mkdtempSync(resolve(tmpdir(), 'vitest-timings-'));
	const timingFile = pathToFileURL(resolve(root, 'vitest.timings.json'));
	writeFileSync(timingFile, JSON.stringify({ 'subdir/a.test.ts': 9 }));
	t.after(() => {
		unlinkSync(timingFile);
		rmdirSync(root);
	});
	const a = { moduleId: resolve(root, 'subdir/a.test.ts'), project: 'one' };
	const b = { moduleId: resolve(root, 'subdir/b.test.ts'), project: 'one' };
	const aAgain = { ...a, project: 'two' };
	const Sequencer = durationSequencer(BaseSequencer, timingFile);
	const all = [b, aAgain, a];
	const shard = (index) => new Sequencer({ config: { shard: { index, count: 2 } } }).shard(all);
	assert.deepEqual(await shard(1), [aAgain, a]);
	assert.deepEqual(await shard(2), [b]);
});

test('timings use suite-relative paths and round the sum of duplicate file runs to centiseconds', () => {
	const report = {
		success: true,
		testResults: [
			{ name: 'D:\\checkout\\src\\core\\z.test.ts', startTime: 123, endTime: 128 },
			{ name: '/checkout/src/core/subdir/a.test.ts', startTime: 30, endTime: 250.9 },
			{ name: 'D:/checkout/src/core/subdir/a.test.ts', startTime: 250, endTime: 275.8 },
		],
	};
	assert.deepEqual(Object.entries(vitestTimings(report, 'core')), [
		['subdir/a.test.ts', 0.25],
		['z.test.ts', 0.01],
	]);
	assert.throws(() => vitestTimings(report, 'ui'), /Invalid ui timing/);
	assert.throws(() => vitestTimings({ ...report, success: false }, 'core'), /successful/);
});

test('worker defaults leave local capacity and respect CI and an explicit override', () => {
	assert.equal(testWorkers({}, 1), 1);
	assert.equal(testWorkers({}, 7), 3);
	assert.equal(testWorkers({}, 32), 8);
	assert.equal(testWorkers({ CI: 'true' }, 2), 2);
	assert.equal(testWorkers({ CI: 'true' }, 4), 4);
	assert.equal(testWorkers({ CI: '1' }, 16), 4);
	assert.equal(testWorkers({ CI: '1', OOXML_TEST_WORKERS: '3' }, 16), 3);
	assert.throws(() => testWorkers({ OOXML_TEST_WORKERS: '2.5' }, 16), /positive integer/);
});
