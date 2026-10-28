import assert from 'node:assert/strict';
import { existsSync, mkdirSync, mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { test } from 'node:test';
import {
	clearStamp,
	defaultJobs,
	hashInputs,
	runSteps,
	stampMatches,
	stampPath,
	writeStamp,
} from './verify-runner.mjs';

const delay = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

/** An exec that records start order and the peak number of concurrent steps. */
function fake(codes = {}, ms = 10) {
	const log = { started: [], finished: [], peak: 0, active: 0 };
	const exec = async (step) => {
		log.started.push(step.id);
		log.peak = Math.max(log.peak, ++log.active);
		await delay(ms);
		log.active--;
		log.finished.push(step.id);
		return { code: codes[step.id] ?? 0, output: `out ${step.id}` };
	};
	return { exec, log };
}

test('a step starts only after the steps it needs finished', async () => {
	const { exec, log } = fake();
	const steps = [{ id: 'c', needs: ['b'] }, { id: 'b', needs: ['a'] }, { id: 'a' }, { id: 'free' }];
	const { failed } = await runSteps(steps, { jobs: 4, exec });
	assert.equal(failed, false);
	assert.ok(log.started.indexOf('a') < log.started.indexOf('b'));
	assert.ok(log.finished.indexOf('a') < log.started.indexOf('b'));
	assert.ok(log.finished.indexOf('b') < log.started.indexOf('c'));
	assert.ok(log.started.includes('free'));
});

test('no more than `jobs` steps run at once', async () => {
	const { exec, log } = fake();
	const steps = Array.from({ length: 8 }, (_, i) => ({ id: `s${i}` }));
	await runSteps(steps, { jobs: 3, exec });
	assert.equal(log.peak, 3);
	const serial = fake();
	await runSteps(steps, { jobs: 1, exec: serial.exec });
	assert.equal(serial.log.peak, 1);
});

test('a failure stops new steps, lets running ones finish and blocks dependants', async () => {
	const { exec, log } = fake({ bad: 1 });
	const steps = [
		{ id: 'bad' },
		{ id: 'other' },
		{ id: 'child', needs: ['bad'] },
		{ id: 'later', needs: ['other'] },
	];
	const { results, failed } = await runSteps(steps, { jobs: 2, exec });
	assert.equal(failed, true);
	const status = Object.fromEntries(results.map((r) => [r.step.id, r.status]));
	assert.equal(status.bad, 'failed');
	assert.equal(status.other, 'ok');
	assert.equal(status.child, 'blocked');
	assert.equal(status.later, 'blocked');
	assert.ok(!log.started.includes('child') && !log.started.includes('later'));
	assert.equal(results.find((r) => r.step.id === 'bad').output, 'out bad');
});

test('skipIf skips without running and still satisfies dependants; onSuccess runs after ok', async () => {
	const { exec, log } = fake();
	let stamped = false;
	const steps = [
		{ id: 'build', skipIf: () => 'inputs unchanged', onSuccess: () => (stamped = true) },
		{ id: 'use', needs: ['build'] },
	];
	const { results } = await runSteps(steps, { jobs: 2, exec });
	assert.deepEqual(log.started, ['use']);
	assert.equal(results.find((r) => r.step.id === 'build').status, 'skipped');
	assert.equal(stamped, false);
	await runSteps([{ id: 'x', onSuccess: () => (stamped = true) }], { jobs: 1, exec });
	assert.equal(stamped, true);
});

test('an unknown need is rejected', async () => {
	await assert.rejects(
		runSteps([{ id: 'a', needs: ['nope'] }], { exec: fake().exec }),
		/unknown step/,
	);
});

test('an exclusive browser step drains running checks and gets the whole machine', async () => {
	const active = new Set();
	const order = [];
	await runSteps([{ id: 'unit' }, { id: 'browser', exclusive: true }, { id: 'types' }], {
		jobs: 3,
		exec: async (step) => {
			if (step.exclusive) assert.equal(active.size, 0);
			else assert.ok(!active.has('browser'));
			active.add(step.id);
			order.push(step.id);
			await delay(5);
			active.delete(step.id);
			return { code: 0, output: '' };
		},
	});
	assert.deepEqual(order, ['unit', 'browser', 'types']);
});

test('defaultJobs is a quarter of the cores, at least one', () => {
	assert.equal(defaultJobs(32), 8);
	assert.equal(defaultJobs(2), 1);
});

test('input hash follows tracked sources and ignores tests and docs', () => {
	const listing = (extra = '') =>
		`100644 aaa 0\tsrc/core/a.ts\n100644 bbb 0\tsrc/core/a.test.ts${extra}\n100644 ccc 0\tsrc/core/docs/x.md\n`;
	const hash = (text, extra = []) =>
		hashInputs({ dirs: ['src/core'], extra, lsFiles: () => text, root: tmpdir() });
	const base = hash(listing());
	assert.equal(hash(listing().replace('bbb', 'zzz')), base);
	assert.notEqual(hash(listing().replace('aaa', 'zzz')), base);
	assert.notEqual(hash(listing('\n100644 ddd 0\tsrc/core/new.ts')), base);
	assert.notEqual(hash(listing(), ['core-stamp']), base);
});

test('stamps match only for the same hash and an existing dist', () => {
	const root = mkdtempSync(join(tmpdir(), 'verify-stamp-'));
	try {
		const dist = join(root, 'dist');
		assert.equal(stampMatches(dist, 'h1'), false);
		// The stamp never lands in `dist`, which is published.
		assert.ok(!stampPath(dist).startsWith(dist));
		writeStamp(dist, 'h1');
		assert.equal(stampMatches(dist, 'h1'), false, 'no dist yet');
		mkdirSync(dist);
		assert.equal(stampMatches(dist, 'h1'), true);
		assert.equal(stampMatches(dist, 'h2'), false);
		clearStamp(dist);
		assert.equal(existsSync(stampPath(dist)), false);
		assert.equal(stampMatches(dist, 'h1'), false);
	} finally {
		rmSync(root, { recursive: true, force: true });
	}
});
