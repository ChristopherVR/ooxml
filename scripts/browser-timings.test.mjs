import assert from 'node:assert/strict';
import { mkdtempSync, mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { execFileSync } from 'node:child_process';
import { test } from 'node:test';
import BrowserTimingReporter from './browser-timing-reporter.mjs';

test('timings keep successful retries and distinguish parameterized tests and frameworks', () => {
	const reporter = new BrowserTimingReporter();
	reporter.onBegin({ rootDir: resolve('e2e/pptx') });
	const sample = (project, title) => ({
		expectedStatus: 'passed',
		parent: { project: () => ({ name: project }) },
		location: { file: resolve('e2e/pptx/chart.spec.ts') },
		titlePath: () => ['', project, 'chart.spec.ts', 'charts', title],
	});
	reporter.onTestEnd(sample('react', 'one'), { status: 'failed', duration: 60_000 });
	reporter.onTestEnd(sample('react', 'one'), { status: 'passed', duration: 200 });
	reporter.onTestEnd(sample('react', 'two'), { status: 'passed', duration: 300 });
	reporter.onTestEnd(sample('vue', 'one'), { status: 'passed', duration: 400 });
	reporter.onTestEnd(sample('vue', 'skip'), { status: 'skipped', duration: 0 });
	assert.deepEqual(Object.values(reporter.tests), [200, 300, 400]);
	assert.deepEqual(JSON.parse(Object.keys(reporter.tests)[0]), [
		'react',
		'chart.spec.ts',
		['chart.spec.ts', 'charts', 'one'],
	]);
});

test('merging a partial run updates observed tests without discarding earlier timings', () => {
	const dir = mkdtempSync(join(tmpdir(), 'ooxml-browser-timings-'));
	try {
		const input = join(dir, 'reports');
		mkdirSync(input);
		writeFileSync(
			join(input, '000-baseline.json'),
			JSON.stringify({ version: 1, tests: { old: 10, updated: 20 } }),
		);
		writeFileSync(
			join(input, 'latest.json'),
			JSON.stringify({ version: 1, tests: { updated: 30, added: 40 } }),
		);
		writeFileSync(
			join(input, 'other.json'),
			JSON.stringify({ version: 2, tests: { updated: 99 } }),
		);
		const output = join(dir, 'history.json');
		execFileSync(process.execPath, ['scripts/merge-browser-timings.mjs', input, output]);
		assert.deepEqual(JSON.parse(readFileSync(output, 'utf8')), {
			version: 1,
			tests: { old: 10, updated: 30, added: 40 },
		});
	} finally {
		rmSync(dir, { recursive: true, force: true });
	}
});
