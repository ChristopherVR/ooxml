import assert from 'node:assert/strict';
import { test } from 'node:test';
import { mergeReports, vitestTimings } from './test-timings.mjs';

const report = (success, ...files) => ({
	success,
	testResults: files.map(([name, ms]) => ({ name, startTime: 1000, endTime: 1000 + ms })),
});

test('shard reports merge into one table keyed by path inside the suite', () => {
	const merged = mergeReports([
		report(true, ['/home/runner/work/ooxml/ooxml/src/core/xml/a.test.ts', 1234]),
		report(true, ['/home/runner/work/ooxml/ooxml/src/core/docx/b.test.ts', 4]),
	]);
	assert.deepEqual(vitestTimings(merged, 'core'), {
		'docx/b.test.ts': 0.01,
		'xml/a.test.ts': 1.23,
	});
});

test('a failed shard keeps the table unchanged', () => {
	const merged = mergeReports([
		report(true, ['/w/src/core/a.test.ts', 10]),
		report(false, ['/w/src/core/b.test.ts', 10]),
	]);
	assert.throws(() => vitestTimings(merged, 'core'), /successful Vitest JSON report/);
});
