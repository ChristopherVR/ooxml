import assert from 'node:assert/strict';
import { test } from 'node:test';
import { balanceFiles, fileDurations } from './browser-shards.mjs';

test('duration balancing retains every file exactly once including new tests', () => {
	const files = ['slow', 'medium', 'fast', 'new'];
	const result = balanceFiles(files, { slow: 100, medium: 60, fast: 20 }, 2);
	assert.deepEqual(result.flatMap((bin) => bin.files).sort(), [...files].sort());
	assert.deepEqual(
		result.map((bin) => bin.duration),
		[120, 120],
	);
	assert.deepEqual(
		result,
		balanceFiles([...files].reverse(), { slow: 100, medium: 60, fast: 20 }, 2),
	);
});

test('empty, corrupt or stale history cannot suppress tests', () => {
	assert.deepEqual(fileDurations({ version: 8 }, 'react'), {});
	assert.deepEqual(fileDurations({ version: 1, tests: { bad: 1 } }, 'react'), {});
	assert.equal(
		balanceFiles(['a', 'b'], { a: -1, b: NaN, removed: 100 }, 3).flatMap((b) => b.files).length,
		2,
	);
	assert.throws(() => balanceFiles([], {}, 0));
});

test('history remains specific to the framework and sums test durations by file', () => {
	const key = (project, title) => JSON.stringify([project, 'a.spec.ts', [title]]);
	assert.deepEqual(
		fileDurations(
			{
				version: 1,
				tests: {
					[key('react', 'one')]: 10,
					[key('react', 'two')]: 20,
					[key('vue', 'one')]: 90,
				},
			},
			'react',
		),
		{ 'a.spec.ts': 30 },
	);
});
