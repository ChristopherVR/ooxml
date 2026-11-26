import assert from 'node:assert/strict';
import { mkdtempSync, readFileSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { test } from 'node:test';
import { SUITE_ROUTES } from './perf-budget-routes.mjs';
import {
	assertNoCheckoutPath,
	limitFor,
	problems,
	readLimits,
	writeLimits,
} from './perf-budgets.mjs';

test('a limit is 3% over the measurement, at least 1 KiB, in whole KiB', () => {
	assert.equal(limitFor(100_000), 101 * 1024);
	assert.equal(limitFor(1_000_000), 1006 * 1024);
	assert.equal(limitFor(10), 2048);
	assert.equal(limitFor(1024), 2048);
});

test('a budget fails when broken, stale, missing or unmeasured', () => {
	const ok = { 'suite/route/home': 50_000 };
	assert.deepEqual(problems(ok, { 'suite/route/home': limitFor(50_000) }), []);
	const [over] = problems({ 'suite/route/home': 60_000 }, { 'suite/route/home': limitFor(50_000) });
	assert.match(over, /is over its budget .* by/);
	const [stale] = problems(
		{ 'suite/route/home': 40_000 },
		{ 'suite/route/home': limitFor(50_000) },
	);
	assert.match(stale, /far under its budget/);
	assert.match(problems(ok, {})[0], /has no budget/);
	assert.match(problems({}, { gone: 1 })[0], /was not measured/);
});

test('a work budget measures exact counts without byte rounding', () => {
	const name = 'work/pptx/large deck missed first-slide previews';
	assert.deepEqual(problems({ [name]: 0 }, { [name]: 0 }), []);
	assert.match(problems({ [name]: 1 }, { [name]: 0 })[0], /1 is over its budget of 0 by 1/);
});

test('update writes sorted limits that check then accepts', () => {
	const dir = mkdtempSync(join(tmpdir(), 'ooxml-perf-budgets-'));
	try {
		const file = join(dir, 'budgets.json');
		const sizes = { 'package/b': 5000, 'package/a': 300_000 };
		writeLimits(sizes, file);
		const limits = readLimits(file);
		assert.deepEqual(Object.keys(limits), ['package/a', 'package/b']);
		assert.deepEqual(problems(sizes, limits), []);
		writeLimits(sizes, file, { 'package/a': 300_000 });
		assert.equal(readLimits(file)['package/a'], 300_000);
	} finally {
		rmSync(dir, { recursive: true, force: true });
	}
});

test('output that names the checkout it was built in is refused', () => {
	const root = 'D:\\work\\ooxml';
	assert.doesNotThrow(() => assertNoCheckoutPath('a.js', '// raw-css:src/x.css?raw', root));
	assert.throws(
		() => assertNoCheckoutPath('a.js', '// raw-css:D:\\work\\ooxml\\src\\x.css?raw', root),
		/a\.js contains the checkout path/,
	);
	assert.throws(() => assertNoCheckoutPath('a.js', 'D:/work/ooxml/src', root), /checkout path/);
});

test('every suite route has a committed budget', () => {
	const limits = readLimits();
	for (const route of Object.keys(SUITE_ROUTES))
		assert.ok(limits[`suite/route/${route}`] > 0, route);
	assert.ok(limits['suite/all javascript'] > 0);
});

test('the default limits come entirely from the committed budget file', () => {
	const file = new URL('./perf-budgets.json', import.meta.url);
	assert.deepEqual(readLimits(), JSON.parse(readFileSync(file, 'utf8')).budgets);
});
