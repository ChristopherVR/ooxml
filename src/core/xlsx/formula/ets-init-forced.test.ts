// With Excel's own smoothing parameters (STAT 1 to 3 of every probe series), the exact seasonal start
// (functions/ets-init-chain.ts and its helpers) reproduces Excel's forecasts to rounding error on
// hundreds of probe series: 4 and more seasons with a partial one, 3 seasons plus a partial one, 2
// seasons plus a partial one (seasons of 4 and more; of 3 from 6 points and of 2 from 6), seasons 2 to
// 12. The only exceptions are the series on which Excel reports a parameter pair at its floor (0.001)
// that does not reproduce its own forecasts.
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { etsRun } from './functions/ets-core';
import { initialStates } from './functions/ets-init';
import { chainStart, isExactLength } from './functions/ets-init-chain';

interface Recorded {
	name: string;
	values: number[];
	results: { value: number }[];
}

const FILES = [
	'2',
	'3',
	'4',
	'5',
	'6',
	'7',
	'8',
	'9',
	'10',
	'11',
	'12',
	'13',
	'14',
	'15',
	'16',
	'18',
];
const FLOORED = new Set([
	's2-10-imp9',
	's3-18-imp11',
	's3-18-imp12',
	's4-13-imp11',
	's4-15-imp6',
	's8-29-imp8',
	's2-9-imp0',
	's3-9-imp0',
	's3-9-imp4',
	's3-9-imp8',
	's3-10-imp8',
	's7-17-imp6',
]);

// Probe groups whose parameter pair Excel reports at the beta floor (0.001) does not reproduce its own
// forecasts, while the same series at other parameters does (file:group).
const FLOORED_GROUPS = ['10:s7-33', '11:s6-23', '11:s12-47'];

function worstError(e: Recorded, m: number): number {
	const y = e.values;
	const n = y.length;
	const v = e.results.map((r) => r.value);
	const p = { alpha: v[0] as number, beta: v[1] as number, gamma: v[2] as number };
	const r = etsRun(y, m, initialStates(y, m, p, etsRun), p);
	let err = 0;
	for (let i = 0; i < v.length - 3; i++) {
		const h = i + 1;
		const f = r.level + h * r.trend + (r.season[(n - 1 + h) % m] as number);
		const ex = v[3 + i] as number;
		err = Math.max(err, Math.abs(f - ex) / Math.abs(ex));
	}
	return err;
}

describe('exact seasonal start at Excel parameters', () => {
	let total = 0;
	for (const f of FILES) {
		const path = fileURLToPath(new URL(`./__fixtures__/ets-init${f}-excel.json`, import.meta.url));
		const data = JSON.parse(readFileSync(path, 'utf8')) as Recorded[];
		it(`reproduces the forecasts of ets-init${f}-excel.json`, () => {
			let checked = 0;
			for (const e of data) {
				const m = /^s(\d+)-/.exec(e.name);
				if (!m || FLOORED.has(e.name) || !chainStart(e.values, Number(m[1]))) continue;
				const group = `${f}:${e.name.replace(/-(base|imp[0-9]+)$/, '')}`;
				if (FLOORED_GROUPS.includes(group) && (e.results[1] as { value: number }).value < 0.0011)
					continue;
				if (!isExactLength(e.values.length, Number(m[1]))) continue;
				expect(worstError(e, Number(m[1]))).toBeLessThan(1e-9);
				checked++;
			}
			total += checked;
		});
	}

	it('checked hundreds of exact-start series', () => {
		expect(total).toBeGreaterThan(300);
	});
});
