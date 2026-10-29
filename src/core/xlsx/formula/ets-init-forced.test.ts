// With Excel's own smoothing parameters (STAT 1 to 3 of every probe series), the exact seasonal start
// (functions/ets-init-chain.ts) reproduces Excel's forecasts to rounding error: ~1200 probe series of
// 4 and more seasons, seasons 2 to 12. The only exceptions are the three series on which Excel
// reports a parameter pair at its floor (0.001) that does not reproduce its own forecasts.
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { etsRun } from './functions/ets-core';
import { initialStates } from './functions/ets-init';
import { chainStart } from './functions/ets-init-chain';

interface Recorded {
	name: string;
	values: number[];
	results: { value: number }[];
}

const FILES = ['2', '3', '4', '5'];
const FLOORED = new Set(['s2-10-imp9', 's3-18-imp11', 's3-18-imp12']);

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
	for (const f of FILES) {
		const path = fileURLToPath(new URL(`./__fixtures__/ets-init${f}-excel.json`, import.meta.url));
		const data = JSON.parse(readFileSync(path, 'utf8')) as Recorded[];
		it(`reproduces the forecasts of ets-init${f}-excel.json`, () => {
			let checked = 0;
			for (const e of data) {
				const m = /^s(\d+)-/.exec(e.name);
				if (!m || FLOORED.has(e.name) || !chainStart(e.values, Number(m[1]))) continue;
				expect(worstError(e, Number(m[1]))).toBeLessThan(1e-9);
				checked++;
			}
			expect(checked).toBeGreaterThan(30);
		});
	}
});
