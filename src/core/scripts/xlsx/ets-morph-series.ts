// Writes probe series that morph one vector into another, y = (1 - t) * u + t * v, so that the
// squared-error surface of the FORECAST.ETS filter moves smoothly with t while Excel's chosen
// smoothing parameters are recorded at each step. Used to find where Excel's choice changes and to
// compare those boundaries with candidate searches:
//   bun scripts/xlsx/ets-morph-series.ts          (from src/core)
// then record Excel's answers with `xlsx/formula/__fixtures__/ets-search-cases.ps1 -In ets-morph-series.json`.
import { writeFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import type { EtsCaseInput } from './ets-series';

function rng(seed: number): () => number {
	let a = seed >>> 0;
	return () => {
		a = (a + 0x6d2b79f5) >>> 0;
		let t = a;
		t = Math.imul(t ^ (t >>> 15), t | 1);
		t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
		return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
	};
}

const STEPS = 120;
const out: EtsCaseInput[] = [];
for (const [fam, n, seed] of [
	['a', 24, 11],
	['b', 40, 12],
	['c', 12, 13],
	['d', 30, 14],
] as const) {
	const r = rng(seed);
	const g = (): number => (r() + r() + r() + r() - 2) * 1.7;
	const u: number[] = [];
	const v: number[] = [];
	let x = 50;
	for (let i = 0; i < n; i++) {
		x += g();
		u.push(x);
		v.push(50 + 3 * g());
	}
	for (let k = 0; k <= STEPS; k++) {
		const t = k / STEPS;
		out.push({
			name: `morph-${fam}${n}-${k}`,
			values: u.map((a, i) => Math.round(((1 - t) * a + t * (v[i] as number)) * 1e6) / 1e6),
			timeline: Array.from({ length: n }, (_x, i) => i + 1),
			formulas: [
				'FORECAST.ETS.STAT({V},{L},1,0)',
				'FORECAST.ETS.STAT({V},{L},2,0)',
				'FORECAST.ETS.STAT({V},{L},7,0)',
				`FORECAST.ETS(${n + 1},{V},{L},0)`,
			],
		});
	}
}
writeFileSync(
	fileURLToPath(new URL('../../xlsx/formula/__fixtures__/ets-morph-series.json', import.meta.url)),
	'[\n' + out.map((r) => '\t' + JSON.stringify(r)).join(',\n') + '\n]\n',
);
console.log(out.length, 'series');
