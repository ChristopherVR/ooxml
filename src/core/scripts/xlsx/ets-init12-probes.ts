// Twelfth round of starting-state probes for FORECAST.ETS (see ets-init.ts): 3 whole seasons plus 3 to
// ceil(m / 2) + 3 points for seasons of 8 to 12, to find where the 3-season rule ends.
//   bun scripts/xlsx/ets-init12-probes.ts          (from src/core)
// then record Excel's answers with `xlsx/formula/__fixtures__/ets-init2-cases.ps1 -Set 12`.
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
const round = (v: number): number => Math.round(v * 1000) / 1000;

function entry(name: string, values: number[], m: number): EtsCaseInput {
	const n = values.length;
	const horizons = m >= 2 ? 2 * m + 1 : 3;
	return {
		name,
		values,
		timeline: values.map((_v, i) => i + 1),
		formulas: [
			...[1, 2, 3].map((k) => `FORECAST.ETS.STAT({V},{L},${k},${m})`),
			...Array.from({ length: horizons }, (_v, h) => `FORECAST.ETS(${n + h + 1},{V},{L},${m})`),
		],
	};
}

export const INIT12_GROUPS: readonly (readonly [number, number])[] = [
	[8, 27],
	[8, 28],
	[8, 29],
	[9, 31],
	[9, 32],
	[9, 33],
	[10, 34],
	[10, 35],
	[10, 36],
	[11, 37],
	[11, 38],
	[11, 39],
	[12, 40],
	[12, 41],
	[12, 42],
	[5, 19],
];

export function buildInit12Cases(): EtsCaseInput[] {
	const out: EtsCaseInput[] = [];
	const r = rng(121212);
	const rand = (amp: number): number => (r() - 0.5) * 2 * amp;
	for (const [m, n] of INIT12_GROUPS) {
		const base = Array.from({ length: n }, (_v, i) =>
			round(30 + 0.4 * i + 5 * Math.sin((2 * Math.PI * i) / m) + rand(2.5)),
		);
		out.push(entry(`s${m}-${n}-base`, base, m));
		for (let i = 0; i < n; i++)
			out.push(
				entry(
					`s${m}-${n}-imp${i}`,
					base.map((v, j) => (j === i ? round(v + 1) : v)),
					m,
				),
			);
	}
	return out;
}

export const ETS_INIT12_SERIES_PATH = fileURLToPath(
	new URL('../../xlsx/formula/__fixtures__/ets-init12-series.json', import.meta.url),
);
if (import.meta.main) {
	writeFileSync(ETS_INIT12_SERIES_PATH, `${JSON.stringify(buildInit12Cases(), null, '\t')}\n`);
}
