// Fifteenth round of starting-state probes for FORECAST.ETS (see ets-init.ts): every length from the season
// plus 1 to 2 seasons minus 1 for seasons of 2 to 9, to find where Excel starts using the season.
//   bun scripts/xlsx/ets-init15-probes.ts          (from src/core)
// then record Excel's answers with `xlsx/formula/__fixtures__/ets-init2-cases.ps1 -Set 15`.
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

export const INIT15_GROUPS: readonly (readonly [number, number])[] = [
	[2, 3],
	[3, 4],
	[3, 5],
	[4, 5],
	[4, 6],
	[4, 7],
	[5, 6],
	[5, 7],
	[5, 8],
	[5, 9],
	[6, 7],
	[6, 8],
	[6, 9],
	[6, 10],
	[6, 11],
	[7, 8],
	[7, 9],
	[7, 10],
	[7, 11],
	[7, 12],
	[7, 13],
	[8, 9],
	[8, 10],
	[8, 11],
	[8, 12],
	[8, 13],
	[8, 14],
	[8, 15],
	[9, 10],
	[9, 11],
	[9, 12],
	[9, 13],
	[9, 14],
	[9, 15],
	[9, 16],
	[9, 17],
];

export function buildInit15Cases(): EtsCaseInput[] {
	const out: EtsCaseInput[] = [];
	const r = rng(151515);
	const rand = (amp: number): number => (r() - 0.5) * 2 * amp;
	for (const [m, n] of INIT15_GROUPS) {
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

export const ETS_INIT15_SERIES_PATH = fileURLToPath(
	new URL('../../xlsx/formula/__fixtures__/ets-init15-series.json', import.meta.url),
);
if (import.meta.main) {
	writeFileSync(ETS_INIT15_SERIES_PATH, `${JSON.stringify(buildInit15Cases(), null, '\t')}\n`);
}
