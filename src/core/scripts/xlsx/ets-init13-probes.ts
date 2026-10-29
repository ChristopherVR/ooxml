// Thirteenth round of starting-state probes for FORECAST.ETS (see ets-init.ts): 2 seasons plus 1 to m - 1 points
// for seasons of 4 to 8 (the lengths between the exact 2-season and 3-season rules).
//   bun scripts/xlsx/ets-init13-probes.ts          (from src/core)
// then record Excel's answers with `xlsx/formula/__fixtures__/ets-init2-cases.ps1 -Set 13`.
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

export const INIT13_GROUPS: readonly (readonly [number, number])[] = [
	[4, 9],
	[4, 10],
	[4, 11],
	[5, 11],
	[5, 12],
	[5, 13],
	[5, 14],
	[6, 13],
	[6, 14],
	[6, 15],
	[6, 16],
	[6, 17],
	[7, 15],
	[7, 16],
	[7, 17],
	[7, 18],
	[7, 19],
	[7, 20],
	[8, 17],
	[8, 18],
	[8, 19],
	[8, 20],
	[8, 21],
	[8, 22],
	[8, 23],
];

export function buildInit13Cases(): EtsCaseInput[] {
	const out: EtsCaseInput[] = [];
	const r = rng(131313);
	const rand = (amp: number): number => (r() - 0.5) * 2 * amp;
	for (const [m, n] of INIT13_GROUPS) {
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

export const ETS_INIT13_SERIES_PATH = fileURLToPath(
	new URL('../../xlsx/formula/__fixtures__/ets-init13-series.json', import.meta.url),
);
if (import.meta.main) {
	writeFileSync(ETS_INIT13_SERIES_PATH, `${JSON.stringify(buildInit13Cases(), null, '\t')}\n`);
}
