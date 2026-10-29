// Eighth round of starting-state probes for FORECAST.ETS (see ets-init.ts): 3 whole seasons plus 1 to m - 2
// points for seasons of 6 to 10, which fixed the 3-season rule (functions/ets-init-three.ts).
//   bun scripts/xlsx/ets-init8-probes.ts          (from src/core)
// then record Excel's answers with `xlsx/formula/__fixtures__/ets-init2-cases.ps1 -Set 8`.
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

export const INIT8_GROUPS: readonly (readonly [number, number])[] = [
	[6, 20],
	[6, 21],
	[7, 22],
	[7, 23],
	[7, 24],
	[8, 25],
	[8, 26],
	[8, 27],
	[9, 28],
	[9, 29],
	[10, 30],
	[10, 31],
	[10, 32],
	[6, 18],
	[4, 16],
	[5, 20],
];

export function buildInit8Cases(): EtsCaseInput[] {
	const out: EtsCaseInput[] = [];
	const r = rng(80818);
	const rand = (amp: number): number => (r() - 0.5) * 2 * amp;
	for (const [m, n] of INIT8_GROUPS) {
		const base = Array.from({ length: n }, (_v, i) =>
			round(30 + 0.4 * i + 5 * Math.sin((2 * Math.PI * i) / m) + rand(0.4)),
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

export const ETS_INIT8_SERIES_PATH = fileURLToPath(
	new URL('../../xlsx/formula/__fixtures__/ets-init8-series.json', import.meta.url),
);
if (import.meta.main) {
	writeFileSync(ETS_INIT8_SERIES_PATH, `${JSON.stringify(buildInit8Cases(), null, '\t')}\n`);
}
