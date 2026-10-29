// Sixteenth round of starting-state probes for FORECAST.ETS (see ets-init.ts): impulse groups for every
// length where Excel uses the season below 2 seasons (m + 2 + ceil(m / 3) <= n < 2m), seasons of 8 to 13,
// plus 2 to 3 (the first lengths with a season), to recover the starting-state maps there.
//   bun scripts/xlsx/ets-init16-probes.ts          (from src/core)
// then record Excel's answers with `xlsx/formula/__fixtures__/ets-init2-cases.ps1 -Set 16`.
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
	const horizons = 2 * m + 1;
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

export function init16Groups(): (readonly [number, number])[] {
	const out: (readonly [number, number])[] = [
		[2, 4],
		[2, 5],
		[3, 5],
	];
	for (let m = 8; m <= 13; m++)
		for (let n = m + 2 + Math.ceil(m / 3); n < 2 * m; n++) out.push([m, n]);
	return out;
}

export function buildInit16Cases(): EtsCaseInput[] {
	const out: EtsCaseInput[] = [];
	const r = rng(161616);
	const rand = (amp: number): number => (r() - 0.5) * 2 * amp;
	for (const [m, n] of init16Groups()) {
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

export const ETS_INIT16_SERIES_PATH = fileURLToPath(
	new URL('../../xlsx/formula/__fixtures__/ets-init16-series.json', import.meta.url),
);
if (import.meta.main) {
	writeFileSync(ETS_INIT16_SERIES_PATH, `${JSON.stringify(buildInit16Cases(), null, '\t')}\n`);
}
