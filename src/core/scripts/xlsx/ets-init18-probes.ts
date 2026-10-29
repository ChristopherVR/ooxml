// Eighteenth round of FORECAST.ETS probes (see ets-init-short.ts): checks the predictions of the rules
// found from ets-init16 and 17 on seasons they were not fitted on. Cutoff series (`c`) for seasons 17 to
// 24, 30 and 40 on both sides of ceil(10 (m + 1) / 7), and impulse groups (`s`) for every seasonal
// length below 2 seasons of the seasons 14 and 16.
//   bun scripts/xlsx/ets-init18-probes.ts          (from src/core)
// then record Excel's answers with `xlsx/formula/__fixtures__/ets-init2-cases.ps1 -Set 18`.
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
	return {
		name,
		values,
		timeline: values.map((_v, i) => i + 1),
		formulas: [
			...[1, 2, 3].map((k) => `FORECAST.ETS.STAT({V},{L},${k},${m})`),
			...Array.from({ length: 2 * m + 1 }, (_v, h) => `FORECAST.ETS(${n + h + 1},{V},{L},${m})`),
		],
	};
}

export function buildInit18Cases(): EtsCaseInput[] {
	const out: EtsCaseInput[] = [];
	const r = rng(181818);
	const rand = (amp: number): number => (r() - 0.5) * 2 * amp;
	const series = (n: number, m: number): number[] =>
		Array.from({ length: n }, (_v, i) =>
			round(30 + 0.4 * i + 5 * Math.sin((2 * Math.PI * i) / m) + rand(2.5)),
		);
	for (const m of [17, 18, 19, 20, 21, 22, 23, 24, 30, 40]) {
		const first = Math.ceil((10 * (m + 1)) / 7);
		for (const n of [first - 2, first - 1, first, first + 1])
			out.push(entry(`c${m}-${n}-0`, series(n, m), m));
	}
	for (const m of [14, 16])
		for (let n = Math.ceil((10 * (m + 1)) / 7); n < 2 * m; n++) {
			const base = series(n, m);
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

export const ETS_INIT18_SERIES_PATH = fileURLToPath(
	new URL('../../xlsx/formula/__fixtures__/ets-init18-series.json', import.meta.url),
);
if (import.meta.main) {
	writeFileSync(ETS_INIT18_SERIES_PATH, `${JSON.stringify(buildInit18Cases(), null, '\t')}\n`);
}
