// Seventeenth round of FORECAST.ETS probes (see ets-init.ts): where does Excel start using the season?
// Every length from the season plus 1 to 2 seasons minus 1 for seasons of 2 to 16, three series each
// (a noisy sine, a sawtooth with small noise, pure noise around a trend). Gamma (STAT 3) is the machine
// epsilon when Excel drops the season, and FORECAST.ETS.SEASONALITY says what it detects.
//   bun scripts/xlsx/ets-init17-probes.ts          (from src/core)
// then record Excel's answers with `xlsx/formula/__fixtures__/ets-init2-cases.ps1 -Set 17`.
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
			`FORECAST.ETS.SEASONALITY({V},{L})`,
			...[1, 2, 3].map((h) => `FORECAST.ETS(${n + h},{V},{L},${m})`),
		],
	};
}

export function buildInit17Cases(): EtsCaseInput[] {
	const out: EtsCaseInput[] = [];
	const r = rng(171717);
	const rand = (amp: number): number => (r() - 0.5) * 2 * amp;
	for (let m = 2; m <= 16; m++)
		for (let n = m + 1; n < 2 * m; n++) {
			const kinds: number[][] = [
				Array.from({ length: n }, (_v, i) =>
					round(30 + 0.4 * i + 5 * Math.sin((2 * Math.PI * i) / m) + rand(2.5)),
				),
				Array.from({ length: n }, (_v, i) => round(20 + 0.2 * i + (10 * (i % m)) / m + rand(0.5))),
				Array.from({ length: n }, (_v, i) => round(50 + 0.8 * i + rand(6))),
			];
			kinds.forEach((v, k) => out.push(entry(`c${m}-${n}-${k}`, v, m)));
		}
	return out;
}

export const ETS_INIT17_SERIES_PATH = fileURLToPath(
	new URL('../../xlsx/formula/__fixtures__/ets-init17-series.json', import.meta.url),
);
if (import.meta.main) {
	writeFileSync(ETS_INIT17_SERIES_PATH, `${JSON.stringify(buildInit17Cases(), null, '\t')}\n`);
}
