// Writes the probe series for Excel's FORECAST.ETS.SEASONALITY detector (bun scripts/xlsx/ets-detect-series.ts
// from src/core): white noise at many lengths and seeds, and sinusoids with controlled noise.
// Record Excel's answers with ets-search-cases.ps1 (-In ets-detect-series.json).
import { writeFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { rng } from './ets-search-series';
import type { EtsCaseInput } from './ets-series';

const round = (v: number): number => Math.round(v * 1000) / 1000;

function entry(name: string, values: number[]): EtsCaseInput {
	return {
		name,
		values,
		timeline: values.map((_v, k) => k + 1),
		formulas: ['FORECAST.ETS.SEASONALITY({V},{L})'],
	};
}

export function detectSeries(): EtsCaseInput[] {
	const out: EtsCaseInput[] = [];
	const lengths = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 40, 48, 60];
	for (const n of lengths)
		for (let seed = 1; seed <= 16; seed++) {
			const r = rng(n * 1000 + seed);
			const g = (): number => (r() + r() + r() + r() - 2) * 1.7;
			out.push(
				entry(
					`noise-${n}-${seed}`,
					Array.from({ length: n }, () => round(50 + 3 * g())),
				),
			);
		}
	for (const period of [2, 3, 4, 5, 6, 7, 8, 12])
		for (const n of [24, 36, 50])
			for (const sd of [0.3, 0.7, 1.2, 2]) {
				const r = rng(period * 100 + n + sd * 7);
				const g = (): number => (r() + r() + r() + r() - 2) * 1.7;
				out.push(
					entry(
						`sin-p${period}-n${n}-s${sd}`,
						Array.from({ length: n }, (_x, i) =>
							round(50 + 3 * Math.sin((2 * Math.PI * i) / period) + sd * g()),
						),
					),
				);
			}
	return out;
}

if (process.argv[1] === fileURLToPath(import.meta.url)) {
	const path = fileURLToPath(
		new URL('../../xlsx/formula/__fixtures__/ets-detect-series.json', import.meta.url),
	);
	const all = detectSeries();
	writeFileSync(path, '[\n' + all.map((s) => '\t' + JSON.stringify(s)).join(',\n') + '\n]\n');
	console.log(all.length, 'series');
}
