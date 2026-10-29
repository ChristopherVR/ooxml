// Writes ets-detect-sweep-series.json (bun scripts/xlsx/ets-detect-sweep.ts from src/core): the white
// noise series of ets-detect-excel.json for which Excel reports no season (the first two of each
// length from 20 to 60), each with a cosine at a Fourier frequency added in ten growing amplitudes
// (0.15 to 3 standard deviations of the series). Two frequencies per series: the strongest allowed
// periodogram bin and the middle one. Record Excel's answers with ets-search-cases.ps1 and save them
// as ets-detect-sweep-excel.json; forecast-ets-detect.test.ts reads both.
import { readFileSync, writeFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import type { EtsCaseInput } from './ets-series';

interface Recorded {
	name: string;
	values: number[];
	results: { value: number }[];
}

const AMPLITUDES = [0.15, 0.3, 0.45, 0.6, 0.8, 1, 1.3, 1.7, 2.2, 3];
const period = (n: number, k: number): number => Math.ceil(n / k - 0.5);

function periodogram(y: readonly number[]): number[] {
	const n = y.length;
	const mean = y.reduce((a, b) => a + b, 0) / n;
	const mx = (n - 1) / 2;
	let sxy = 0;
	let sxx = 0;
	y.forEach((v, i) => {
		sxy += (i - mx) * (v - mean);
		sxx += (i - mx) * (i - mx);
	});
	const z = y.map((v, i) => v - mean - (sxy / sxx) * (i - mx));
	return Array.from({ length: Math.floor(n / 2) + 1 }, (_v, k) => {
		let c = 0;
		let s = 0;
		z.forEach((v, i) => {
			c += v * Math.cos((2 * Math.PI * i * k) / n);
			s += v * Math.sin((2 * Math.PI * i * k) / n);
		});
		return c * c + s * s;
	});
}

export function sweepSeries(recorded: readonly Recorded[]): EtsCaseInput[] {
	const taken = new Map<number, number>();
	const out: EtsCaseInput[] = [];
	for (const r of recorded) {
		const n = r.values.length;
		if (!r.name.startsWith('noise-') || r.results[0]?.value !== 0 || n < 20 || n > 60) continue;
		const seen = (taken.get(n) ?? 0) + 1;
		taken.set(n, seen);
		if (seen > 2) continue;
		const y = r.values;
		const mean = y.reduce((a, b) => a + b, 0) / n;
		const sd = Math.sqrt(y.reduce((a, b) => a + (b - mean) ** 2, 0) / n);
		const power = periodogram(y);
		const ks = Array.from({ length: 12 }, (_v, i) => i + 1).filter(
			(k) => k <= n / 2 && period(n, k) >= 2 && n / period(n, k) >= 2.5,
		);
		const strongest = ks.reduce((a, k) => ((power[k] ?? 0) > (power[a] ?? 0) ? k : a), ks[0] ?? 1);
		const middle = ks[Math.floor(ks.length / 2)] ?? 1;
		for (const k of new Set([strongest, middle]))
			for (const a of AMPLITUDES) {
				const values = y.map(
					(v, i) => Math.round((v + a * sd * Math.cos((2 * Math.PI * k * i) / n)) * 100) / 100,
				);
				out.push({
					name: `${r.name}|k${k}|a${a}`,
					values,
					timeline: values.map((_v, i) => i + 1),
					formulas: ['FORECAST.ETS.SEASONALITY({V},{L})'],
				});
			}
	}
	return out;
}

if (process.argv[1] === fileURLToPath(import.meta.url)) {
	const dir = '../../xlsx/formula/__fixtures__/';
	const source = JSON.parse(
		readFileSync(fileURLToPath(new URL(`${dir}ets-detect-excel.json`, import.meta.url)), 'utf8'),
	) as Recorded[];
	const all = sweepSeries(source);
	const path = fileURLToPath(new URL(`${dir}ets-detect-sweep-series.json`, import.meta.url));
	writeFileSync(path, '[\n' + all.map((s) => '\t' + JSON.stringify(s)).join(',\n') + '\n]\n');
	console.log(all.length, 'series');
}
