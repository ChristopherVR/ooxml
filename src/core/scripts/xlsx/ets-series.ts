// Writes the input series for the FORECAST.ETS comparison with real Excel:
//   bun scripts/xlsx/ets-series.ts          (from src/core)
// then record Excel's answers with `xlsx/formula/__fixtures__/ets-cases.ps1`.
import { writeFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';

export interface EtsCaseInput {
	name: string;
	/** Values; null is an empty cell. */
	values: (number | null)[];
	timeline: number[];
	/** Formulas; {V} and {L} stand for the values and timeline ranges. */
	formulas: string[];
}

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
const seq = (n: number, from = 1, step = 1): number[] =>
	Array.from({ length: n }, (_v, i) => from + i * step);
const noise = (seed: number, amp: number, n: number): number[] => {
	const r = rng(seed);
	return Array.from({ length: n }, () => (r() - 0.5) * 2 * amp);
};
/** Excel serial of the first day of month `index` counted from January 2022. */
const monthStart = (index: number): number =>
	Math.round(Date.UTC(2022, index, 1) / 86400000) + 25569;

/** Forecast, interval, seasonality and statistics calls for a series of `n` points. */
function calls(n: number, from: number, step: number, extra: string[] = []): string[] {
	const t = (k: number): number => from + (n - 1 + k) * step;
	return [
		...[1, 2, 5].map((k) => `FORECAST.ETS(${t(k)},{V},{L})`),
		`FORECAST.ETS(${t(-2)},{V},{L})`,
		...[1, 3].map((k) => `FORECAST.ETS.CONFINT(${t(k)},{V},{L})`),
		`FORECAST.ETS.CONFINT(${t(2)},{V},{L},0.8)`,
		'FORECAST.ETS.SEASONALITY({V},{L})',
		...[1, 2, 3, 4, 5, 6, 7, 8].map((k) => `FORECAST.ETS.STAT({V},{L},${k})`),
		...extra,
	];
}

export function buildEtsCases(): EtsCaseInput[] {
	const out: EtsCaseInput[] = [];
	const add = (name: string, values: (number | null)[], timeline: number[], formulas: string[]) =>
		out.push({ name, values, timeline, formulas });
	const lin = (n: number, a: number, b: number, e: number[]) =>
		seq(n).map((x, i) => round(a + b * x + (e[i] ?? 0)));
	add('linear-exact', lin(12, 5, 2, []), seq(12), calls(12, 1, 1));
	add('linear-noisy', lin(20, 10, 1.5, noise(1, 3, 20)), seq(20), calls(20, 1, 1));
	add(
		'flat-noisy',
		noise(2, 5, 24).map((e) => round(50 + e)),
		seq(24),
		calls(24, 1, 1),
	);
	const s4 = [10, 20, 15, 30];
	const n4 = noise(3, 0.8, 24);
	add(
		'seasonal4-trend',
		seq(24).map((x, i) => round(s4[i % 4]! + 1.2 * x + n4[i]!)),
		seq(24),
		calls(24, 1, 1, [
			'FORECAST.ETS(25,{V},{L},4)',
			'FORECAST.ETS(25,{V},{L},0)',
			'FORECAST.ETS(25,{V},{L},8)',
			'FORECAST.ETS.CONFINT(26,{V},{L},0.95,4)',
			'FORECAST.ETS.STAT({V},{L},1,4)',
			'FORECAST.ETS.STAT({V},{L},7,0)',
		]),
	);
	const m12 = seq(12).map((k) => 20 * Math.sin((2 * Math.PI * k) / 12));
	const n12 = noise(4, 2, 36);
	add(
		'monthly-dates-seasonal12',
		seq(36).map((x, i) => round(100 + 0.8 * x + m12[i % 12]! + n12[i]!)),
		seq(36, 0).map(monthStart),
		[
			`FORECAST.ETS(${monthStart(36)},{V},{L})`,
			`FORECAST.ETS(${monthStart(40)},{V},{L})`,
			`FORECAST.ETS.CONFINT(${monthStart(37)},{V},{L})`,
			'FORECAST.ETS.SEASONALITY({V},{L})',
			...[1, 2, 3, 4, 5, 6, 7, 8].map((k) => `FORECAST.ETS.STAT({V},{L},${k})`),
		],
	);
	const w = [0, 5, 8, 9, 12, 30, 25];
	const n7 = noise(5, 2, 35);
	add(
		'seasonal7-daily',
		seq(35).map((x, i) => round(200 + w[i % 7]! * 3 + n7[i]!)),
		seq(35, 45000),
		calls(35, 45000, 1),
	);
	// Missing timeline points (gaps) with both completion modes.
	const gapT = [1, 2, 3, 5, 6, 8, 9, 10, 11, 13, 14, 15];
	const gn = noise(6, 1.2, 12);
	add(
		'gaps',
		gapT.map((x, i) => round(8 + 1.1 * x + gn[i]!)),
		gapT,
		[
			...calls(15, 1, 1),
			'FORECAST.ETS(16,{V},{L},1,0)',
			'FORECAST.ETS(16,{V},{L},1,1)',
			'FORECAST.ETS.CONFINT(16,{V},{L},0.95,1,0)',
			'FORECAST.ETS.STAT({V},{L},4,1,0)',
		],
	);
	// Duplicate timestamps, each aggregation.
	const dupT = [1, 1, 2, 3, 3, 3, 4, 5, 5, 6, 7, 7, 8, 9, 10, 10];
	const dn = noise(7, 4, 16);
	add(
		'duplicates',
		dupT.map((x, i) => round(30 + 2.5 * x + dn[i]!)),
		dupT,
		[1, 2, 3, 4, 5, 6, 7].flatMap((agg) => [
			`FORECAST.ETS(11,{V},{L},1,1,${agg})`,
			`FORECAST.ETS(13,{V},{L},0,1,${agg})`,
		]),
	);
	add(
		'growth',
		seq(18).map((x) => round(5 * Math.exp(0.15 * x))),
		seq(18),
		calls(18, 1, 1),
	);
	add('short-6', [4, 7, 5, 9, 8, 12], seq(6), [...calls(6, 1, 1), 'FORECAST.ETS(8,{V},{L},0)']);
	let walk = 100;
	add(
		'random-walk',
		noise(8, 6, 30).map((e) => round((walk += e))),
		seq(30),
		calls(30, 1, 1),
	);
	add('constant', Array(10).fill(7), seq(10), calls(10, 1, 1));
	const sn = noise(9, 1, 24);
	add(
		'step-change',
		seq(24).map((x, i) => round((x > 12 ? 40 : 20) + sn[i]!)),
		seq(24),
		calls(24, 1, 1),
	);
	const yn = noise(10, 15, 14);
	add(
		'yearly-step',
		seq(14).map((x, i) => round(1000 * 1.05 ** x + yn[i]!)),
		seq(14, 2010),
		calls(14, 2010, 1),
	);
	const v = lin(15, 3, 2.2, noise(11, 2, 15));
	const order = [5, 1, 10, 2, 15, 3, 8, 4, 6, 7, 9, 11, 14, 12, 13];
	add(
		'unsorted',
		order.map((t) => v[t - 1]!),
		order,
		calls(15, 1, 1),
	);
	const n3 = noise(12, 4, 30);
	const s3 = [0, 14, -9];
	add(
		'seasonal3-noisy',
		seq(30).map((x, i) => round(60 + s3[i % 3]! + n3[i]!)),
		seq(30),
		calls(30, 1, 1, ['FORECAST.ETS(31,{V},{L},3)', 'FORECAST.ETS(31,{V},{L},6)']),
	);
	add('blank-value', [3, 5, null, 9, 11, 13, 15, 17], seq(8), [
		'FORECAST.ETS(9,{V},{L})',
		'FORECAST.ETS(9,{V},{L},1,0)',
		'FORECAST.ETS.STAT({V},{L},1)',
	]);
	return out;
}

export const ETS_SERIES_PATH = fileURLToPath(
	new URL('../../xlsx/formula/__fixtures__/ets-series.json', import.meta.url),
);
if (import.meta.main)
	writeFileSync(ETS_SERIES_PATH, `${JSON.stringify(buildEtsCases(), null, '\t')}\n`);
