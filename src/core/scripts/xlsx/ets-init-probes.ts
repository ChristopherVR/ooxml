// Series probed against Excel to identify the FORECAST.ETS starting states (see ets-init.ts):
//   bun scripts/xlsx/ets-init-probes.ts          (from src/core)
// then record Excel's answers with `xlsx/formula/__fixtures__/ets-init-cases.ps1`. For each series
// Excel is asked for the smoothing parameters (STAT 1 to 3) and the forecasts at 2m + 1 horizons,
// which fix the final state of the filter, from which the starting state is recovered.
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
	const arg = `,${m}`;
	return {
		name,
		values,
		timeline: values.map((_v, i) => i + 1),
		formulas: [
			...[1, 2, 3].map((k) => `FORECAST.ETS.STAT({V},{L},${k}${arg})`),
			...Array.from({ length: horizons }, (_v, h) => `FORECAST.ETS(${n + h + 1},{V},{L}${arg})`),
		],
	};
}

export function buildInitCases(): EtsCaseInput[] {
	const out: EtsCaseInput[] = [];
	const r = rng(77);
	const rand = (amp: number): number => (r() - 0.5) * 2 * amp;
	// Seasonal: a base series with an impulse at every position, and random ones.
	for (const [m, n] of [
		[2, 4],
		[2, 6],
		[2, 8],
		[3, 6],
		[3, 9],
		[3, 12],
		[4, 8],
		[4, 12],
		[5, 10],
		[6, 12],
	] as const) {
		const base = Array.from({ length: n }, (_v, i) => round(20 + 0.5 * i + rand(6)));
		out.push(entry(`s${m}-${n}-base`, base, m));
		for (let i = 0; i < n; i++)
			out.push(
				entry(
					`s${m}-${n}-imp${i}`,
					base.map((v, j) => (j === i ? round(v + 5) : v)),
					m,
				),
			);
		for (let k = 0; k < 3; k++)
			out.push(
				entry(
					`s${m}-${n}-rnd${k}`,
					Array.from({ length: n }, () => round(20 + rand(8))),
					m,
				),
			);
	}
	// Short non-seasonal series.
	for (const n of [3, 4, 5, 6, 7, 8]) {
		const base = Array.from({ length: n }, (_v, i) => round(10 + 0.8 * i + rand(3)));
		out.push(entry(`n${n}-base`, base, 0));
		for (let i = 0; i < n; i++)
			out.push(
				entry(
					`n${n}-imp${i}`,
					base.map((v, j) => (j === i ? round(v + 4) : v)),
					0,
				),
			);
		for (let k = 0; k < 3; k++)
			out.push(
				entry(
					`n${n}-rnd${k}`,
					Array.from({ length: n }, (_v, i) => round(10 + i + rand(3))),
					0,
				),
			);
	}
	return out;
}

/** Series for the seasonality detector: sinusoids, spike trains and noise; one call each. */
export function buildDetectionCases(): EtsCaseInput[] {
	const out: EtsCaseInput[] = [];
	const r = rng(2024);
	const rand = (amp: number): number => (r() - 0.5) * 2 * amp;
	const add = (name: string, values: number[]): void => {
		out.push({
			name,
			values,
			timeline: values.map((_v, i) => i + 1),
			formulas: ['FORECAST.ETS.SEASONALITY({V},{L})'],
		});
	};
	const sine = (p: number, n: number, trend: number, noise: number): number[] =>
		Array.from({ length: n }, (_v, i) =>
			round(50 + 10 * Math.sin((2 * Math.PI * i) / p) + trend * i + rand(noise)),
		);
	for (const p of [2, 3, 4, 5, 6, 7, 8, 10, 12])
		for (const cycles of [3, 4, 6])
			if (p * cycles <= 60) add(`det-sin${p}-n${p * cycles}`, sine(p, p * cycles, 0, 0));
	for (const p of [3, 4, 5, 7, 12]) {
		add(`det-sin${p}-n${p * 3 + 1}`, sine(p, p * 3 + 1, 0, 0));
		add(`det-sin${p}-n${p * 4 + 2}`, sine(p, Math.min(60, p * 4 + 2), 0, 0));
		add(`det-sintrend${p}-n${p * 4}`, sine(p, Math.min(60, p * 4), 0.6, 0));
		add(`det-sintrend${p}-n${p * 3}`, sine(p, p * 3, 1.5, 0));
	}
	for (const p of [4, 7, 12])
		for (const sd of [2, 5, 9, 14]) add(`det-noisy${p}-sd${sd}`, sine(p, 48, 0, sd));
	for (const n of [8, 10, 12, 15, 20, 24, 30, 36, 48, 60, 100])
		for (let s = 0; s < 4; s++)
			add(
				`det-noise-n${n}-${s}`,
				Array.from({ length: n }, () => round(50 + rand(8))),
			);
	for (const p of [3, 4, 5, 7]) {
		add(
			`det-spike${p}-n${6 * p}`,
			Array.from({ length: 6 * p }, (_v, i) => (i % p === 0 ? 20 : 0)),
		);
		add(
			`det-spike${p}-noisy`,
			Array.from({ length: 6 * p }, (_v, i) => round((i % p === 0 ? 20 : 0) + rand(4))),
		);
	}
	add(
		'det-two-4-6',
		Array.from({ length: 48 }, (_v, i) =>
			round(50 + 8 * Math.sin((2 * Math.PI * i) / 4) + 6 * Math.sin((2 * Math.PI * i) / 6)),
		),
	);
	return out;
}

export const ETS_INIT_SERIES_PATH = fileURLToPath(
	new URL('../../xlsx/formula/__fixtures__/ets-init-series.json', import.meta.url),
);
if (import.meta.main)
	writeFileSync(
		ETS_INIT_SERIES_PATH,
		`${JSON.stringify([...buildInitCases(), ...buildDetectionCases()], null, '\t')}\n`,
	);
