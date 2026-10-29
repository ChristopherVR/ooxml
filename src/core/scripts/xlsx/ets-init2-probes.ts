// Second set of starting-state probes for FORECAST.ETS (see ets-init.ts): impulse series over more
// shapes than the first set (non-seasonal 9 to 16 points, seasons 2 to 6 with partial last cycles).
//   bun scripts/xlsx/ets-init2-probes.ts          (from src/core)
// then record Excel's answers with `xlsx/formula/__fixtures__/ets-init2-cases.ps1`. Each group is a
// base series and the base with 5 added at one position, so the difference of the recovered
// starting states is the weight of that position (the map is linear in the data).
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

/** The (season, length) groups: `m` 0 is a series without a season. */
export const INIT2_GROUPS: readonly (readonly [number, number])[] = [
	...[9, 10, 11, 12, 14, 16].map((n) => [0, n] as const),
	...[6, 7, 8, 9, 10].map((n) => [2, n] as const),
	...[6, 7, 8, 9, 10, 11, 12].map((n) => [3, n] as const),
	...[8, 9, 10, 12].map((n) => [4, n] as const),
	...[10, 12].map((n) => [5, n] as const),
	...[12, 13].map((n) => [6, n] as const),
];

export function buildInit2Cases(): EtsCaseInput[] {
	const out: EtsCaseInput[] = [];
	const r = rng(4242);
	const rand = (amp: number): number => (r() - 0.5) * 2 * amp;
	for (const [m, n] of INIT2_GROUPS) {
		const season = (i: number): number => (m >= 2 ? 4 * Math.sin((2 * Math.PI * i) / m) : 0);
		const base = Array.from({ length: n }, (_v, i) =>
			round(20 + 0.5 * i + season(i) + rand(m >= 2 ? 1.5 : 3)),
		);
		const tag = m >= 2 ? `s${m}-${n}` : `n${n}`;
		out.push(entry(`${tag}-base`, base, m));
		for (let i = 0; i < n; i++)
			out.push(
				entry(
					`${tag}-imp${i}`,
					base.map((v, j) => (j === i ? round(v + 5) : v)),
					m,
				),
			);
	}
	return out;
}

/** A third round: longer seasonal series, to see which lengths share one algorithm. */
export const INIT3_GROUPS: readonly (readonly [number, number])[] = [
	...[8, 9, 10, 11, 12, 14].map((n) => [2, n] as const),
	...[9, 13, 15, 18].map((n) => [3, n] as const),
	...[13, 14, 15, 16, 18, 20].map((n) => [4, n] as const),
	...[15].map((n) => [5, n] as const),
	...[18].map((n) => [6, n] as const),
];

export function buildInit3Cases(): EtsCaseInput[] {
	const out: EtsCaseInput[] = [];
	const r = rng(9137);
	const rand = (amp: number): number => (r() - 0.5) * 2 * amp;
	for (const [m, n] of INIT3_GROUPS) {
		const base = Array.from({ length: n }, (_v, i) =>
			round(20 + 0.5 * i + 4 * Math.sin((2 * Math.PI * i) / m) + rand(1.5)),
		);
		out.push(entry(`s${m}-${n}-base`, base, m));
		for (let i = 0; i < n; i++)
			out.push(
				entry(
					`s${m}-${n}-imp${i}`,
					base.map((v, j) => (j === i ? round(v + 5) : v)),
					m,
				),
			);
	}
	return out;
}

/** A fourth round: gentle impulses (1) on calmer series, so Excel's search stays on small steps. */
export const INIT4_GROUPS: readonly (readonly [number, number])[] = [
	...[10].map((n) => [2, n] as const),
	...[10, 11, 18, 21].map((n) => [3, n] as const),
	...[16, 17, 20].map((n) => [4, n] as const),
	...[20].map((n) => [5, n] as const),
	...[24].map((n) => [6, n] as const),
];

export function buildInit4Cases(): EtsCaseInput[] {
	const out: EtsCaseInput[] = [];
	const r = rng(31337);
	const rand = (amp: number): number => (r() - 0.5) * 2 * amp;
	for (const [m, n] of INIT4_GROUPS) {
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

const fixturePath = (name: string): string =>
	fileURLToPath(new URL(`../../xlsx/formula/__fixtures__/${name}`, import.meta.url));
export const ETS_INIT2_SERIES_PATH = fixturePath('ets-init2-series.json');
export const ETS_INIT3_SERIES_PATH = fixturePath('ets-init3-series.json');
export const ETS_INIT4_SERIES_PATH = fixturePath('ets-init4-series.json');
if (import.meta.main) {
	writeFileSync(ETS_INIT2_SERIES_PATH, `${JSON.stringify(buildInit2Cases(), null, '\t')}\n`);
	writeFileSync(ETS_INIT3_SERIES_PATH, `${JSON.stringify(buildInit3Cases(), null, '\t')}\n`);
	writeFileSync(ETS_INIT4_SERIES_PATH, `${JSON.stringify(buildInit4Cases(), null, '\t')}\n`);
}
