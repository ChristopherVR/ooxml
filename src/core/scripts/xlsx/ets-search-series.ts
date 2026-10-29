// Writes the probe series for Excel's FORECAST.ETS parameter search (non-seasonal, season passed as
// 0, so detection and seasonal starting values play no part):
//   bun scripts/xlsx/ets-search-series.ts [set]          (from src/core)
// then record Excel's answers with `xlsx/formula/__fixtures__/ets-search-cases.ps1 -In ... -Out ...`.
// Sets: `all` (the default, written to ets-search-series.json), or one of its parts: `random` (300
// mixed series), `walks` (invariance probes on random walks), `sweep` (AR(1) with a swept coefficient).
import { writeFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import type { EtsCaseInput } from './ets-series';

export function rng(seed: number): () => number {
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

export type Kind = 'walk' | 'drift' | 'trend' | 'ar' | 'shift' | 'noise';
export const KINDS: Kind[] = ['walk', 'drift', 'trend', 'ar', 'shift', 'noise'];

export function make(kind: Kind, n: number, r: () => number): number[] {
	const g = (): number => (r() + r() + r() + r() - 2) * 1.7;
	const out: number[] = [];
	let x = 50;
	for (let i = 0; i < n; i++) {
		if (kind === 'walk') x += g();
		else if (kind === 'drift') x += 0.6 + g();
		else if (kind === 'trend') x = 50 + 0.8 * i + 2 * g();
		else if (kind === 'ar') x = 50 + 0.7 * (x - 50) + 2 * g();
		else if (kind === 'shift') x = (i < n / 2 ? 50 : 58) + 1.5 * g();
		else x = 50 + 3 * g();
		out.push(round(x));
	}
	return out;
}

/** The probe calls: alpha, beta, RMSE, two forecasts, gamma, MASE, SMAPE, MAE, a longer and a back forecast, an interval. */
export function search(name: string, values: number[]): EtsCaseInput {
	const n = values.length;
	return {
		name,
		values,
		timeline: Array.from({ length: n }, (_x, k) => k + 1),
		formulas: [
			'FORECAST.ETS.STAT({V},{L},1,0)',
			'FORECAST.ETS.STAT({V},{L},2,0)',
			'FORECAST.ETS.STAT({V},{L},7,0)',
			`FORECAST.ETS(${n + 1},{V},{L},0)`,
			`FORECAST.ETS(${n + 2},{V},{L},0)`,
			'FORECAST.ETS.STAT({V},{L},3,0)',
			'FORECAST.ETS.STAT({V},{L},4,0)',
			'FORECAST.ETS.STAT({V},{L},5,0)',
			'FORECAST.ETS.STAT({V},{L},6,0)',
			`FORECAST.ETS(${n + 4},{V},{L},0)`,
			`FORECAST.ETS(${n - 3},{V},{L},0)`,
			`FORECAST.ETS.CONFINT(${n + 1},{V},{L},0.95,0)`,
		],
	};
}

/** Search probes: [name, values]; the lengths cycle through 8 to 41. */
export function buildSearchCases(count = 300): EtsCaseInput[] {
	const r = rng(2024);
	return Array.from({ length: count }, (_v, i) => {
		const kind = KINDS[i % KINDS.length] as Kind;
		const n = 8 + ((i * 7) % 34);
		return search(`${kind}-${n}-${i}`, make(kind, n, r));
	});
}

/** The same few random walks with a slope added, an offset, a scale and shorter lengths. */
export function buildWalkCases(): EtsCaseInput[] {
	const out: EtsCaseInput[] = [];
	for (let w = 0; w < 5; w++) {
		const base = make('walk', 30, rng(900 + w));
		for (const k of [-0.5, 0, 0.25, 0.5, 1, 2])
			out.push(
				search(
					`w${w}-slope${k}`,
					base.map((v, i) => round(v + k * i)),
				),
			);
		out.push(
			search(
				`w${w}-offset`,
				base.map((v) => round(v + 1000)),
			),
		);
		out.push(
			search(
				`w${w}-scale`,
				base.map((v) => round(v * 10)),
			),
		);
		for (const n of [10, 14, 18, 22, 26]) out.push(search(`w${w}-n${n}`, base.slice(0, n)));
	}
	return out;
}

/** AR(1) series on shared innovations, the coefficient swept from 0 (noise) to 1 (random walk). */
export function buildSweepCases(): EtsCaseInput[] {
	const out: EtsCaseInput[] = [];
	for (const seed of [31, 32, 33])
		for (const n of [12, 30]) {
			const r = rng(seed);
			const eps = Array.from({ length: n }, () => (r() + r() + r() + r() - 2) * 1.7);
			for (let k = 0; k <= 20; k++) {
				const phi = k / 20;
				let x = 0;
				out.push(
					search(
						`sweep-s${seed}-n${n}-p${k}`,
						eps.map((e) => {
							x = phi * x + e;
							return round(50 + x);
						}),
					),
				);
			}
		}
	return out;
}

const sets: Record<string, () => EtsCaseInput[]> = {
	random: buildSearchCases,
	walks: buildWalkCases,
	sweep: buildSweepCases,
	all: () => [...buildSearchCases(), ...buildWalkCases(), ...buildSweepCases()],
};
const SERIES_DIR = '../../xlsx/formula/__fixtures__/';
export const ETS_SEARCH_SERIES_PATH = fileURLToPath(
	new URL(`${SERIES_DIR}ets-search-series.json`, import.meta.url),
);
if (import.meta.main) {
	const set = process.argv[2] ?? 'all';
	const file = set === 'all' ? 'ets-search-series.json' : `ets-search-series-${set}.json`;
	const out = fileURLToPath(new URL(SERIES_DIR + file, import.meta.url));
	const cases = (sets[set] ?? buildSearchCases)();
	writeFileSync(out, `[\n${cases.map((c) => `\t${JSON.stringify(c)}`).join(',\n')}\n]\n`);
}
