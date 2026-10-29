// Excel's FORECAST.ETS starting state as a linear map of the data, read off the impulse probes
// (ets-init-excel.json and the ets-init2/3/4 files, series written by scripts/xlsx/ets-init2-probes.ts):
// a group is a base series and the base with a known amount added at one position, so the
// difference of the recovered states, divided by the amount, is that position's weight. The map is
// parameter-free and keeps a periodic pattern out of the trend, so a probe that Excel answered from
// extreme smoothing parameters (the recovery then fails the first-value law) is rebuilt from the rule
// that the weights of one position (mod season) sum to 0 for the trend and to 1 for that position's
// own state, 0 for the others.
import { type EtsInitCase, recoveredInits } from './ets-init-compare';

export interface StartMap {
	/** Group name, `s<season>-<length>` or `n<length>` without a season. */
	key: string;
	/** The probe file the group comes from (groups of the same name differ between files). */
	source: string;
	/** Season length, 0 without. */
	m: number;
	n: number;
	/**
	 * Weights over the data of the level plus season plus trend per position (the fitted value for
	 * that position at time 0): one row per position, a single row without a season.
	 */
	z: number[][];
	/** Weights of the starting trend. */
	trend: number[];
	/** Probes that could not be recovered or rebuilt (the map is incomplete when above 0). */
	missing: number;
}

const FILES = [
	'ets-init-excel.json',
	'ets-init2-excel.json',
	'ets-init3-excel.json',
	'ets-init4-excel.json',
	'ets-init5-excel.json',
	'ets-init6-excel.json',
	'ets-init7-excel.json',
	'ets-init8-excel.json',
	'ets-init9-excel.json',
	'ets-init10-excel.json',
	'ets-init11-excel.json',
	'ets-init12-excel.json',
];

const FIRST_VALUE_TOLERANCE = 1e-6;

function groupMap(key: string, source: string, cases: EtsInitCase[]): StartMap | undefined {
	const base = cases.find((c) => c.name.endsWith('-base'));
	if (!base) return undefined;
	const n = base.y.length;
	const m = base.m;
	const positions = m >= 2 ? m : 1;
	const probes = Array.from({ length: n }, (_v, i) =>
		cases.find((c) => c.name.endsWith(`-imp${i}`)),
	);
	const usable = (c: EtsInitCase | undefined): c is EtsInitCase =>
		c !== undefined && Math.abs(c.firstError) < FIRST_VALUE_TOLERANCE;
	if (!usable(base)) return undefined;
	const column = (c: EtsInitCase | undefined, row: number): number => {
		if (!usable(c)) return Number.NaN;
		const amount = (c.y[probes.indexOf(c)] as number) - (base.y[probes.indexOf(c)] as number);
		const value = (x: EtsInitCase): number =>
			row < positions ? (x.c[row] as number) + x.trend : x.trend;
		return (value(c) - value(base)) / amount;
	};
	const rows = Array.from({ length: positions + 1 }, (_v, r) => probes.map((c) => column(c, r)));
	const lost = probes.map((c, i) => (usable(c) ? -1 : i)).filter((i) => i >= 0);
	let missing = 0;
	rows.forEach((row, r) => {
		for (const i of lost) {
			const same = lost.filter((j) => j % positions === i % positions);
			if (same.length !== 1) {
				missing++;
				continue;
			}
			let others = 0;
			for (let t = i % positions; t < n; t += positions) if (t !== i) others += row[t] as number;
			const total = r === positions ? 0 : r === i % positions ? 1 : 0;
			row[i] = total - others;
		}
	});
	return {
		key,
		source,
		m,
		n,
		z: rows.slice(0, positions),
		trend: rows[positions] as number[],
		missing,
	};
}

/** Every probe group of the four files with the map recovered from it. */
export function recoveredMaps(): StartMap[] {
	const out: StartMap[] = [];
	for (const file of FILES) {
		const groups = new Map<string, EtsInitCase[]>();
		for (const c of recoveredInits(file)) {
			const key = c.name.replace(/-(base|imp\d+|rnd\d+)$/, '');
			groups.set(key, [...(groups.get(key) ?? []), c]);
		}
		for (const [key, cases] of groups) {
			const map = groupMap(key, file, cases);
			if (map) out.push(map);
		}
	}
	return out;
}
