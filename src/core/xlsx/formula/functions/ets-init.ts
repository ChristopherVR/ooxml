// Initial states of the exponential smoothing model by least squares. For fixed smoothing
// parameters the one-step fitted values are linear in the initial level, trend and seasonal
// states, so the states that minimise the squared error come from a small linear system. The
// seasonal states are constrained to sum to zero (level and seasonals are otherwise confounded).
import type { Smoothing } from './ets-core';

export interface EtsInit {
	level: number;
	trend: number;
	season: number[];
}

type Runner = (
	y: readonly number[],
	m: number,
	init: EtsInit,
	p: Smoothing,
) => { fitted: number[] };

/** Longest season whose states are solved exactly; longer ones start from a decomposition. */
const MAX_EXACT_SEASON = 24;

const mean = (xs: readonly number[]): number => xs.reduce((a, b) => a + b, 0) / (xs.length || 1);

/** A cheap start for long seasons: first-cycle means around a linear level. */
function heuristicStates(y: readonly number[], m: number): EtsInit {
	if (m < 2) {
		const trend = (y[1] ?? 0) - (y[0] ?? 0);
		return { level: (y[0] ?? 0) - trend, trend, season: [] };
	}
	const first = y.slice(0, m);
	const second = y.slice(m, 2 * m);
	const trend = (mean(second) - mean(first)) / m;
	const detrended = first.map((v, j) => v - trend * (j + 1));
	const level = mean(detrended);
	return { level, trend, season: detrended.map((d) => d - level) };
}

/** Solves the symmetric system `a x = b` in place by Gaussian elimination with pivoting. */
function solveSystem(a: number[][], b: number[]): number[] {
	const k = b.length;
	for (let i = 0; i < k; i++) {
		let p = i;
		for (let r = i + 1; r < k; r++)
			if (Math.abs((a[r] as number[])[i] as number) > Math.abs((a[p] as number[])[i] as number))
				p = r;
		[a[i], a[p]] = [a[p] as number[], a[i] as number[]];
		[b[i], b[p]] = [b[p] as number, b[i] as number];
		const row = a[i] as number[];
		const pivot = row[i] as number;
		if (Math.abs(pivot) < 1e-12) continue;
		for (let r = i + 1; r < k; r++) {
			const other = a[r] as number[];
			const f = (other[i] as number) / pivot;
			for (let c = i; c < k; c++) other[c] = (other[c] as number) - f * (row[c] as number);
			b[r] = (b[r] as number) - f * (b[i] as number);
		}
	}
	const x = new Array<number>(k).fill(0);
	for (let i = k - 1; i >= 0; i--) {
		let s = b[i] as number;
		for (let c = i + 1; c < k; c++) s -= ((a[i] as number[])[c] as number) * (x[c] as number);
		const pivot = (a[i] as number[])[i] as number;
		x[i] = Math.abs(pivot) < 1e-12 ? 0 : s / pivot;
	}
	return x;
}

/** The initial states that minimise the squared one-step error for these smoothing parameters. */
export function initialStates(y: readonly number[], m: number, p: Smoothing, run: Runner): EtsInit {
	const seasonal = m >= 2;
	if (seasonal && m > MAX_EXACT_SEASON) return heuristicStates(y, m);
	const n = y.length;
	const zeros = new Array<number>(n).fill(0);
	const blank: EtsInit = {
		level: 0,
		trend: 0,
		season: seasonal ? new Array<number>(m).fill(0) : [],
	};
	const base = run(y, m, blank, p).fitted;
	// Responses of the fitted values to each unit initial state, with y = 0.
	const unit = (patch: Partial<EtsInit>): number[] =>
		run(zeros, m, { ...blank, ...patch }, p).fitted;
	const columns: number[][] = [unit({ level: 1 }), unit({ trend: 1 })];
	if (seasonal) {
		const last = unit({ season: Array.from({ length: m }, (_v, j) => (j === m - 1 ? 1 : 0)) });
		for (let j = 0; j < m - 1; j++) {
			const s = unit({ season: Array.from({ length: m }, (_v, i) => (i === j ? 1 : 0)) });
			columns.push(s.map((v, i) => v - (last[i] as number)));
		}
	}
	const k = columns.length;
	const target = y.map((v, i) => v - (base[i] as number));
	const gram = columns.map((ci) =>
		columns.map((cj) => ci.reduce((s, v, i) => s + v * (cj[i] as number), 0)),
	);
	for (let i = 0; i < k; i++)
		(gram[i] as number[])[i] = ((gram[i] as number[])[i] as number) + 1e-9;
	const rhs = columns.map((ci) => ci.reduce((s, v, i) => s + v * (target[i] as number), 0));
	const x = solveSystem(gram, rhs);
	const season = seasonal ? (x.slice(2) as number[]) : [];
	if (seasonal) season.push(-season.reduce((a, b) => a + b, 0));
	return { level: x[0] as number, trend: x[1] as number, season };
}
