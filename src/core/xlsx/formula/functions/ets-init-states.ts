// The chain of 4 and more seasons behind the seasonal FORECAST.ETS start (laws in ets-init-chain.ts).

const mean = (xs: readonly number[]): number => xs.reduce((a, b) => a + b, 0) / (xs.length || 1);

/** Seasonal states by position (not yet relative to position 0), from the first `C * m` values. */
export function chainStates(y: readonly number[], m: number, seasons: number): number[] {
	const h = Math.floor(seasons / 2);
	const first = seasons - h;
	const at = (t: number): number => y[t] as number;
	const meanFirst = (t: number): number =>
		mean(Array.from({ length: first }, (_v, c) => at(t + c * m)));
	const slope = (j: number): number => {
		const early = mean(Array.from({ length: h }, (_v, c) => at(j + c * m)));
		const late = mean(Array.from({ length: h }, (_v, c) => at(j + (seasons - h + c) * m)));
		return (late - early) / ((seasons - h) * m);
	};
	const j0 = Math.floor(m / 2);
	const state = new Map<number, number>([[j0 - 1, 0]]);
	for (let t = j0; t <= j0 + m - 2; t++)
		state.set(t, (state.get(t - 1) as number) + meanFirst(t) - meanFirst(t - 1) - slope(t - j0));
	const byPosition: number[] = new Array<number>(m).fill(0);
	for (const [t, s] of state) byPosition[t % m] = s;
	return byPosition;
}

/** Seasonal states relative to position 0 (`s[k] - s[0]`), from the first `C * m` values. */
export function chainSeasonals(y: readonly number[], m: number, seasons: number): number[] {
	const byPosition = chainStates(y, m, seasons);
	const origin = byPosition[0] as number;
	return byPosition.map((s) => s - origin);
}
