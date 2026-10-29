// Evaluates the FORECAST.ETS cases recorded from real Excel (__fixtures__/ets-excel.json and
// ets-search-excel.json) with this engine. Used by forecast-ets-excel.test.ts and by the error report in docs.
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { book, engine, get } from './test-helpers';

export interface EtsRecorded {
	name: string;
	values: (number | null)[];
	timeline: number[];
	results: { formula: string; type: string; value?: number | string }[];
}

export type EtsKind = 'forecast' | 'confint' | 'seasonality' | `stat${number}`;

export interface EtsRow {
	name: string;
	formula: string;
	kind: EtsKind;
	excel: number | string;
	ours: unknown;
	/** Relative error (absolute when Excel's value is 0); NaN when the results are not comparable. */
	error: number;
}

export function recorded(): EtsRecorded[] {
	const path = fileURLToPath(new URL('./__fixtures__/ets-excel.json', import.meta.url));
	return JSON.parse(readFileSync(path, 'utf8')) as EtsRecorded[];
}

/**
 * The 491 non-seasonal probe series of ets-search-series.json with Excel's answers
 * (ets-search-excel.json), used to identify how Excel searches for alpha, beta and gamma.
 */
export function searchRecorded(): EtsRecorded[] {
	const path = fileURLToPath(new URL('./__fixtures__/ets-search-excel.json', import.meta.url));
	return JSON.parse(readFileSync(path, 'utf8')) as EtsRecorded[];
}

export function kindOf(formula: string): EtsKind {
	if (formula.startsWith('FORECAST.ETS.CONFINT')) return 'confint';
	if (formula.startsWith('FORECAST.ETS.SEASONALITY')) return 'seasonality';
	const stat = /^FORECAST\.ETS\.STAT\([^,]+,[^,]+,(\d)/.exec(formula);
	if (stat) return `stat${stat[1]}` as EtsKind;
	return 'forecast';
}

export function evaluateEts(c: EtsRecorded): EtsRow[] {
	const n = c.values.length;
	const cells: Record<string, number> = {};
	c.values.forEach((v, i) => {
		if (v !== null) cells[`A${i + 1}`] = v;
	});
	c.timeline.forEach((t, i) => (cells[`B${i + 1}`] = t));
	const formulas = c.results.map((r) =>
		r.formula.replace(/\{V\}/g, `A1:A${n}`).replace(/\{L\}/g, `B1:B${n}`),
	);
	const wb = book(
		{
			...cells,
			...Object.fromEntries(formulas.map((f, i) => [`D${i + 1}`, `=${f}` as const])),
		},
		['Sheet1'],
	);
	engine(wb).recalculateAll();
	return c.results.map((r, i) => {
		const ours = get(wb, `D${i + 1}`);
		const excel = r.value as number | string;
		let error = Number.NaN;
		if (r.type === 'number' && typeof ours === 'number')
			error = Math.abs(ours - (excel as number)) / Math.max(1, Math.abs(excel as number));
		return { name: c.name, formula: r.formula, kind: kindOf(r.formula), excel, ours, error };
	});
}
