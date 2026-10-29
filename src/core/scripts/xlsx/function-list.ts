// Generates viewers/xlsx/docs/functions.md from the function registry:
//   bun scripts/xlsx/function-list.ts          (from src/core) rewrites the file
// `xlsx/formula/function-list.test.ts` fails when the committed file differs from this output.
import { readFileSync, writeFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { FUNCTION_CATALOG } from '../../xlsx/formula/functions/registry';

const root = (path: string): URL => new URL(path, import.meta.url);

/** Known deviations from Excel, by function name. Every name must exist in the registry. */
export const DEVIATIONS: Readonly<Record<string, string>> = {
	'FORECAST.ETS':
		"The filter, starting values and error statistics were identified against Excel on 34 series: 21 reproduce to the last digit, typical series are within 0.5%, the worst is 14% (zero-filled gaps, where this finds a period of 5 and Excel none; the seasonal calls on SUM-aggregated duplicates are within 5%; with a season of 0 every aggregation is exact). Excel's search for alpha, beta and gamma minimises the squared one-step error of the training part plus the squared forecast error of the hold-out (the last 40% of the series, at most the larger of 10, two seasons and 20%; scale, offset and sign do not matter) and searches a fixed lattice of 25 (alpha, beta) pairs: without a season the lowest-cost pair is the one Excel fits on every one of about 2100 non-seasonal probes (alpha on all of them, every error statistic STAT 4-7 to the last digit, series of 4 to 7 points being extended to 7 for the search, 8 to 10 points using the plain one-step error and 3 points always 0.1 and 0.099). What Excel reports is the pair's alpha and either its beta or, for a large beta, 0.001 with a forecast of the last value plus the least-squares slope (the pair still measures the errors); which of the two happens is only partly predicted (a few pairs always drift, the rest by the hold-out mean and the slope strength: 83% of 643 probes, so about 1 forecast in 6 of those series differs, up to 13%), and 95% of all forecasts on the 491 search probes equal Excel's; with a season whose starting states are Excel's own (4 whole seasons, 5 and more, 3 whole seasons plus up to ceil(m/2) + 1 points, exactly 2 seasons; seasons of 4 and more for the last two) the same cost is used over the 20 (alpha, beta, gamma) triples Excel is seen to report, on the series padded the way the starting trend pads it (its triple on 2307 of 2319 recorded series), and elsewhere the plain one-step error. Starting states: the starting trend is exact (recovered to 1e-15 from Excel's forecasts) for every series without a season, 3 to 16 points. For seasonal series the start is a fixed linear function of the data, recovered exactly (ets-init-chain.ts, ets-init-chain.test.ts): for 4 whole seasons or 5 and more (only the first 5 seasons are read) the position states follow each other along m consecutive times through first-seasons means and half-split slopes of the positions and the start trend is the least-squares slope of the series without them, which reproduces Excel's forecasts at its own parameters to rounding error on about 1200 probe series (seasons 2 to 12). 3 whole seasons plus up to ceil(m/2) + 1 points and exactly 2 seasons (seasons of 4 and more) are exact too (ets-init-three.ts: a blend of a forward and a backward chain with a slope-sum drift and a corner term, a trend fitted on the series padded to 3m + ceil(m/2) points, and 2 seasons extended to 3), to rounding error on about 1100 more probe series (seasons 4 to 12). Not exact: 4 seasons plus a partial one (a mean over the windows of 4 whole seasons, within 5% of Excel's forecasts at its parameters), 3 seasons plus more than ceil(m/2) + 1 points, 2 seasons plus a partial one and seasons of 2 or 3 below 4 whole seasons (least squares, up to 28% off). Ruled out against the recovered maps: single and multi-pass backcasting with fixed parameters on a 0.01 lattice, weighted regressions with geometric weights, plain seasonal-dummy regression and classical decomposition.",
	'FORECAST.ETS.CONFINT':
		'RMSE times sqrt(1 + sum of (alpha + j*beta)^2) as in Excel: equal to Excel where the fitted parameters are, up to 37% off where they differ (mostly seasonal models, whose gamma term was not checked against Excel).',
	'FORECAST.ETS.SEASONALITY':
		'First peak of the autocorrelation of the detrended series that reaches 0.3 (lags 2 to a third of the series); agrees with Excel on 113 of 126 recorded series, all noiseless waves, trends and impulse trains included. 0 means none. Excel also reports a period on about a third of short white-noise series (8 to 36 points) by a rule that is not the autocorrelation, and this does not. The noise picks of Excel do not follow the best-fitting seasonal ETS either (checked for periods up to 8), nor the lowest hold-out cost of the seasonal fit made here (rechecked with the exact starting states for periods whose series hold 4 or more seasons, on the 44 noise series: the lowest-cost period is neither the period Excel reports nor no period), nor the largest autocorrelation, periodogram peak or differenced autocorrelation (the period is the lag of the largest autocorrelation on 6 of the 10 noisy series where Excel reports one, but the same statistic is as large on series where it reports none).',
	'FORECAST.ETS.STAT':
		'Step size agrees; the error statistics are measured on the last points as in Excel and equal it where the parameters do; alpha is found on 25 of 32 recorded series, beta on 29 (on the 491 non-seasonal probes: alpha on every one, beta on 477).',
	BESSELJ: 'About 1e-7 from the value Microsoft documents for some arguments.',
	GETPIVOTDATA: "Reads this editor's worksheet-range pivots; `#REF!` for anything else.",
	GROUPBY: 'Subset: one aggregating function, grand totals only, flat output.',
	PIVOTBY: 'Subset: one aggregating function, one column field, grand totals only.',
	TRIMRANGE: 'Returns the trimmed values as an array rather than a reference.',
	ODDFPRICE: 'Quasi-coupon method; long odd first periods not checked against Excel.',
	ODDFYIELD: 'Quasi-coupon method; long odd first periods not checked against Excel.',
	ODDLPRICE: 'Quasi-coupon method; not checked against Excel.',
	ODDLYIELD: 'Quasi-coupon method; not checked against Excel.',
};

const FIXTURES = [
	'excel-results.json',
	'depreciation-excel.json',
	'amorlinc-excel.json',
	'lookup-native.json',
	'ets-excel.json',
];

const upperCalls = (formula: string): string[] =>
	[...formula.matchAll(/\b([A-Z][A-Z0-9.]*)\s*\(/g)].map((m) => m[1] as string);

/** Function names used by cases whose expected value was recorded from real Excel. */
export function excelVerifiedNames(): Set<string> {
	const names = new Set<string>();
	const read = (name: string): string =>
		readFileSync(root(`../../xlsx/formula/__fixtures__/${name}`), 'utf8');
	for (const file of FIXTURES) {
		const json = JSON.parse(read(file)) as unknown;
		const list = (Array.isArray(json) ? json : (json as { cases: unknown[] }).cases) as (
			| { formula: string; type?: string }
			| { results: { formula: string; type?: string }[] }
		)[];
		const cases = list.flatMap((c) => ('results' in c ? c.results : [c]));
		for (const c of cases)
			if (c.type !== 'rejected') for (const n of upperCalls(c.formula)) names.add(n);
	}
	return names;
}

const cell = (text: string): string => text.replace(/\\/g, '\\\\').replace(/\|/g, '\\|');

export function buildFunctionList(): string {
	const verified = excelVerifiedNames();
	const byCategory = new Map<string, typeof FUNCTION_CATALOG>();
	for (const fn of FUNCTION_CATALOG)
		byCategory.set(fn.category, [...(byCategory.get(fn.category) ?? []), fn]);
	const checked = FUNCTION_CATALOG.filter((f) => verified.has(f.name)).length;
	const out = [
		'# Supported functions',
		'',
		'<!-- Generated by `bun scripts/xlsx/function-list.ts` in src/core; do not edit by hand. -->',
		'',
		`The calculation engine implements ${FUNCTION_CATALOG.length} functions. ${checked} of them appear in tests whose`,
		'expected values were recorded from real Excel (`src/core/xlsx/formula/__fixtures__`); the rest are',
		'covered by unit tests only, with expected values worked out by hand or from documentation.',
		'A function missing from this list returns `#NAME?`.',
		'',
		'Verification: **Excel** means at least one recorded Excel result uses the function (it may be',
		'nested in a larger formula, and one case does not prove every argument combination);',
		'**Unit** means unit tests only.',
		'',
	];
	for (const [category, fns] of [...byCategory].sort(([a], [b]) => a.localeCompare(b))) {
		out.push(
			`## ${category}`,
			'',
			'| Function | Syntax | Verification | Notes |',
			'| --- | --- | --- | --- |',
		);
		for (const fn of fns)
			out.push(
				`| \`${fn.name}\` | \`${cell(fn.syntax)}\` | ${verified.has(fn.name) ? 'Excel' : 'Unit'} | ${cell(DEVIATIONS[fn.name] ?? '')} |`,
			);
		out.push('');
	}
	return `${out.join('\n')}`;
}

export const FUNCTION_LIST_PATH = fileURLToPath(root('../../../../viewers/xlsx/docs/functions.md'));

if (import.meta.main) writeFileSync(FUNCTION_LIST_PATH, buildFunctionList());
