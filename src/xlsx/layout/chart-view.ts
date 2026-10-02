import type { CellValue, ChartObject, ChartSeries, ChartType, Workbook } from '../model.js';
import { isCellError } from '../model.js';
import { autoSeriesColor } from './chart-colors.js';
import { niceScale, PERCENT_SCALE, type AxisScale } from './chart-scale.js';
import { resolveColor } from './colors.js';

export interface ChartSeriesView {
	name: string;
	values: (number | null)[];
	/** Scatter X values (numeric categories, or 1..n when the categories are text). */
	xValues?: (number | null)[];
	color: string;
	/** Per-point colours (pie and doughnut vary colours by point). */
	pointColors?: string[];
}

export interface ValueAxisView extends AxisScale {
	/** Ticks are fractions shown as percentages (percent-stacked charts). */
	percent: boolean;
}

/** Neutral data for a chart painter: series resolved, colours chosen, axes scaled. */
export interface ChartViewModel {
	type: ChartType;
	grouping: 'clustered' | 'stacked' | 'percentStacked' | 'standard';
	title?: string;
	showLegend: boolean;
	legendPosition: 'r' | 'l' | 't' | 'b' | 'tr';
	categories: string[];
	series: ChartSeriesView[];
	/** Value (Y) axis; for bar charts it runs horizontally. */
	valueAxis?: ValueAxisView;
	/** Scatter X axis. */
	xAxis?: AxisScale;
	/** Bars run horizontally (`bar`), not vertically (`column`). */
	horizontal: boolean;
	/** False for chart types this painter does not draw (bubble, stock, surface). */
	supported: boolean;
}

export type EvaluateRef = (ref: string) => CellValue[];

const SUPPORTED: ReadonlySet<ChartType> = new Set([
	'bar',
	'column',
	'line',
	'area',
	'pie',
	'doughnut',
	'scatter',
	'radar',
]);

function toText(value: CellValue | string | number): string {
	if (value === null) return '';
	if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
	if (isCellError(value)) return value.error;
	return String(value);
}

const toNumber = (value: CellValue | number | null): number | null =>
	typeof value === 'number' && Number.isFinite(value) ? value : null;

function resolveRef(ref: string | undefined, evaluateRef: EvaluateRef): CellValue[] | undefined {
	if (!ref) return undefined;
	try {
		const values = evaluateRef(ref);
		return values.length ? values : undefined;
	} catch {
		return undefined;
	}
}

function seriesName(series: ChartSeries, index: number, evaluateRef: EvaluateRef): string {
	const fromRef = resolveRef(series.nameRef, evaluateRef)
		?.map(toText)
		.filter((t) => t !== '')
		.join(' ');
	return fromRef || series.name || `Series${index + 1}`;
}

/**
 * Resolves a chart's series from live cells (`evaluateRef` returns the values of a reference such
 * as `Sheet1!$B$2:$B$9`, flattened row-major), falling back to the values cached in the chart part.
 */
export function chartView(
	workbook: Workbook,
	sheetIndex: number,
	chart: ChartObject,
	evaluateRef: EvaluateRef,
): ChartViewModel {
	void sheetIndex;
	const type = chart.chartType;
	const theme = workbook.theme;
	const radial = type === 'pie' || type === 'doughnut';
	const grouping =
		chart.grouping ?? (type === 'bar' || type === 'column' ? 'clustered' : 'standard');
	let categories: string[] = [];
	const rawCategories: (CellValue | string | number)[][] = [];
	const series: ChartSeriesView[] = chart.series.map((s, i) => {
		const values = (resolveRef(s.valuesRef, evaluateRef)?.map(toNumber) ??
			s.values.map(toNumber)) as (number | null)[];
		const cats = resolveRef(s.categoriesRef, evaluateRef) ?? s.categories;
		rawCategories.push(cats);
		const color = resolveColor(s.color, theme) ?? autoSeriesColor(theme, i);
		return { name: seriesName(s, i, evaluateRef), values, color };
	});
	const longest = Math.max(0, ...series.map((s) => s.values.length));
	const firstCats = rawCategories.find((c) => c.length > 0) ?? [];
	categories = Array.from({ length: Math.max(longest, firstCats.length) }, (_, i) =>
		firstCats[i] !== undefined ? toText(firstCats[i] ?? null) : String(i + 1),
	);

	if (radial) {
		const first = series[0];
		if (first) first.pointColors = first.values.map((_, i) => autoSeriesColor(theme, i));
		for (const s of series.slice(1))
			s.pointColors = s.values.map((_, i) => autoSeriesColor(theme, i));
	}

	const model: ChartViewModel = {
		type,
		grouping,
		showLegend: chart.showLegend,
		legendPosition: chart.legendPosition ?? 'r',
		categories,
		series,
		horizontal: type === 'bar',
		supported: SUPPORTED.has(type),
	};
	if (chart.title) model.title = chart.title;

	if (type === 'scatter') {
		series.forEach((s, i) => {
			const cats = rawCategories[i] ?? [];
			const numeric = cats.length > 0 && cats.every((c) => typeof c === 'number');
			s.xValues = s.values.map((_, k) =>
				numeric ? toNumber((cats[k] ?? null) as CellValue) : k + 1,
			);
		});
		const xs = series.flatMap((s) => s.xValues ?? []).filter((x): x is number => x !== null);
		model.xAxis = xs.length ? niceScale(Math.min(...xs), Math.max(...xs)) : niceScale(0, 1);
	}
	if (!radial) model.valueAxis = valueAxis(model);
	return model;
}

function valueAxis(model: ChartViewModel): ValueAxisView {
	const stackable =
		model.type === 'bar' ||
		model.type === 'column' ||
		model.type === 'line' ||
		model.type === 'area';
	if (stackable && model.grouping === 'percentStacked')
		return { ...PERCENT_SCALE, ticks: [...PERCENT_SCALE.ticks], percent: true };
	let min = Infinity;
	let max = -Infinity;
	if (stackable && model.grouping === 'stacked') {
		for (let i = 0; i < model.categories.length; i++) {
			let pos = 0;
			let neg = 0;
			for (const s of model.series) {
				const v = s.values[i] ?? 0;
				if (v >= 0) pos += v;
				else neg += v;
			}
			min = Math.min(min, neg);
			max = Math.max(max, pos);
		}
	} else {
		for (const s of model.series)
			for (const v of s.values)
				if (v !== null) {
					min = Math.min(min, v);
					max = Math.max(max, v);
				}
	}
	if (min === Infinity) {
		min = 0;
		max = 1;
	}
	return { ...niceScale(min, max), percent: false };
}

/** Per-category stacked totals of absolute values (for percent-stacked charts). */
export function categoryTotals(model: ChartViewModel): number[] {
	return model.categories.map((_, i) =>
		model.series.reduce((sum, s) => sum + Math.abs(s.values[i] ?? 0), 0),
	);
}
