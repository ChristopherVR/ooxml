// The series side of an xlsx `ChartPatch` applied to the chart's `ChartSpace`: which series the
// patch keeps, adds or drops (the first chart group, which the summary lists), and what changed on
// each kept series (name, references, cached data, primary paint, effects, marker paint and point
// overrides). Paint lands where the summary reads it back (`chartSummarySeries`): the outline of a
// line, scatter or radar series, the marker of a marker-only scatter series, the fill otherwise,
// so the summary derived from the edited model shows the patched colours. Edits go through the
// `chart/edit` operations where one exists; the rest edit the draft's series in place.
import type { ChartSpace } from '../../chart/model';
import type {
	ChartDataSource,
	ChartShapeProperties,
	ChartSpaceSeries,
} from '../../chart/model-series';
import { assignOrDelete, editChartSpace } from '../../chart/edit/draft';
import {
	addChartSeries,
	moveChartSeries,
	removeChartSeries,
	renameChartSeries,
} from '../../chart/edit/series';
import { categorySource, numberSource } from '../../chart/edit/series-data';
import { setChartDataPointFill } from '../../chart/edit/fill';
import { drawingFillXml } from '../../drawingml/write-fill';
import type { DrawingFill } from '../../drawingml/types';
import type { ChartObject, ChartSeries } from '../model';
import { chartSeriesDrawingFill } from '../write/chart-colors';
import { validChartEffects } from '../write/chart-space';

const same = (a: unknown, b: unknown) => JSON.stringify(a) === JSON.stringify(b);

/** The identity of a series across a patch: its references and literal name. */
const identity = (series: ChartSeries) =>
	JSON.stringify([series.nameRef, series.valuesRef, series.categoriesRef, series.name]);

/**
 * Matches each series of `after` to one of `before` (same position first, else the same
 * references and name); `undefined` for a series the patch adds.
 */
function matchSeries(before: readonly ChartSeries[], after: readonly ChartSeries[]) {
	if (before.length === after.length) return after.map((_, i) => i);
	const free = new Set(before.map((_, i) => i));
	return after.map((series) => {
		const found = [...free].find((i) => identity(before[i]!) === identity(series));
		if (found !== undefined) free.delete(found);
		return found;
	});
}

/** The first group's series list made to follow `after`: dropped, added and reordered. */
function syncMembership(
	space: ChartSpace,
	before: ChartObject,
	after: ChartObject,
): { space: ChartSpace; match: (number | undefined)[] } {
	const match = matchSeries(before.series, after.series);
	if (before.series.length === after.series.length) return { space, match };
	let next = space;
	const kept = new Set(match.filter((i): i is number => i !== undefined));
	for (let i = before.series.length - 1; i >= 0; i--)
		if (!kept.has(i)) next = removeChartSeries(next, i);
	// Positions after the drops, in `before` order, then the added series appended.
	const order = [...kept].sort((a, b) => a - b);
	match.forEach((source, i) => {
		if (source !== undefined) return;
		const series = after.series[i]!;
		next = addChartSeries(next, {
			values: series.values,
			categories: series.categories,
			groupIndex: 0,
			...(series.name === undefined ? {} : { name: series.name }),
			...(series.nameRef === undefined ? {} : { nameFormula: series.nameRef }),
			...(series.valuesRef === undefined ? {} : { valuesFormula: series.valuesRef }),
			...(series.categoriesRef === undefined ? {} : { categoriesFormula: series.categoriesRef }),
		});
		order.push(-1 - i);
	});
	// Move each series to its place in `after`.
	const wanted = match.map((source, i) => (source === undefined ? -1 - i : source));
	for (let to = 0; to < wanted.length; to++) {
		const from = order.indexOf(wanted[to]!);
		if (from !== to) {
			next = moveChartSeries(next, from, to);
			order.splice(to, 0, ...order.splice(from, 1));
		}
	}
	return { space: next, match };
}

/** The source a summary reference names, with the formula changed (an existing reference only). */
function withFormula(source: ChartDataSource | undefined, formula: string | undefined) {
	if (!source?.formula || formula === undefined || source.formula === formula) return source;
	return { ...source, formula };
}

/** The shape properties the series' primary paint lives in, created on the draft when missing. */
function primaryPaint(
	series: ChartSpaceSeries,
	chartType: ChartObject['chartType'],
	scatterStyle: string | undefined,
): { fill?: DrawingFill } {
	if (chartType === 'scatter' && scatterStyle === 'marker')
		return ((series.marker ??= {}).spPr ??= {});
	const shape = (series.spPr ??= {});
	return ['line', 'scatter', 'radar'].includes(chartType) ? (shape.line ??= {}) : shape;
}

/** The point override paint of a summary series point: a writable fill, else the solid colour. */
function pointPaint(series: ChartSeries, point: number): DrawingFill | undefined {
	const fill = series.pointFills?.[point];
	if (fill && drawingFillXml(fill)) return fill;
	const color = series.pointColors?.[point];
	return color ? { kind: 'solid', color } : undefined;
}

/** Copies a marker's shape properties (not its resolved symbol) onto a model marker. */
function markerPaint(
	target: { marker?: ChartSpaceSeries['marker'] },
	spPr: ChartShapeProperties | undefined,
) {
	if (!spPr && !target.marker) return;
	assignOrDelete((target.marker ??= {}), 'spPr', structuredClone(spPr));
}

/** Removes a point override's own fill; the override and its other settings stay. */
function withoutPointFill(space: ChartSpace, index: number, point: number): ChartSpace {
	const override = space.plotArea.groups[0]?.series[index]?.dataPoints.find(
		(item) => item.index === point,
	);
	if (!override?.spPr?.fill) return space;
	return editChartSpace(space, (draft) => {
		const spPr = draft.plotArea.groups[0]!.series[index]!.dataPoints.find(
			(item) => item.index === point,
		)!.spPr!;
		delete spPr.fill;
	});
}

/** Applies what changed between two summary series to the model series at `index`. */
function syncSeries(
	space: ChartSpace,
	index: number,
	before: ChartSeries | undefined,
	after: ChartSeries,
	chart: ChartObject,
	refresh: boolean,
): ChartSpace {
	let next = space;
	// A name taken from a cell follows the cell: only a cache refresh rewrites it.
	if (before && after.name !== undefined && after.name !== before.name)
		if (!after.nameRef || refresh)
			next = renameChartSeries(next, index, after.name, { keepReference: !!after.nameRef });
	const changedPoints = new Set(
		[
			...Object.keys(before?.pointColors ?? {}),
			...Object.keys(before?.pointFills ?? {}),
			...Object.keys(after.pointColors ?? {}),
			...Object.keys(after.pointFills ?? {}),
		].filter(
			(key) =>
				/^\d+$/.test(key) &&
				!same(before && pointPaint(before, Number(key)), pointPaint(after, Number(key))),
		),
	);
	for (const key of changedPoints) {
		const paint = pointPaint(after, Number(key));
		if (paint) next = setChartDataPointFill(next, index, Number(key), paint);
		else next = withoutPointFill(next, index, Number(key));
	}
	return editChartSpace(next, (draft) => {
		const group = draft.plotArea.groups[0]!;
		const series = group.series[index]!;
		const xy = series.yValues !== undefined || series.xValues !== undefined;
		const valueKey = xy ? 'yValues' : 'values';
		const categoryKey = xy ? 'xValues' : 'categories';
		if (series.tx?.reference && after.nameRef !== before?.nameRef)
			assignOrDelete(series.tx, 'reference', withFormula(series.tx.reference, after.nameRef));
		assignOrDelete(series, valueKey, withFormula(series[valueKey], after.valuesRef));
		assignOrDelete(series, categoryKey, withFormula(series[categoryKey], after.categoriesRef));
		if (before && !same(before.values, after.values))
			series[valueKey] = numberSource(after.values, series[valueKey]);
		if (before && !same(before.categories, after.categories))
			series[categoryKey] = categorySource(after.categories, series[categoryKey]);
		if (!before || before.effectsXml !== after.effectsXml) {
			const effects = validChartEffects(after.effectsXml);
			if (effects || series.spPr) assignOrDelete((series.spPr ??= {}), 'effectsXml', effects);
		}
		if (
			!before ||
			!same(before.drawingColor, after.drawingColor) ||
			!same(before.fill, after.fill) ||
			!same(before.color, after.color)
		)
			primaryPaint(series, chart.chartType, group.scatterStyle).fill = chartSeriesDrawingFill(
				chart,
				after,
				index,
			);
		if (!same(before?.marker?.spPr, after.marker?.spPr) && (after.marker?.spPr || series.marker))
			markerPaint(series, after.marker?.spPr);
		for (const point of series.dataPoints) {
			if (point.index === undefined || !point.marker?.spPr) continue;
			const was = before?.pointMarkers?.[point.index]?.spPr;
			const now = after.pointMarkers?.[point.index]?.spPr;
			if (!same(was, now)) markerPaint(point, now);
		}
	});
}

/**
 * The series of `after` applied to `space`, where `before` is the summary `space` was showing:
 * membership first, then each series' changes. `refresh` marks a cache refresh from the cells,
 * which also rewrites the cached name of a series named by a cell.
 */
export function chartSpaceWithSeries(
	space: ChartSpace,
	before: ChartObject,
	after: ChartObject,
	refresh = false,
): ChartSpace {
	if (!space.plotArea.groups[0]) return space;
	const synced = syncMembership(space, before, after);
	let next = synced.space;
	after.series.forEach((series, index) => {
		const source = synced.match[index];
		const was = source === undefined ? undefined : before.series[source];
		if (was && same(was, series)) return;
		next = syncSeries(next, index, was, series, after, refresh);
	});
	return next;
}

/**
 * Rewrites every series reference of `space` in place (rows inserted, a sheet renamed), as the
 * structure edits rewrite the summary's references.
 */
export function rewriteChartSpaceReferences(
	space: ChartSpace,
	rewrite: (formula: string) => string,
): void {
	for (const group of space.plotArea.groups)
		for (const series of group.series)
			for (const source of [
				series.tx?.reference,
				series.categories,
				series.values,
				series.xValues,
				series.yValues,
				series.bubbleSizes,
			])
				if (source?.formula) source.formula = rewrite(source.formula);
}
