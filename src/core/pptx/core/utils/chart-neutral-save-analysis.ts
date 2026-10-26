import {
	chartSeriesEntryAt,
	chartTrendlineType,
	setChartTrendline,
	setChartErrorBars,
	setChartMarker,
	setChartSeriesOptions,
	setChartSeriesAxis,
	editChartShapeProperties,
	numberSource,
	type ChartTrendlineEdit,
	type ChartErrorBarsEdit,
	type ChartSeriesOptionsEdit,
	type ChartShapeProperties,
	type ChartTrendlineLabel,
} from '../../../chart';
import type {
	PptxChartData,
	PptxChartTrendline,
	PptxChartErrBars,
	PptxChartMarker,
} from '../types';
import type { NeutralChartEdit } from './chart-neutral-save-plan';
import { same } from './chart-neutral-save-common';
import { neutralShapeEdit } from './chart-neutral-save-shape';
import { planNeutralChartPoints } from './chart-neutral-save-points';

const lineShape = (value: PptxChartTrendline | PptxChartErrBars | undefined) =>
	value && {
		strokeColor: value.color,
		strokeWidth: 'trendlineType' in value ? value.lineWidth : value.width,
		strokeDashStyle: 'trendlineType' in value ? value.lineDashStyle : value.dashStyle,
	};
function shape(
	current: ChartShapeProperties | undefined,
	before: PptxChartTrendline | PptxChartErrBars | undefined,
	after: PptxChartTrendline | PptxChartErrBars,
) {
	return editChartShapeProperties(current, neutralShapeEdit(lineShape(before), lineShape(after)));
}
function label(
	current: ChartTrendlineLabel | undefined,
	old: PptxChartTrendline['label'],
	next: PptxChartTrendline['label'],
): ChartTrendlineLabel | null {
	if (next == null) return null;
	const result = structuredClone(current ?? {});
	if (!same(old?.layout, next.layout) && next.layout !== undefined) {
		const { ext: _ext, ...layout } = next.layout;
		result.layout = { ...result.layout, ...layout };
	}
	if (next.numberFormatCode !== undefined || next.sourceLinked !== undefined)
		result.numberFormat = {
			formatCode: result.numberFormat?.formatCode ?? 'General',
			...result.numberFormat,
			...(next.numberFormatCode === undefined ? {} : { formatCode: next.numberFormatCode }),
			...(next.sourceLinked === undefined ? {} : { sourceLinked: next.sourceLinked }),
		};
	return result;
}
function trendline(
	edits: NeutralChartEdit[],
	seriesIndex: number,
	index: number,
	before: PptxChartTrendline | undefined,
	after: PptxChartTrendline,
) {
	if (!same(before?.label?.layout?.ext, after.label?.layout?.ext)) return false;
	const edit: ChartTrendlineEdit = {};
	if (before?.trendlineType !== after.trendlineType)
		edit.type = chartTrendlineType(after.trendlineType);
	const fields = {
		name: 'name',
		order: 'order',
		period: 'period',
		forward: 'forward',
		backward: 'backward',
		intercept: 'intercept',
		displayEq: 'displayEquation',
		displayRSq: 'displayRSquared',
	} as const;
	for (const [source, target] of Object.entries(fields)) {
		const key = source as keyof typeof fields;
		if (!same(before?.[key], after[key])) Object.assign(edit, { [target]: after[key] ?? null });
	}
	edits.push((space) => {
		const current = chartSeriesEntryAt(space, seriesIndex).series.trendlines?.[index];
		const applied = { ...edit };
		if (!same(lineShape(before), lineShape(after)))
			applied.spPr = shape(current?.spPr, before, after);
		if (!same(before?.label, after.label))
			applied.label = label(current?.label, before?.label, after.label);
		return setChartTrendline(space, seriesIndex, index, applied);
	});
	return true;
}
function errorBars(
	edits: NeutralChartEdit[],
	seriesIndex: number,
	before: PptxChartErrBars | undefined,
	after: PptxChartErrBars,
) {
	const edit: ChartErrorBarsEdit = {};
	for (const [key, target] of [
		['barType', 'barType'],
		['valType', 'valueType'],
		['val', 'value'],
		['noEndCap', 'noEndCap'],
	] as const)
		if (!same(before?.[key], after[key])) Object.assign(edit, { [target]: after[key] ?? null });
	if (after.valType === 'cust')
		for (const [key, target] of [
			['customPlus', 'plus'],
			['customMinus', 'minus'],
		] as const)
			if (!same(before?.[key], after[key]))
				edit[target] = after[key] === undefined ? null : numberSource(after[key]!, undefined);
	edits.push((space) => {
		const current = chartSeriesEntryAt(space, seriesIndex).series.errorBars?.find(
			(item) => (item.direction ?? 'y') === after.direction,
		);
		const applied = { ...edit };
		if (!same(lineShape(before), lineShape(after)))
			applied.spPr = shape(current?.spPr, before, after);
		return setChartErrorBars(space, seriesIndex, after.direction, applied);
	});
}
function marker(
	edits: NeutralChartEdit[],
	seriesIndex: number,
	before: PptxChartMarker | undefined,
	after: PptxChartMarker | undefined,
) {
	edits.push((space) => {
		if (!after) return setChartMarker(space, seriesIndex, undefined);
		const current = chartSeriesEntryAt(space, seriesIndex).series.marker;
		const next = { ...current, symbol: after.symbol, size: after.size ?? null };
		if (!same(before?.spPr, after.spPr))
			next.spPr = editChartShapeProperties(
				current?.spPr,
				neutralShapeEdit(before?.spPr, after.spPr),
			);
		return setChartMarker(space, seriesIndex, next);
	});
}
/** Claims only source-preserving analysis, marker and series-option fields; data caches stay untouched. */
export function planNeutralChartAnalysis(
	before: PptxChartData,
	probe: PptxChartData,
	edits: NeutralChartEdit[],
): void {
	if (before.series.length !== probe.series.length) return;
	probe.series.forEach((series, seriesIndex) => {
		const old = before.series[seriesIndex]!;
		if (!same(old.trendlines, series.trendlines)) {
			const changes: NeutralChartEdit[] = [];
			const previous = old.trendlines ?? [],
				next = series.trendlines ?? [];
			if (
				next.every(
					(value, index) =>
						same(previous[index], value) ||
						trendline(changes, seriesIndex, index, previous[index], value),
				)
			) {
				for (let index = previous.length - 1; index >= next.length; index--)
					changes.push((space) => setChartTrendline(space, seriesIndex, index, null));
				edits.push(...changes);
				series.trendlines = old.trendlines;
			}
		}
		if (!same(old.errBars, series.errBars)) {
			const previous = old.errBars ?? [],
				next = series.errBars ?? [];
			if (
				[previous, next].every(
					(list) => new Set(list.map((item) => item.direction)).size === list.length,
				)
			) {
				for (const item of previous)
					if (!next.some((value) => value.direction === item.direction))
						edits.push((space) => setChartErrorBars(space, seriesIndex, item.direction, null));
				for (const item of next) {
					const was = previous.find((value) => value.direction === item.direction);
					if (!same(was, item)) errorBars(edits, seriesIndex, was, item);
				}
				series.errBars = old.errBars;
			}
		}
		if (!same(old.marker, series.marker)) {
			marker(edits, seriesIndex, old.marker, series.marker);
			series.marker = old.marker;
		}
		if (old.axisId !== series.axisId) {
			const id = series.axisId ?? null;
			edits.push((space) => setChartSeriesAxis(space, seriesIndex, id));
			series.axisId = old.axisId;
		}
		const options: ChartSeriesOptionsEdit = {};
		for (const key of ['smooth', 'invertIfNegative', 'explosion', 'shape'] as const)
			if (!same(old[key], series[key])) {
				Object.assign(options, { [key]: series[key] ?? null });
				Object.assign(series, { [key]: old[key] });
			}
		if (Object.keys(options).length)
			edits.push((space) => setChartSeriesOptions(space, seriesIndex, options));
	});
	planNeutralChartPoints(before, probe, edits);
}
