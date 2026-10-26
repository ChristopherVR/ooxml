// Shared series analysis edits. Unsupported chart families are preserved but are not offered
// authoring controls. Each edit copies the model and retains unrelated labels/styles/extensions.
import type { ChartSpace } from '../model';
import type {
	ChartErrorBars,
	ChartErrorBarDirection,
	ChartTrendline,
	ChartTrendlineType,
	ChartTrendlineLabel,
} from '../model-analysis';
import type { ChartDataSource, ChartShapeProperties } from '../model-series';
import { assignOrDelete, chartSeriesEntryAt, editChartSpace } from './draft';

export interface ChartTrendlineEdit {
	type?: ChartTrendlineType;
	name?: string | null;
	order?: number | null;
	period?: number | null;
	forward?: number | null;
	backward?: number | null;
	intercept?: number | null;
	displayEquation?: boolean | null;
	displayRSquared?: boolean | null;
	label?: ChartTrendlineLabel | null;
	spPr?: ChartShapeProperties | null;
}

/** Whether a series can author trendlines or error bars; imported unsupported data stays intact. */
export function chartSeriesAnalysisSupported(space: ChartSpace, seriesIndex: number): boolean {
	const { group } = chartSeriesEntryAt(space, seriesIndex);
	return (
		['bar', 'line', 'area', 'scatter', 'bubble'].includes(group.kind) &&
		!group.element.includes('3D') &&
		!['stacked', 'percentStacked'].includes(group.grouping ?? '')
	);
}

function admitted(space: ChartSpace, seriesIndex: number): void {
	if (!chartSeriesAnalysisSupported(space, seriesIndex))
		throw new RangeError('This chart family does not support series analysis editing');
}

const TYPES = new Set<ChartTrendlineType>(['linear', 'exp', 'log', 'poly', 'power', 'movingAvg']);

/** Adds/edits a trendline by index, or removes it with null; no existing trendline is replaced implicitly. */
export function setChartTrendline(
	space: ChartSpace,
	seriesIndex: number,
	index: number,
	edit: ChartTrendlineEdit | null,
): ChartSpace {
	admitted(space, seriesIndex);
	return editChartSpace(space, (draft) => {
		const series = chartSeriesEntryAt(draft, seriesIndex).series;
		const list = (series.trendlines ??= []);
		if (
			!Number.isSafeInteger(index) ||
			index < 0 ||
			index > list.length ||
			(edit === null && index === list.length)
		)
			throw new RangeError('No trendline at this index');
		if (edit === null) {
			list.splice(index, 1);
			return;
		}
		const next: ChartTrendline = { ...(list[index] ?? { trendlineType: 'linear' }) };
		if (edit.type !== undefined) {
			if (!TYPES.has(edit.type)) throw new RangeError('Unsupported trendline type');
			next.trendlineType = edit.type;
			if (edit.type !== 'poly') delete next.order;
			if (edit.type !== 'movingAvg') delete next.period;
		}
		for (const key of ['order', 'period', 'forward', 'backward', 'intercept'] as const) {
			const value = edit[key];
			if (value === undefined) continue;
			if (
				value !== null &&
				(!Number.isFinite(value) ||
					(key === 'order' && (!Number.isInteger(value) || value < 2 || value > 6)) ||
					(key === 'period' && (!Number.isInteger(value) || value < 2 || value > 2147483647)) ||
					((key === 'forward' || key === 'backward') && value < 0))
			)
				throw new RangeError(`Invalid trendline ${key}`);
			assignOrDelete(next, key, value ?? undefined);
		}
		if (next.trendlineType === 'poly' && next.order === undefined) next.order = 2;
		if (next.trendlineType === 'movingAvg' && next.period === undefined) next.period = 2;
		if (next.order !== undefined && next.trendlineType !== 'poly')
			throw new RangeError('Only polynomial trendlines have an order');
		if (next.period !== undefined && next.trendlineType !== 'movingAvg')
			throw new RangeError('Only moving averages have a period');
		if (edit.name !== undefined) assignOrDelete(next, 'name', edit.name ?? undefined);
		if (edit.spPr !== undefined)
			assignOrDelete(next, 'spPr', structuredClone(edit.spPr ?? undefined));
		if (edit.label !== undefined)
			assignOrDelete(next, 'label', structuredClone(edit.label ?? undefined));
		for (const key of ['displayEquation', 'displayRSquared'] as const)
			if (edit[key] !== undefined) assignOrDelete(next, key, edit[key] ?? undefined);
		list[index] = next;
	});
}

export interface ChartErrorBarsEdit {
	barType?: 'both' | 'minus' | 'plus';
	valueType?: 'cust' | 'fixedVal' | 'percentage' | 'stdDev' | 'stdErr';
	noEndCap?: boolean | null;
	value?: number | null;
	plus?: ChartDataSource | null;
	minus?: ChartDataSource | null;
	spPr?: ChartShapeProperties | null;
}

function numericSource(source: ChartDataSource | undefined): boolean {
	return (
		!!source &&
		['numRef', 'numLit'].includes(source.kind ?? '') &&
		(!source.cache ||
			source.cache.points.every(
				(point) =>
					point.value.trim() !== '' &&
					Number.isFinite(Number(point.value)) &&
					Number(point.value) >= 0,
			))
	);
}

/** Edits the unique x/y error-bar set. Custom sources are explicit numeric sources, never guessed from cells. */
export function setChartErrorBars(
	space: ChartSpace,
	seriesIndex: number,
	direction: ChartErrorBarDirection,
	edit: ChartErrorBarsEdit | null,
): ChartSpace {
	admitted(space, seriesIndex);
	const { group } = chartSeriesEntryAt(space, seriesIndex);
	if (
		!['x', 'y'].includes(direction) ||
		(direction === 'x' && !['scatter', 'bubble'].includes(group.kind))
	)
		throw new RangeError('Only scatter and bubble charts support x error bars');
	return editChartSpace(space, (draft) => {
		const series = chartSeriesEntryAt(draft, seriesIndex).series;
		const list = (series.errorBars ??= []);
		const matches = list.flatMap((item, index) =>
			(item.direction ?? 'y') === direction ? [index] : [],
		);
		if (matches.length > 1) throw new RangeError('Ambiguous error-bar direction');
		const index = matches[0] ?? list.length;
		if (edit === null) {
			if (matches.length) list.splice(index, 1);
			return;
		}
		const next: ChartErrorBars = {
			...(list[index] ?? { direction, barType: 'both', valueType: 'fixedVal', value: 1 }),
		};
		if (edit.barType !== undefined) {
			if (!['both', 'minus', 'plus'].includes(edit.barType))
				throw new RangeError('Invalid error-bar side');
			next.barType = edit.barType;
		}
		if (edit.valueType !== undefined) {
			if (!['cust', 'fixedVal', 'percentage', 'stdDev', 'stdErr'].includes(edit.valueType))
				throw new RangeError('Invalid error-bar amount type');
			next.valueType = edit.valueType;
			if (edit.valueType !== 'cust') {
				delete next.plus;
				delete next.minus;
			}
			if (edit.valueType === 'cust' || edit.valueType === 'stdErr') delete next.value;
			else if (next.value === undefined) next.value = 1;
		}
		if (edit.value !== undefined) {
			if (edit.value !== null && (!Number.isFinite(edit.value) || edit.value < 0))
				throw new RangeError('Error-bar amounts must be nonnegative');
			assignOrDelete(next, 'value', edit.value ?? undefined);
		}
		for (const key of ['plus', 'minus', 'spPr'] as const)
			if (edit[key] !== undefined)
				assignOrDelete(next, key, structuredClone(edit[key] ?? undefined));
		if (edit.noEndCap !== undefined) assignOrDelete(next, 'noEndCap', edit.noEndCap ?? undefined);
		if (next.valueType === 'cust') {
			if (
				(next.barType !== 'minus' && !numericSource(next.plus)) ||
				(next.barType !== 'plus' && !numericSource(next.minus))
			)
				throw new RangeError('Custom error bars need numeric sources for each selected side');
		} else if (edit.plus !== undefined || edit.minus !== undefined)
			throw new RangeError('Only custom error bars accept source ranges');
		if (['cust', 'stdErr'].includes(next.valueType ?? '') && next.value !== undefined)
			throw new RangeError('This error-bar type has no fixed amount');
		list[index] = next;
	});
}
