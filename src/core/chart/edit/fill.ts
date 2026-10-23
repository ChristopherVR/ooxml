// Fill, outline, gradient and transparency edits of series, data points and chart elements, over
// their shape properties (`c:spPr`). The writer patches a changed fill or line into the kept
// `c:spPr` source, so geometry, effects and 3-D settings survive. Series colour follows Excel's
// Format pane (xlsx `edit/chart-series-fill.ts`): a new colour keeps the old transparency and
// replaces the points' own fills; a line, scatter or radar series is coloured by its line.
import type { DrawingColor, DrawingFill, DrawingLine } from '../../drawingml/types';
import type { ChartSpace } from '../model';
import type { ChartShapeProperties, ChartSpaceSeries } from '../model-series';
import { assignOrDelete, chartSeriesEntryAt, editChartSpace } from './draft';
import { editChartGradientFill, type ChartGradientFillEdit } from './gradient';
import { withAlphaOf, withDrawingColorTransparency } from './paint';

/** Chart elements with their own shape properties. */
export const CHART_ELEMENT_PARTS = [
	'chartArea',
	'plotArea',
	'title',
	'legend',
	'floor',
	'sideWall',
	'backWall',
] as const;
export type ChartElementPart = (typeof CHART_ELEMENT_PARTS)[number];

/** The shape properties of an element on the draft, created when missing. */
function elementShape(draft: ChartSpace, part: ChartElementPart): ChartShapeProperties | undefined {
	switch (part) {
		case 'chartArea':
			return (draft.spPr ??= {});
		case 'plotArea':
			return (draft.plotArea.spPr ??= {});
		case 'title':
			if (!draft.title) throw new RangeError('The chart has no title');
			return (draft.title.spPr ??= {});
		case 'legend':
			if (!draft.legend) throw new RangeError('The chart has no legend');
			return (draft.legend.spPr ??= {});
		default:
			return ((draft[part] ??= {}).spPr ??= {});
	}
}

/** The fill an element declares, if any (an inherited style fill is not seen here). */
export function chartElementShape(
	space: ChartSpace,
	part: ChartElementPart,
): ChartShapeProperties | undefined {
	switch (part) {
		case 'chartArea':
			return space.spPr;
		case 'plotArea':
			return space.plotArea.spPr;
		case 'title':
			return space.title?.spPr;
		case 'legend':
			return space.legend?.spPr;
		default:
			return space[part]?.spPr;
	}
}

/** Sets (or, with `undefined`, removes so the style applies) an element's fill. */
export function setChartElementFill(
	space: ChartSpace,
	part: ChartElementPart,
	fill: DrawingFill | undefined,
): ChartSpace {
	return editChartSpace(space, (draft) =>
		assignOrDelete(elementShape(draft, part)!, 'fill', structuredClone(fill)),
	);
}

/** Sets (or removes) an element's outline. */
export function setChartElementLine(
	space: ChartSpace,
	part: ChartElementPart,
	line: DrawingLine | undefined,
): ChartSpace {
	return editChartSpace(space, (draft) =>
		assignOrDelete(elementShape(draft, part)!, 'line', structuredClone(line)),
	);
}

/** Whether a series is coloured by its line (line, stock, radar other than filled, scatter). */
function paintsLine(space: ChartSpace, index: number): boolean {
	const { group } = chartSeriesEntryAt(space, index);
	if (group.kind === 'radar') return group.radarStyle !== 'filled';
	return group.kind === 'line' || group.kind === 'stock' || group.kind === 'scatter';
}

function clearPointPaint(series: ChartSpaceSeries, key: 'fill' | 'line'): void {
	for (const point of series.dataPoints) if (point.spPr) delete point.spPr[key];
}

/** Sets a series' fill (`undefined` returns it to the automatic colour); points lose their own. */
export function setChartSeriesFill(
	space: ChartSpace,
	index: number,
	fill: DrawingFill | undefined,
): ChartSpace {
	return editChartSpace(space, (draft) => {
		const { series } = chartSeriesEntryAt(draft, index);
		assignOrDelete((series.spPr ??= {}), 'fill', structuredClone(fill));
		clearPointPaint(series, 'fill');
	});
}

/** Sets a series' outline, merged into the current one (`null` removes it). */
export function setChartSeriesLine(
	space: ChartSpace,
	index: number,
	line: DrawingLine | null,
): ChartSpace {
	return editChartSpace(space, (draft) => {
		const shape = (chartSeriesEntryAt(draft, index).series.spPr ??= {});
		assignOrDelete(
			shape,
			'line',
			line === null ? undefined : { ...shape.line, ...structuredClone(line) },
		);
	});
}

/**
 * Sets the colour of a series as the Format pane's colour picker does: the fill, or the line of a
 * series drawn as a line, keeping the previous colour's transparency. `null` removes the paint so
 * the series takes its automatic colour again.
 */
export function setChartSeriesColor(
	space: ChartSpace,
	index: number,
	color: DrawingColor | null,
): ChartSpace {
	return editChartSpace(space, (draft) => {
		const { series } = chartSeriesEntryAt(draft, index);
		const shape = (series.spPr ??= {});
		const line = paintsLine(draft, index);
		const current = line ? shape.line?.fill : shape.fill;
		const previous = current?.kind === 'solid' ? current.color : undefined;
		const fill: DrawingFill | undefined = color
			? { kind: 'solid', color: withAlphaOf(color, previous) }
			: undefined;
		if (line) assignOrDelete((shape.line ??= {}), 'fill', fill);
		else assignOrDelete(shape, 'fill', fill);
		clearPointPaint(series, line ? 'line' : 'fill');
	});
}

/** Sets a solid series fill's transparency (0..100); other fills are left alone. */
export function setChartSeriesTransparency(
	space: ChartSpace,
	index: number,
	percent: number,
): ChartSpace {
	return editChartSpace(space, (draft) => {
		const shape = chartSeriesEntryAt(draft, index).series.spPr;
		const target = paintsLine(draft, index) ? shape?.line : shape;
		if (target?.fill?.kind !== 'solid') return;
		target.fill = {
			kind: 'solid',
			color: withDrawingColorTransparency(target.fill.color, percent),
		};
	});
}

/** Sets (or removes, with `undefined`) one data point's fill, through its `c:dPt` override. */
export function setChartDataPointFill(
	space: ChartSpace,
	index: number,
	point: number,
	fill: DrawingFill | undefined,
): ChartSpace {
	if (!Number.isInteger(point) || point < 0)
		throw new RangeError(`No chart point at index ${String(point)}`);
	return editChartSpace(space, (draft) => {
		const { series } = chartSeriesEntryAt(draft, index);
		let override = series.dataPoints.find((item) => item.index === point);
		if (!override) {
			if (!fill) return;
			override = { index: point };
			series.dataPoints.push(override);
			series.dataPoints.sort((a, b) => (a.index ?? 0) - (b.index ?? 0));
		}
		assignOrDelete((override.spPr ??= {}), 'fill', structuredClone(fill));
		const spPr = override.spPr;
		const empty = !spPr.fill && !spPr.line && !spPr.effectsXml && !spPr.sourceXml;
		if (empty && Object.keys(override).every((key) => key === 'index' || key === 'spPr'))
			series.dataPoints = series.dataPoints.filter((item) => item !== override);
	});
}

/** A gradient edit of a series fill; `undefined` when it does not apply or changes nothing. */
export function editChartSeriesGradient(
	space: ChartSpace,
	index: number,
	edit: ChartGradientFillEdit,
	defaultColor: DrawingColor,
): { space: ChartSpace; stopIndex: number } | undefined {
	const current = chartSeriesEntryAt(space, index).series.spPr?.fill;
	const base = current?.kind === 'solid' ? current.color : defaultColor;
	const result = editChartGradientFill(current, edit, base);
	return (
		result && { space: setChartSeriesFill(space, index, result.fill), stopIndex: result.stopIndex }
	);
}

/** A gradient edit of an element fill (a new gradient starts from its solid colour, else white). */
export function editChartElementGradient(
	space: ChartSpace,
	part: ChartElementPart,
	edit: ChartGradientFillEdit,
): { space: ChartSpace; stopIndex: number } | undefined {
	const current = chartElementShape(space, part)?.fill;
	const base: DrawingColor =
		current?.kind === 'solid' ? current.color : { kind: 'scheme', value: 'lt1', transforms: [] };
	const result = editChartGradientFill(current, edit, base);
	return (
		result && { space: setChartElementFill(space, part, result.fill), stopIndex: result.stopIndex }
	);
}
