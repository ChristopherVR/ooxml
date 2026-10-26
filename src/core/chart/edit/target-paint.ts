// Fill and outline edits addressed by a Format pane target (`target.ts`), so one call serves every
// element a pane can select: chart and plot area, title and legend (`fill.ts`), a series, an axis
// line, an axis's gridlines and the chart's or one series' data labels. Series fills keep Excel's
// semantics from `fill.ts` (a series fill replaces its points' own fills); gridline styling turns
// the gridlines on, as `setChartGridlineShape` does.
import type { DrawingFill, DrawingLine } from '../../drawingml/types';
import type { ChartSpace } from '../model';
import type { ChartShapeProperties } from '../model-series';
import { chartAxisAt } from './axes';
import { assignOrDelete, chartSeriesEntryAt, editChartSpace } from './draft';
import {
	chartElementShape,
	setChartElementFill,
	setChartSeriesFill,
	type ChartElementPart,
} from './fill';
import { withDrawingColorTransparency } from './paint';
import type { ChartFormatTarget } from './target';

const ELEMENT: readonly string[] = ['chartArea', 'plotArea', 'title', 'legend'];
const isElement = (
	target: ChartFormatTarget,
): target is Extract<ChartFormatTarget, { kind: ChartElementPart }> =>
	ELEMENT.includes(target.kind);

/** The shape properties a target declares, if any (an inherited style is not seen here). */
export function chartTargetShape(
	space: ChartSpace,
	target: ChartFormatTarget,
): ChartShapeProperties | undefined {
	if (isElement(target)) return chartElementShape(space, target.kind);
	switch (target.kind) {
		case 'series':
			return chartSeriesEntryAt(space, target.seriesIndex).series.spPr;
		case 'axis':
			return chartAxisAt(space, { index: target.axisIndex }).spPr;
		case 'majorGridlines':
			return chartAxisAt(space, { index: target.axisIndex }).majorGridlinesSpPr;
		case 'minorGridlines':
			return chartAxisAt(space, { index: target.axisIndex }).minorGridlinesSpPr;
		case 'dataLabels':
			return target.seriesIndex === undefined
				? space.plotArea.groups.find((group) => group.dataLabels)?.dataLabels?.spPr
				: chartSeriesEntryAt(space, target.seriesIndex).series.dataLabels?.spPr;
		default:
			return undefined;
	}
}

/** Every shape a target writes on the draft, created when missing; throws for a missing element. */
export function chartTargetDraftShapes(
	draft: ChartSpace,
	target: ChartFormatTarget,
): ChartShapeProperties[] {
	switch (target.kind) {
		case 'chartArea':
			return [(draft.spPr ??= {})];
		case 'plotArea':
			return [(draft.plotArea.spPr ??= {})];
		case 'title':
			if (!draft.title) throw new RangeError('The chart has no title');
			return [(draft.title.spPr ??= {})];
		case 'legend':
			if (!draft.legend) throw new RangeError('The chart has no legend');
			return [(draft.legend.spPr ??= {})];
		case 'series':
			return [(chartSeriesEntryAt(draft, target.seriesIndex).series.spPr ??= {})];
		case 'axis':
			return [(chartAxisAt(draft, { index: target.axisIndex }).spPr ??= {})];
		case 'majorGridlines':
		case 'minorGridlines': {
			const axis = chartAxisAt(draft, { index: target.axisIndex });
			const major = target.kind === 'majorGridlines';
			axis[major ? 'majorGridlines' : 'minorGridlines'] = true;
			return [(axis[major ? 'majorGridlinesSpPr' : 'minorGridlinesSpPr'] ??= {})];
		}
		case 'dataLabels': {
			if (target.seriesIndex === undefined) {
				for (const group of draft.plotArea.groups)
					if (group.kind !== 'surface') group.dataLabels ??= { labels: [] };
			} else {
				const entry = chartSeriesEntryAt(draft, target.seriesIndex);
				if (entry.group.kind === 'surface')
					throw new RangeError('Surface charts have no data labels');
				entry.series.dataLabels ??= { labels: [] };
			}
			const labels =
				target.seriesIndex === undefined
					? draft.plotArea.groups.flatMap((group) => (group.dataLabels ? [group.dataLabels] : []))
					: [chartSeriesEntryAt(draft, target.seriesIndex).series.dataLabels].filter(
							(item) => item !== undefined,
						);
			if (labels.length === 0) throw new RangeError('The chart shows no data labels there');
			return labels.map((item) => (item.spPr ??= {}));
		}
	}
}

/**
 * Sets (or, with `undefined`, removes so the chart style applies) the fill of a target. Axes and
 * gridlines are lines: their fill is refused.
 */
export function setChartTargetFill(
	space: ChartSpace,
	target: ChartFormatTarget,
	fill: DrawingFill | undefined,
): ChartSpace {
	if (target.kind === 'dataLabels' && fill === undefined) {
		const labels =
			target.seriesIndex === undefined
				? space.plotArea.groups.some((group) => group.dataLabels)
				: chartSeriesEntryAt(space, target.seriesIndex).series.dataLabels;
		if (!labels) throw new RangeError('The chart shows no data labels there');
	}
	const current = chartTargetShape(space, target)?.fill;
	const next =
		fill &&
		current?.kind === fill.kind &&
		'sourceXml' in current &&
		current.sourceXml &&
		(fill.kind === 'pattern' || fill.kind === 'picture' || fill.kind === 'gradient') &&
		!fill.sourceXml
			? { ...fill, sourceXml: current.sourceXml }
			: fill;
	if (isElement(target)) return setChartElementFill(space, target.kind, next);
	if (target.kind === 'series') return setChartSeriesFill(space, target.seriesIndex, next);
	if (target.kind !== 'dataLabels') throw new RangeError(`A chart ${target.kind} has no fill`);
	return editChartSpace(space, (draft) => {
		for (const shape of chartTargetDraftShapes(draft, target))
			assignOrDelete(shape, 'fill', structuredClone(next));
	});
}

/**
 * Sets a target's outline: replaced by `line` (`undefined` removes it so the style applies), or,
 * with `merge`, merged into the current outline (width, dash or cap alone).
 */
export function setChartTargetLine(
	space: ChartSpace,
	target: ChartFormatTarget,
	line: DrawingLine | undefined,
	merge = false,
): ChartSpace {
	return editChartSpace(space, (draft) => {
		for (const shape of chartTargetDraftShapes(draft, target)) {
			const next =
				merge && line ? { ...shape.line, ...structuredClone(line) } : structuredClone(line);
			assignOrDelete(shape, 'line', next);
		}
		if (target.kind === 'series' && !merge)
			for (const point of chartSeriesEntryAt(draft, target.seriesIndex).series.dataPoints)
				if (point.spPr) delete point.spPr.line;
	});
}

/** Sets the transparency (0..100) of a target's solid fill or solid outline; other paints stay. */
export function setChartTargetTransparency(
	space: ChartSpace,
	target: ChartFormatTarget,
	which: 'fill' | 'line',
	percent: number,
): ChartSpace {
	return editChartSpace(space, (draft) => {
		for (const shape of chartTargetDraftShapes(draft, target)) {
			const holder = which === 'fill' ? shape : shape.line;
			if (holder?.fill?.kind !== 'solid') continue;
			holder.fill = {
				kind: 'solid',
				color: withDrawingColorTransparency(holder.fill.color, percent),
			};
		}
	});
}
