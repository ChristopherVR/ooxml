// A chart edit as data: what a Format pane (or an assistant, or a collaborator's operation log)
// asks for, one variant per edit operation of this folder. `applyChartEditIntent` runs it over a
// `ChartSpace`; a product whose chart model is not a `ChartSpace` yet (PowerPoint's
// `PptxChartData`, Excel's `ChartObject` summary) maps the variants it supports onto its own edit
// path instead, so every product reads the same request.
import type { DrawingFill, DrawingLine } from '../../drawingml/types';
import type { ChartSpace } from '../model';
import { setChartAxis, type ChartAxisEdit } from './axes';
import { setChartGroupOptions, type ChartGroupOptionsEdit } from './group-options';
import {
	setChartDataLabels,
	setChartLegend,
	type ChartDataLabelsEdit,
	type ChartLegendEdit,
} from './legend-labels';
import type { ChartFormatTarget } from './target';
import { setChartTargetFill, setChartTargetLine, setChartTargetTransparency } from './target-paint';

export type ChartEditIntent =
	/** Replace a target's fill; `undefined` returns it to the chart style ("Automatic"). */
	| { op: 'fill'; target: ChartFormatTarget; fill: DrawingFill | undefined }
	/** Replace a target's outline, or with `merge` change only the given outline settings. */
	| { op: 'line'; target: ChartFormatTarget; line: DrawingLine | undefined; merge?: boolean }
	/** Transparency (0..100) of a target's solid fill or solid outline, keeping its colour. */
	| { op: 'transparency'; target: ChartFormatTarget; paint: 'fill' | 'line'; percent: number }
	/** Axis options of the axis at `axisIndex` of `plotArea.axes`. */
	| { op: 'axis'; axisIndex: number; edit: ChartAxisEdit }
	| { op: 'legend'; edit: ChartLegendEdit }
	/** The chart's data labels, or one series' own labels. */
	| { op: 'dataLabels'; seriesIndex?: number; edit: ChartDataLabelsEdit }
	/** Series Options of the group holding series `seriesIndex`. */
	| { op: 'groupOptions'; seriesIndex: number; edit: ChartGroupOptionsEdit };

/** Applies `intent`, returning a new model; throws like the operation it names. */
export function applyChartEditIntent(space: ChartSpace, intent: ChartEditIntent): ChartSpace {
	switch (intent.op) {
		case 'fill':
			return setChartTargetFill(space, intent.target, intent.fill);
		case 'line':
			return setChartTargetLine(space, intent.target, intent.line, intent.merge);
		case 'transparency':
			return setChartTargetTransparency(space, intent.target, intent.paint, intent.percent);
		case 'axis':
			return setChartAxis(space, { index: intent.axisIndex }, intent.edit);
		case 'legend':
			return setChartLegend(space, intent.edit);
		case 'dataLabels':
			return setChartDataLabels(
				space,
				intent.edit,
				intent.seriesIndex === undefined ? {} : { series: intent.seriesIndex },
			);
		case 'groupOptions':
			return setChartGroupOptions(space, intent.seriesIndex, intent.edit);
	}
}
