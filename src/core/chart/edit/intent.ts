// A chart edit as data: what a Format pane (or an assistant, or a collaborator's operation log)
// asks for, one variant per edit operation of this folder. `applyChartEditIntent` runs it over a
// `ChartSpace`; a product whose chart model is not a `ChartSpace` yet (PowerPoint's
// `PptxChartData`, Excel's `ChartObject` summary) maps the variants it supports onto its own edit
// path instead, so every product reads the same request.
import type { DrawingFill, DrawingLine } from '../../drawingml/types';
import type { ChartSpace, ChartView3D, ChartMarker } from '../model';
import type { ChartErrorBarDirection } from '../model-analysis';
import {
	setChartTrendline,
	setChartErrorBars,
	type ChartTrendlineEdit,
	type ChartErrorBarsEdit,
} from './analysis';
import { setChartMarker } from './marker';
import { setChartTargetLineOptions, type ChartLineEdit } from './line-options';
import { setChartTargetPattern, type ChartPatternEdit } from './pattern';
import { setChartTargetEffect, type ChartEffectEdit } from './effects';
import { setChartTargetText, type ChartTextEdit } from './text';
import { setChartTargetThreeD, setChartView3D, type ChartThreeDEdit } from './three-d';
import { setChartSeriesSecondaryAxis } from './secondary-axis';
import { setChartType, type ChartTypeTarget } from './type';
import { setChartAxis, type ChartAxisEdit } from './axes';
import { setChartGroupOptions, type ChartGroupOptionsEdit } from './group-options';
import {
	setChartDataLabels,
	setChartLegend,
	type ChartDataLabelsEdit,
	type ChartLegendEdit,
} from './legend-labels';
import { applyChartQuickLayout } from './quick-layout';
import { applyChartColorPalette, applyChartStylePreset, setChartStyle } from './style';
import type { ChartFormatTarget } from './target';
import { setChartTargetFill, setChartTargetLine, setChartTargetTransparency } from './target-paint';

export type ChartEditIntent =
	| { op: 'pattern'; target: ChartFormatTarget; edit: ChartPatternEdit }
	| { op: 'lineOptions'; target: ChartFormatTarget; edit: ChartLineEdit }
	| { op: 'type'; type: ChartTypeTarget }
	| { op: 'marker'; seriesIndex: number; edit: Partial<ChartMarker> | null }
	| { op: 'trendline'; seriesIndex: number; index: number; edit: ChartTrendlineEdit | null }
	| {
			op: 'errorBars';
			seriesIndex: number;
			direction: ChartErrorBarDirection;
			edit: ChartErrorBarsEdit | null;
	  }
	| { op: 'secondaryAxis'; seriesIndex: number; secondary: boolean }
	| { op: 'effect'; target: ChartFormatTarget; edit: ChartEffectEdit }
	| { op: 'text'; target: ChartFormatTarget; edit: ChartTextEdit }
	| { op: 'threeD'; target: ChartFormatTarget; edit: ChartThreeDEdit }
	| { op: 'view3D'; edit: Partial<ChartView3D> }
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
	| { op: 'groupOptions'; seriesIndex: number; edit: ChartGroupOptionsEdit }
	/** A built-in chart style: 1..48 (`c:style`) or 101..148 (`c14:style`). */
	| { op: 'style'; id: number }
	/** A curated Chart Styles preset (`chart/style-presets`). */
	| { op: 'stylePreset'; preset: string }
	/** Change Colors: recolour from palette `from` to palette `to` (the caller records `to`). */
	| { op: 'colorPalette'; from: number; to: number }
	/** Quick Layout `layout` (1..11). */
	| { op: 'quickLayout'; layout: number };

/** Applies `intent`, returning a new model; throws like the operation it names. */
export function applyChartEditIntent(space: ChartSpace, intent: ChartEditIntent): ChartSpace {
	switch (intent.op) {
		case 'pattern':
			return setChartTargetPattern(space, intent.target, intent.edit);
		case 'lineOptions':
			return setChartTargetLineOptions(space, intent.target, intent.edit);
		case 'type':
			return setChartType(space, intent.type);
		case 'marker':
			return setChartMarker(space, intent.seriesIndex, intent.edit);
		case 'trendline':
			return setChartTrendline(space, intent.seriesIndex, intent.index, intent.edit);
		case 'errorBars':
			return setChartErrorBars(space, intent.seriesIndex, intent.direction, intent.edit);
		case 'secondaryAxis':
			return setChartSeriesSecondaryAxis(space, intent.seriesIndex, intent.secondary);
		case 'effect':
			return setChartTargetEffect(space, intent.target, intent.edit);
		case 'text':
			return setChartTargetText(space, intent.target, intent.edit);
		case 'threeD':
			return setChartTargetThreeD(space, intent.target, intent.edit);
		case 'view3D':
			return setChartView3D(space, intent.edit);
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
		case 'style':
			return setChartStyle(space, intent.id);
		case 'stylePreset':
			return applyChartStylePreset(space, intent.preset);
		case 'colorPalette':
			return applyChartColorPalette(space, intent.from, intent.to);
		case 'quickLayout':
			return applyChartQuickLayout(space, intent.layout);
	}
}
