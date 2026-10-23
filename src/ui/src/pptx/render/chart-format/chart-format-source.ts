/**
 * PowerPoint's source for the shared chart Format pane: what a
 * `PptxChartData` reads as (`ChartFormatSource` of `src/ui/src/chart`). The
 * element view is the core's `readChartElementFormat`, so the pane keeps every
 * PowerPoint section (pattern and picture fills, effects, 3-D Format, Text
 * Options, the secondary axis) as extension sections the neutral source lacks.
 *
 * @module render/chart-format/chart-format-source
 */
import type { PptxChartData } from 'ooxml-core/pptx';
import {
	chartSupportsSecondaryAxis,
	isChartSeriesOnSecondaryAxis,
	readChartElementFormat,
} from 'ooxml-core/pptx';

import type {
	ChartFormatAxisView,
	ChartFormatLabelsView,
	ChartFormatSource,
} from '../../../chart/format-pane-view';
import { PATTERN_PRESET_OPTIONS } from '../fill-pattern-label-keys';
import { listChartFormatTargets, resolveUnlistedChartFormatTarget } from './chart-format-targets';

const AXIS_KINDS = { catAx: 'cat', valAx: 'val', dateAx: 'date', serAx: 'ser' } as const;

const PATTERN_PRESETS = PATTERN_PRESET_OPTIONS.map((option) => ({
	value: option.value,
	label: { key: option.labelKey, fallback: option.value },
}));

/** The Format pane source of `data`. */
export function pptxChartFormatSource(data: PptxChartData): ChartFormatSource {
	return {
		targets: listChartFormatTargets(data),
		resolve: (key) => resolveUnlistedChartFormatTarget(data, key),
		element: (target) => readChartElementFormat(data, target),
		patternPresets: PATTERN_PRESETS,
		seriesOptions: (seriesIndex) => ({
			chartType: data.chartType === 'combo' ? undefined : data.chartType,
			secondaryAxis: chartSupportsSecondaryAxis(data)
				? isChartSeriesOnSecondaryAxis(data, seriesIndex)
				: undefined,
			overlap: data.barOverlap,
			gapWidth: data.barGapWidth,
			firstSliceAngle: data.firstSliceAngle,
			holeSize: data.doughnutHoleSize,
		}),
		axisOptions: (axisIndex): ChartFormatAxisView | undefined => {
			const axis = data.axes?.[axisIndex];
			if (!axis) return undefined;
			return {
				kind: AXIS_KINDS[axis.axisType] ?? 'cat',
				min: axis.min,
				max: axis.max,
				majorUnit: axis.majorUnit,
				minorUnit: axis.minorUnit,
				crosses: axis.crosses,
				crossesAt: axis.crossesAt,
				displayUnits:
					axis.displayUnits && axis.displayUnits !== 'custom' ? axis.displayUnits : undefined,
				logScale: axis.logScale,
				logBase: axis.logBase,
				reverse: axis.orientation === 'maxMin',
				majorTickMark: axis.majorTickMark,
				minorTickMark: axis.minorTickMark,
				tickLabelPosition: axis.tickLblPos,
				numberFormat: axis.numFmt && !axis.numFmt.sourceLinked ? axis.numFmt.formatCode : '',
			};
		},
		legendOptions: () => ({
			position: data.style?.legendPosition,
			overlay: data.style?.legendOverlay,
		}),
		labelOptions: (seriesIndex): ChartFormatLabelsView => {
			const series = seriesIndex === undefined ? undefined : data.series[seriesIndex];
			const chartLevel = data.style?.dataLabels ?? {};
			// One series' labels: its own options, each flag falling back to the chart-level one.
			const labels = series
				? {
						...chartLevel,
						...Object.fromEntries(
							Object.entries(series.dataLabelOptions ?? {}).filter(([, v]) => v !== undefined),
						),
					}
				: chartLevel;
			return {
				chartType: data.chartType,
				seriesScoped: series !== undefined,
				showSeriesName: labels.showSeriesName,
				showCategory: labels.showCategory,
				showValue: labels.showValue,
				showPercent: labels.showPercent,
				showLegendKey: labels.showLegendKey,
				position: labels.position,
			};
		},
	};
}
