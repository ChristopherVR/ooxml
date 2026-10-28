import { boxWhiskerDistribution, boxMeanMarkerPath } from 'ooxml-core/chart';
import type { PptxChartBoxWhiskerOptions, PptxChartData, PptxElement } from 'ooxml-core/pptx';

import { resolveChartTitleText } from './chart-auto-title';
import { groupRowsByCategory } from './chart-box-whisker-stats';
import type { BoxStats } from './chart-box-whisker-stats';
import { buildValueAxisGridlinesAndLabels, findValueAxis } from './chart-cx-axis-units';
import { distributionRange } from './chart-distribution-range';
import type {
	ChartViewModel,
	PlotLayout,
	SvgCircle,
	SvgLine,
	SvgPrimitive,
	SvgRect,
	SvgPath,
	ValueRange,
} from './chart-view-model';
import {
	buildCategoryLabels,
	buildLegend,
	buildZeroLine,
	computePlotLayout,
	seriesColor,
	valueToY,
} from './chart-view-model';

const WHISKER_COLOR = '#64748b';
const MEDIAN_COLOR = '#1e293b';

interface BoxPoint {
	x: number;
	y: number;
	seriesIndex: number;
	/** The raw source row this observation came from, its stable point identity. */
	rowIndex: number;
	outlier: boolean;
}

export interface BoxWhiskerGeometry {
	stats: BoxStats;
	boxX: number;
	boxW: number;
	xMid: number;
	yMin: number;
	yMax: number;
	yQ1: number;
	yQ3: number;
	yMed: number;
	yMean: number;
	fill: string;
	seriesIndex: number;
	categoryIndex: number;
	options: PptxChartBoxWhiskerOptions | undefined;
	points: BoxPoint[];
}

/**
 * Build one box per (series, category) pair, the way PowerPoint's
 * `cx:boxWhisker` chart actually groups data: a series' raw rows repeat their
 * category label once per underlying observation (see
 * `groupRowsByCategory`'s doc comment), and each series draws its OWN box in
 * every category it has observations for, side by side with the other
 * series' boxes in that same category (COM-verified: `charts-com.pptx` slide
 * 32 draws 3 colour-coded boxes per category, not one box mixing all series).
 */
export function computeBoxWhiskerGeometry(
	chartData: PptxChartData,
	rawCategories: ReadonlyArray<string>,
	layout: PlotLayout,
	range: ValueRange,
	colorPalette: readonly string[] | undefined,
): BoxWhiskerGeometry[] {
	const { uniqueCategories, rowIndexesByCategory } = groupRowsByCategory(rawCategories);
	const catCount = Math.max(uniqueCategories.length, 1);
	const seriesCount = Math.max(chartData.series.length, 1);
	const groupWidth = layout.plotWidth / catCount;
	const slotWidth = groupWidth / seriesCount;
	const boxW = slotWidth * 0.7;

	const output: BoxWhiskerGeometry[] = [];
	uniqueCategories.forEach((category, categoryIndex) => {
		const rows = rowIndexesByCategory.get(category) ?? [];
		chartData.series.forEach((series, seriesIndex) => {
			const observations = rows
				.map((row) => ({ row, value: series.values[row] }))
				.filter((item): item is { row: number; value: number } => item.value !== undefined);
			const values = observations.map((item) => item.value);
			const options = series.boxWhiskerOptions;
			const distribution = boxWhiskerDistribution(
				values,
				options?.quartileMethod ?? 'exclusive',
				options !== undefined,
			);
			if (!distribution) return;
			const { stats, whiskerMin, whiskerMax, mean } = distribution;
			const boxX =
				layout.plotLeft +
				groupWidth * categoryIndex +
				slotWidth * seriesIndex +
				(slotWidth - boxW) / 2;
			output.push({
				stats,
				boxX,
				boxW,
				xMid: boxX + boxW / 2,
				yMin: valueToY(whiskerMin, range, layout.plotTop, layout.plotBottom),
				yMax: valueToY(whiskerMax, range, layout.plotTop, layout.plotBottom),
				yQ1: valueToY(stats.q1, range, layout.plotTop, layout.plotBottom),
				yQ3: valueToY(stats.q3, range, layout.plotTop, layout.plotBottom),
				yMed: valueToY(stats.median, range, layout.plotTop, layout.plotBottom),
				yMean: valueToY(mean, range, layout.plotTop, layout.plotBottom),
				fill: seriesColor(series, seriesIndex, colorPalette),
				seriesIndex,
				categoryIndex,
				options,
				points: observations.map((item, index) => ({
					x: boxX + boxW * (0.2 + (0.6 * (index + 1)) / (observations.length + 1)),
					y: valueToY(item.value, range, layout.plotTop, layout.plotBottom),
					seriesIndex,
					rowIndex: item.row,
					outlier: item.value < whiskerMin || item.value > whiskerMax,
				})),
			});
		});
	});
	return output;
}

function whiskerPrimitives(geometry: BoxWhiskerGeometry): SvgPrimitive[] {
	const g = geometry;
	return [
		{
			kind: 'line',
			x1: g.xMid,
			y1: g.yMax,
			x2: g.xMid,
			y2: g.yQ3,
			stroke: WHISKER_COLOR,
			strokeWidth: 1,
		},
		{
			kind: 'line',
			x1: g.xMid,
			y1: g.yQ1,
			x2: g.xMid,
			y2: g.yMin,
			stroke: WHISKER_COLOR,
			strokeWidth: 1,
		},
		{
			kind: 'line',
			x1: g.boxX + g.boxW * 0.25,
			y1: g.yMax,
			x2: g.boxX + g.boxW * 0.75,
			y2: g.yMax,
			stroke: WHISKER_COLOR,
			strokeWidth: 1,
		},
		{
			kind: 'line',
			x1: g.boxX + g.boxW * 0.25,
			y1: g.yMin,
			x2: g.boxX + g.boxW * 0.75,
			y2: g.yMin,
			stroke: WHISKER_COLOR,
			strokeWidth: 1,
		},
		{
			kind: 'rect',
			x: g.boxX,
			y: Math.min(g.yQ1, g.yQ3),
			w: g.boxW,
			h: Math.abs(g.yQ1 - g.yQ3),
			fill: g.fill,
			rx: 1,
			opacity: 0.8,
			part: { role: 'dataPoint', seriesIndex: g.seriesIndex, pointIndex: g.categoryIndex },
		},
		{
			kind: 'line',
			x1: g.boxX,
			y1: g.yMed,
			x2: g.boxX + g.boxW,
			y2: g.yMed,
			stroke: MEDIAN_COLOR,
			strokeWidth: 2,
		},
	] satisfies Array<SvgLine | SvgRect>;
}

function optionPrimitives(geometry: BoxWhiskerGeometry): SvgPrimitive[] {
	const options = geometry.options;
	const output: SvgPrimitive[] = [];
	if (options?.showMeanLine) {
		output.push({
			kind: 'line',
			x1: geometry.boxX,
			y1: geometry.yMean,
			x2: geometry.boxX + geometry.boxW,
			y2: geometry.yMean,
			stroke: '#0f766e',
			strokeWidth: 1.5,
		} satisfies SvgLine);
	}
	if (options?.showMeanMarker !== false) {
		output.push({
			kind: 'path',
			d: boxMeanMarkerPath(geometry.xMid, geometry.yMean),
			fill: 'none',
			stroke: '#0f766e',
			strokeWidth: 1.5,
		} satisfies SvgPath);
	}
	for (const point of geometry.points) {
		if (
			(point.outlier && !options?.showOutlierPoints) ||
			(!point.outlier && !options?.showInnerPoints)
		) {
			continue;
		}
		output.push({
			kind: 'circle',
			cx: point.x,
			cy: point.y,
			r: 2.25,
			fill: point.outlier ? '#dc2626' : MEDIAN_COLOR,
			part: { role: 'dataPoint', seriesIndex: point.seriesIndex, pointIndex: point.rowIndex },
		} satisfies SvgCircle);
	}
	return output;
}

export function buildBoxWhiskerViewModel(
	element: PptxElement,
	chartData: PptxChartData,
	categoryLabels: ReadonlyArray<string>,
): ChartViewModel {
	const layout = computePlotLayout(element.width, element.height, chartData, true);
	const range = distributionRange(chartData.series);
	const geometries = computeBoxWhiskerGeometry(
		chartData,
		categoryLabels,
		layout,
		range,
		chartData.colorPalette,
	);
	const primitives = geometries.flatMap((geometry) => [
		...whiskerPrimitives(geometry),
		...optionPrimitives(geometry),
	]);
	const { gridlines, axisLabels } = buildValueAxisGridlinesAndLabels(
		range,
		layout,
		findValueAxis(chartData.axes),
	);
	const { legend, legendX, legendY, legendAnchor } = buildLegend(
		chartData.series,
		chartData.colorPalette,
		layout.svgWidth,
		chartData.style?.legendPosition ?? 'b',
		layout.svgHeight,
		layout.plotTop,
	);
	const { uniqueCategories } = groupRowsByCategory(categoryLabels);
	return {
		svgWidth: layout.svgWidth,
		svgHeight: layout.svgHeight,
		title: resolveChartTitleText(chartData),
		titleX: layout.svgWidth / 2,
		titleY: 12,
		gridlines,
		axisLabels,
		zeroLine: buildZeroLine(range, layout),
		categoryLabels: buildCategoryLabels(uniqueCategories, layout, 'bar'),
		primitives,
		dataLabels: [],
		legend: chartData.style?.hasLegend ? legend : [],
		legendX,
		legendY,
		legendAnchor,
	};
}
