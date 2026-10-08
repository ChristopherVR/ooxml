/**
 * chart-cartesian-area.ts: area-chart plot-primitive builder, split out of
 * `chart-cartesian-line-area.ts` (which still owns `buildLines`) to keep each
 * module within the repo's ~300-LOC limit.
 *
 * @module chart-cartesian-area
 */
import type { PptxChartData } from 'ooxml-core/pptx';

import { pushMarker } from './chart-cartesian-plots';
import type { SeriesPlotResult } from './chart-cartesian-plots';
import { resolveMarkerLabelPlacement } from './chart-data-label-anchor';
import { pushPointLabel } from './chart-data-label-callout';
import {
	buildDataLabelText,
	dataLabelFontOverride,
	resolveDataLabelTextStyle,
} from './chart-data-label-text';
import { resolveDataPointFill } from './chart-datapoint-style';
import { DEFAULT_CHART_DATA_LABEL_PX } from './chart-font';
import {
	clipPolygonToBand,
	clippedSeriesLine,
	insideBand,
	plotBand,
	visibleSpanCentre,
} from './chart-plot-clip';
import { computeStackedSeriesPlots } from './chart-stacked-series';
import type { LineAreaStacking } from './chart-stacked-series';
import type {
	ChartPartRef,
	PlotLayout,
	SvgPolyline,
	SvgPrimitive,
	SvgText,
	ValueRange,
} from './chart-view-model';
import {
	buildMarkTooltip,
	computeLinePoints,
	linePointsToSvgString,
	seriesColor,
	valueToY,
} from './chart-view-model';

/**
 * Build area-chart primitives (fill polygon + outline).
 *
 * Stacked/percentStacked fills the band between this series' running-sum top
 * and the previous series' top (its own base) instead of down to the zero
 * baseline, so layers stack visually like a PowerPoint stacked area chart
 * rather than each series washing over the ones below it. Data labels and
 * marker tooltips read the series' own value (or percent share), matching
 * `buildLines` and stacked bar's label convention.
 */
export function buildAreas(
	chartData: PptxChartData,
	catCount: number,
	layout: PlotLayout,
	range: ValueRange,
	sourceIndices: ReadonlyArray<number>,
	xPositions?: ReadonlyArray<number>,
	stacking: LineAreaStacking = 'clustered',
): SeriesPlotResult {
	const primitives: SvgPrimitive[] = [],
		dataLabels: SvgText[] = [],
		labelBoxes: SvgPrimitive[] = [],
		showLabels = chartData.style?.hasDataLabels,
		isStackedMode = stacking !== 'clustered',
		isPercent = stacking === 'percentStacked',
		baselineY = valueToY(0, range, layout.plotTop, layout.plotBottom),
		band = plotBand(layout),
		allDisplayValues = chartData.series.map((series) =>
			series.values.length === 0
				? undefined
				: sourceIndices.map((sourceIndex) => series.values[sourceIndex] ?? 0),
		),
		stackedPlots = isStackedMode
			? computeStackedSeriesPlots(
					allDisplayValues.map((values) => values ?? []),
					catCount,
					isPercent,
				)
			: undefined;

	for (let si = 0; si < chartData.series.length; si++) {
		const series = chartData.series[si],
			displayValues = allDisplayValues[si];
		if (!displayValues) {
			continue;
		}
		// eslint-disable-next-line one-var -- pre-existing, unrelated to this change
		const plot = stackedPlots?.[si],
			topValues = plot ? plot.cumulative : displayValues,
			pts = computeLinePoints(topValues, catCount, layout, range).map((point, index) => ({
				...point,
				x: xPositions?.[index] ?? point.x,
			})),
			c = seriesColor(series, si, chartData.colorPalette),
			firstPt = pts[0],
			lastPt = pts[pts.length - 1],
			// Stacked band: fill between this series' cumulative top and its own
			// base (the previous series' top), like a stream-graph layer, at full
			// opacity so adjacent bands read as distinct rather than washed.
			// Unstacked: fill down to the zero line. Either fill is clipped to the
			// plot, which an authored c:min / c:max can cut.
			fillPolygon = plot
				? [
						...pts,
						...computeLinePoints(plot.base, catCount, layout, range)
							.map((point, index) => ({ ...point, x: xPositions?.[index] ?? point.x }))
							.reverse(),
					]
				: firstPt && lastPt
					? [{ x: firstPt.x, y: baselineY }, ...pts, { x: lastPt.x, y: baselineY }]
					: [],
			clippedFill = clipPolygonToBand(fillPolygon, band);
		if (clippedFill.length > 0) {
			primitives.push({
				kind: 'polyline',
				points: linePointsToSvgString(clippedFill),
				stroke: 'none',
				strokeWidth: 0,
				fill: c,
				...(plot ? {} : { opacity: 0.25 }),
				part: { role: 'series', seriesIndex: si },
			} satisfies SvgPolyline);
		}
		primitives.push(
			...clippedSeriesLine(
				pts,
				false,
				{ stroke: c, strokeWidth: 2, fill: 'none', part: { role: 'series', seriesIndex: si } },
				band,
			),
		);
		pts.forEach((pt, displayIndex) => {
			if (!insideBand(pt, band)) {
				return;
			}
			const idx = sourceIndices[displayIndex] ?? displayIndex,
				part: ChartPartRef = { role: 'dataPoint', seriesIndex: si, pointIndex: idx };
			pushMarker(
				primitives,
				series,
				idx,
				pt.x,
				pt.y,
				resolveDataPointFill(series, idx, c) ?? c,
				2,
				part,
				undefined,
				buildMarkTooltip(
					series.name,
					chartData.categories[idx],
					series.values[idx] ?? 0,
					series.numberFormat,
				),
			);
		});
		if (showLabels) {
			const labelValues = plot ? plot.own : displayValues;
			labelValues.forEach((val, displayIndex) => {
				const pt = pts[displayIndex];
				if (!pt) {
					return;
				}
				if (isPercent) {
					if (val === 0) {
						return;
					}
					// Centre on the VISIBLE segment (the filled band between this
					// series' running-sum top and its own base), matching
					// PowerPoint's own percentStacked convention and stacked bar's
					// `y + h/2` placement, rather than a fixed offset above the top
					// line: at extreme category skews a thin band's label used to
					// float noticeably off the band it names.
					const baseVal = plot?.base[displayIndex] ?? 0,
						topVal = plot?.cumulative[displayIndex] ?? val,
						centreY = visibleSpanCentre(
							valueToY(baseVal, range, layout.plotTop, layout.plotBottom),
							valueToY(topVal, range, layout.plotTop, layout.plotBottom),
							band,
						);
					if (centreY === undefined) {
						return;
					}
					dataLabels.push({
						kind: 'text',
						x: pt.x,
						y: centreY,
						text: `${Math.round(val)}%`,
						fontSize: DEFAULT_CHART_DATA_LABEL_PX,
						fill: '#ffffff',
						textAnchor: 'middle',
						dominantBaseline: 'central',
						fontWeight: 'bold',
						...dataLabelFontOverride(
							resolveDataLabelTextStyle(
								chartData,
								series,
								sourceIndices[displayIndex] ?? displayIndex,
							),
						),
					});
					return;
				}
				// eslint-disable-next-line one-var -- an early return sits between this const and the previous one
				const pointIndex = sourceIndices[displayIndex] ?? displayIndex,
					label = buildDataLabelText({ chartData, series, pointIndex, value: val });
				if (label === undefined) {
					return;
				}
				// PowerPoint centres an area label in the band it names, the
				// middle of this series' own slice at the category (COM,
				// callouts-com.pptx), and points its callout / leader line there.
				const bandBase = plot
						? valueToY(plot.base[displayIndex] ?? 0, range, layout.plotTop, layout.plotBottom)
						: baselineY,
					// The middle of the part of the band the plot shows: an authored
					// c:min / c:max can cut it, and a band wholly cut away has no label.
					midY = visibleSpanCentre(pt.y, bandBase, band);
				if (midY === undefined) {
					return;
				}
				// eslint-disable-next-line one-var -- an early return sits between this const and the previous one
				const mid = { x: pt.x, y: midY },
					anchor = resolveMarkerLabelPlacement(
						chartData,
						series,
						pointIndex,
						mid,
						{ width: layout.svgWidth, height: layout.svgHeight },
						6,
						'ctr',
					);
				pushPointLabel(dataLabels, labelBoxes, chartData, series, pointIndex, mid, {
					kind: 'text',
					x: anchor.x,
					y: anchor.y,
					text: label.text,
					fontSize: DEFAULT_CHART_DATA_LABEL_PX,
					fill: label.color ?? '#334155',
					textAnchor: anchor.textAnchor,
					...(anchor.dominantBaseline ? { dominantBaseline: anchor.dominantBaseline } : {}),
					...dataLabelFontOverride(resolveDataLabelTextStyle(chartData, series, pointIndex)),
				});
			});
		}
	}
	return { primitives: [...primitives, ...labelBoxes], dataLabels };
}
