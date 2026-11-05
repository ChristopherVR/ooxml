/**
 * SDK input for what a chart shows and how its series, labels and text look at
 * creation time: blank points, combination series, the secondary axis, series
 * lines, markers and fills, per-point colours, data labels and fonts. Each
 * field maps onto the chart model through the same operations an edit uses
 * (`chart-content-operations.ts`), so a caller can describe a chart fully in
 * `ChartInput` without post-processing.
 *
 * Units follow the chart operations: font sizes, line and marker outline
 * widths in points, marker sizes in points (2..72), angles in degrees.
 *
 * @module sdk/chart-content-input
 */

import type {
	PptxChartDataLabelPosition,
	PptxChartMarkerSymbol,
	PptxChartScatterStyle,
} from '../../types/chart';
import type { StrokeDashType } from '../../types/common';

/** Chart text font. Sizes are in points; colours are `#RRGGBB` or `RRGGBB`. */
export interface ChartFontInput {
	/** Font size in points. */
	size?: number;
	bold?: boolean;
	italic?: boolean;
	/** Hex colour. */
	color?: string;
	/** Latin typeface, e.g. `Calibri`. */
	family?: string;
}

/**
 * Data labels of a series (`c:ser/c:dLbls`), or of every series when given on
 * `ChartInput.dataLabels` (a series' own `dataLabels` overrides it field by field).
 */
export interface ChartDataLabelsInput {
	/** Show the value (`c:showVal`). */
	showValue?: boolean;
	/** Show the category name (`c:showCatName`). */
	showCategory?: boolean;
	/** Show the series name (`c:showSerName`). */
	showSeriesName?: boolean;
	/** Show the percentage (`c:showPercent`, pie family). */
	showPercent?: boolean;
	/** Show the legend key (`c:showLegendKey`). */
	showLegendKey?: boolean;
	/**
	 * Label position (`c:dLblPos`). Only the positions PowerPoint offers for the
	 * series' chart type are accepted (`chartModelLabelPositions`): bars take
	 * `ctr`, `inEnd`, `inBase`, `outEnd` (no `outEnd` when stacked), lines,
	 * scatter and bubble `ctr`, `l`, `r`, `t`, `b`, the pie family `ctr`,
	 * `inEnd`, `outEnd`, `bestFit`; area, radar, doughnut, surface and 3-D bars
	 * and lines take none.
	 */
	position?: PptxChartDataLabelPosition;
	/** Number format code of the labels (`c:numFmt`, not linked to the source). */
	numberFormat?: string;
	/** Label font (`c:dLbls/c:txPr`). */
	font?: ChartFontInput;
	/** Custom text angle in degrees, -90..90 (`c:txPr/a:bodyPr/@rot`). */
	rotation?: number;
	/** `false` keeps each label on one line (`a:bodyPr/@wrap="none"`), `true` wraps it. */
	wrap?: boolean;
	/** Label box fill: a hex colour or `'none'` (`c:dLbls/c:spPr`). */
	fill?: string;
}

/** Series line (`c:ser/c:spPr/a:ln`). */
export interface ChartSeriesLineInput {
	/** Line width in points. */
	width?: number;
	/** Preset dash (`a:prstDash`); `custom` is not accepted. */
	dash?: Exclude<StrokeDashType, 'custom'>;
}

/** Series marker (`c:ser/c:marker`). Line, scatter and radar series. */
export interface ChartSeriesMarkerInput {
	symbol: Exclude<PptxChartMarkerSymbol, 'picture'>;
	/** Marker size in points (`ST_MarkerSize`, 2..72). */
	size?: number;
	/** Marker fill, a hex colour. */
	fill?: string;
	/** Marker outline: a colour and width in points, or `'none'`. */
	line?: { color?: string; width?: number } | 'none';
}

/** Series chart types a combination chart can mix. */
export type ChartComboSeriesType = 'bar' | 'line' | 'area';

/** How a chart draws blank (`null`) points (`c:dispBlanksAs`). */
export type ChartDisplayBlanksAs = 'gap' | 'zero' | 'span';

/** The `ChartInput` fields described here. */
export interface ChartContentInput {
	/** How blank (`null`) values are drawn: a gap (default), zero, or a line spanning them. */
	dispBlanksAs?: ChartDisplayBlanksAs;
	/** How a scatter chart joins its points (`c:scatterStyle`). Scatter only. */
	scatterStyle?: PptxChartScatterStyle;
	/** Data labels for every series; a series' own `dataLabels` overrides them field by field. */
	dataLabels?: ChartDataLabelsInput;
	/** Chart title font (`c:title`). */
	titleFont?: ChartFontInput;
	/** Legend font (`c:legend/c:txPr`). */
	legendFont?: ChartFontInput;
	/**
	 * Colours, written where PowerPoint draws them: one per point (`c:dPt`) on
	 * the pie family (pie, pie3D, doughnut, ofPie), whose points vary in colour,
	 * and one per series otherwise. Cycled when shorter. A series' own `color`,
	 * `gradientFill` or `pointColors` wins.
	 */
	palette?: string[];
}

/** The `ChartSeriesInput` fields described here. */
export interface ChartSeriesContentInput {
	/** Plots the series as bars, a line or an area in a combination chart (bar, line or area charts). */
	chartType?: ChartComboSeriesType;
	/**
	 * Plots the series against a secondary value axis on the far side, crossing
	 * at the maximum, with a hidden secondary category axis (bar, line, area,
	 * scatter and bubble charts).
	 */
	secondaryAxis?: boolean;
	/** Series line width and dash, or `'none'` for no line (no border on a bar). */
	line?: ChartSeriesLineInput | 'none';
	/** Smooths a line or scatter series (`c:smooth`). */
	smooth?: boolean;
	marker?: ChartSeriesMarkerInput;
	/** Labels of this series' points. */
	dataLabels?: ChartDataLabelsInput;
	/** Scatter and bubble X values, one per value (`c:xVal`). Default: the categories. */
	xValues?: number[];
	/** Bubble sizes, one per value (`c:bubbleSize`). Bubble only; default 1. */
	bubbleSizes?: number[];
	/** `'none'` leaves the series unfilled (`a:noFill`). Area-filled chart types only. */
	fill?: 'none';
	/**
	 * One fill colour per point (`c:dPt`), in category order, at most one per value; `null` keeps
	 * the point's colour. On a series drawn with markers (line, scatter, stock, unfilled radar) it
	 * fills the point's marker (`c:dPt/c:marker`), which keeps the series marker's symbol, size and
	 * outline; such a series whose marker symbol is `'none'` refuses it.
	 */
	pointColors?: (string | null)[];
}
