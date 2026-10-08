// The chart summary the SVG painter draws: the first chart group's type and series with their
// cached data and DrawingML paint, plus the legend, title and formatting metadata. Spreadsheet
// charts extend it with their anchor and legacy colours; Word and PowerPoint build it from a
// `ChartSpace` (`chartSummaryFromSpace`).
import type { DrawingColor, DrawingFill } from '../../drawingml/types';
import type { ChartStyleDefinition } from '../style-definition';

/** The chart families a summary records; the painter draws all but bubble, stock and surface. */
export type ChartSummaryType =
	| 'bar'
	| 'column'
	| 'line'
	| 'pie'
	| 'doughnut'
	| 'area'
	| 'scatter'
	| 'radar'
	| 'bubble'
	| 'stock'
	| 'surface';

export interface ChartSummarySeries {
	name?: string;
	/** Formula reference for the series name (`Sheet1!$B$1`). */
	nameRef?: string;
	categoriesRef?: string;
	valuesRef?: string;
	/** Cached categories and values from the chart part, used when references cannot resolve. */
	categories: (string | number)[];
	values: (number | null)[];
	/** DrawingML fill/line color, including theme transforms. */
	drawingColor?: DrawingColor;
	/** Explicit colors of individual points (indexed by c:dPt/c:idx). */
	pointColors?: Record<number, DrawingColor>;
	/** Non-solid primary DrawingML fill/line paint. Solid fills use drawingColor. */
	fill?: DrawingFill;
	/** Non-solid individual point fills; solid point overrides use pointColors. */
	pointFills?: Record<number, DrawingFill>;
	/** Opaque imported DrawingML effect list, retained during chart regeneration. */
	effectsXml?: string;
}

export interface ChartSummary<S extends ChartSummarySeries = ChartSummarySeries> {
	chartType: ChartSummaryType;
	/** Bar/column grouping. */
	grouping?: 'clustered' | 'stacked' | 'percentStacked' | 'standard';
	/** Gap between category clusters, as a percentage of one bar's width (0..500). */
	barGapWidth?: number;
	/** Overlap of adjacent series, in percent (-100..100). */
	barOverlap?: number;
	title?: string;
	series: S[];
	showLegend: boolean;
	legendPosition?: 'r' | 'l' | 't' | 'b' | 'tr';
	/** Office color-style id (ChartColor). Manual series/point colors retain precedence. */
	colorPalette?: number;
	/** Imported Office chart-style metadata. Style authoring is not yet supported. */
	styleDefinition?: ChartStyleDefinition;
	/** Direct formatting. Axis flags and fills serialize; other metadata edits remain unsupported. */
	formatting?: ChartStyleDefinition;
}

/** A document theme resolved for chart painting. */
export interface ChartThemePalette {
	/** 12 colours in theme order (lt1, dk1, lt2, dk2, accent1-6, hlink, folHlink) as `RRGGBB`. */
	colors: string[];
	majorFont: string;
	minorFont: string;
}
