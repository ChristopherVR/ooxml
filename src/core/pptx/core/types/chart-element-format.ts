/**
 * Shape formatting PowerPoint's chart Format task panes edit on one chart
 * element (Format Chart Area, Format Data Series, Format Axis, ...): the
 * Fill, Border/Line and Effects sections, which all write the element's own
 * `c:spPr` (CT_ShapeProperties).
 *
 * The chart model already carries a few of these as flat fields
 * (`PptxChartStyle.chartAreaFill`, `PptxChartSeries.color`, an axis'
 * `majorGridlinesSpPr`, ...) that the renderer reads. The element format
 * adds what those fields cannot hold (transparency, line compound/cap/join,
 * arrows, shadow, glow, soft edges, and the legend/axis/data-label boxes
 * the model had no field for) as an edit record keyed by
 * {@link PptxChartFormatTargetKey}. The save path writes each record into
 * the matching element's `c:spPr`, so an element the panes never touched
 * keeps its authored XML.
 *
 * @module types/chart-element-format
 */

/**
 * The chart element a Format pane targets.
 *
 * Axis and gridline targets name the axis by its index in
 * `PptxChartData.axes`; series and data-label targets by the series index in
 * `PptxChartData.series`. A data-label target without a series index covers
 * every series (the chart-level "Data Labels" entry of the element picker).
 */
export type PptxChartFormatTarget =
	| { kind: 'chartArea' }
	| { kind: 'plotArea' }
	| { kind: 'title' }
	| { kind: 'legend' }
	| { kind: 'axis'; axisIndex: number }
	| { kind: 'majorGridlines'; axisIndex: number }
	| { kind: 'minorGridlines'; axisIndex: number }
	| { kind: 'series'; seriesIndex: number }
	| { kind: 'dataLabels'; seriesIndex?: number };

/**
 * A {@link PptxChartFormatTarget} as a stable string (`chartArea`, `axis:1`,
 * `series:0`, `dataLabels`, `dataLabels:2`, ...), used as the record key.
 */
export type PptxChartFormatTargetKey = string;

/** Fill of a chart element (Fill section). */
export interface PptxChartElementFill {
	/**
	 * `none`: `a:noFill`. `solid`: `a:solidFill`. `automatic`: no fill element,
	 * so the chart style decides (PowerPoint's "Automatic").
	 */
	kind: 'none' | 'solid' | 'automatic';
	/** `#RRGGBB`, for `solid`. */
	color?: string;
	/** 0 (opaque) through 100 (clear), written as `a:alpha`. */
	transparency?: number;
}

/** Line end decoration (`a:headEnd` / `a:tailEnd` `@type`). */
export type PptxChartLineEndType = 'none' | 'triangle' | 'stealth' | 'diamond' | 'oval' | 'arrow';

/** Border or line of a chart element (Border / Line section). */
export interface PptxChartElementLine {
	/** `none`: `a:ln/a:noFill`. `solid`: `a:ln/a:solidFill`. `automatic`: no line fill. */
	kind: 'none' | 'solid' | 'automatic';
	color?: string;
	/** 0 through 100, written as `a:alpha`. */
	transparency?: number;
	/** Width in points (`a:ln/@w`, EMU on disk). */
	width?: number;
	/** `a:prstDash/@val` (`solid`, `dash`, `sysDot`, `lgDashDot`, ...). */
	dash?: string;
	/** `a:ln/@cmpd`. */
	compound?: 'sng' | 'dbl' | 'thickThin' | 'thinThick' | 'tri';
	/** `a:ln/@cap`. */
	cap?: 'rnd' | 'sq' | 'flat';
	/** `a:round` / `a:bevel` / `a:miter`. */
	join?: 'round' | 'bevel' | 'miter';
	/** `a:headEnd/@type`. */
	beginArrow?: PptxChartLineEndType;
	/** `a:tailEnd/@type`. */
	endArrow?: PptxChartLineEndType;
}

/** `a:outerShdw`. */
export interface PptxChartElementShadow {
	color: string;
	/** 0 through 100. */
	transparency?: number;
	/** Blur radius in points. */
	blur: number;
	/** Distance in points. */
	distance: number;
	/** Direction in degrees, clockwise from the positive x axis. */
	angle: number;
}

/** `a:glow`. */
export interface PptxChartElementGlow {
	color: string;
	/** 0 through 100. */
	transparency?: number;
	/** Radius (size) in points. */
	radius: number;
}

/** Effects section: each `null` removes the effect, `undefined` leaves it as authored. */
export interface PptxChartElementEffects {
	shadow?: PptxChartElementShadow | null;
	glow?: PptxChartElementGlow | null;
	/** Soft-edge radius in points (`a:softEdge/@rad`). */
	softEdge?: number | null;
}

/**
 * The pane edits recorded for one element. Each section present replaces
 * that part of the element's `c:spPr`; an absent section is left as authored.
 * `reset` (Reset to Match Style) drops the element's `c:spPr` so the chart
 * style draws it again; later edits recorded after a reset apply on top.
 */
export interface PptxChartElementFormat {
	fill?: PptxChartElementFill;
	line?: PptxChartElementLine;
	effects?: PptxChartElementEffects;
	reset?: boolean;
}

/** Element formats by target key. */
export type PptxChartElementFormats = Record<PptxChartFormatTargetKey, PptxChartElementFormat>;
