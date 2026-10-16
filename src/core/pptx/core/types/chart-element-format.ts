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
import type { ChartGradientFill } from '../../../chart/gradient-definition';

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

/** `a:pattFill`: a preset pattern (`ST_PresetPatternVal`) in two colours. */
export interface PptxChartPatternFill {
	/** `@prst`, e.g. `pct20`, `ltUpDiag`, `smGrid`. */
	preset: string;
	/** Foreground (`a:fgClr`), `#RRGGBB`. */
	foreground: string;
	/** Background (`a:bgClr`), `#RRGGBB`. */
	background: string;
}

/** `a:blipFill`: a picture fill. */
export interface PptxChartPictureFill {
	/**
	 * The image as a `data:` URL, what the renderer paints and what the save
	 * path embeds as a new media part. Absent until a loaded picture's
	 * relationship has been resolved.
	 */
	src?: string;
	/**
	 * The chart part's relationship id (`a:blip/@r:embed`) of a loaded picture.
	 * A picture chosen in the pane has none until it is saved.
	 */
	relId?: string;
	/** `a:tile` instead of `a:stretch` (PowerPoint's "Tile picture as texture"). */
	tile?: boolean;
	/** 0 through 100, written as `a:blip/a:alphaModFix/@amt`. */
	transparency?: number;
}

/** Fill of a chart element (Fill section). */
export interface PptxChartElementFill {
	/**
	 * `none`: `a:noFill`. `solid`: `a:solidFill`. `gradient`: `a:gradFill`.
	 * `pattern`: `a:pattFill`. `picture`: `a:blipFill`. `automatic`: no fill
	 * element, so the chart style decides (PowerPoint's "Automatic").
	 */
	kind: 'none' | 'solid' | 'gradient' | 'pattern' | 'picture' | 'automatic';
	/** `#RRGGBB`, for `solid`. */
	color?: string;
	/** 0 (opaque) through 100 (clear), written as `a:alpha`. */
	transparency?: number;
	/** For `gradient`. */
	gradient?: ChartGradientFill;
	/** For `pattern`. */
	pattern?: PptxChartPatternFill;
	/** For `picture`. */
	picture?: PptxChartPictureFill;
}

/** Line end decoration (`a:headEnd` / `a:tailEnd` `@type`). */
export type PptxChartLineEndType = 'none' | 'triangle' | 'stealth' | 'diamond' | 'oval' | 'arrow';

/** Border or line of a chart element (Border / Line section). */
export interface PptxChartElementLine {
	/**
	 * `none`: `a:ln/a:noFill`. `solid`: `a:ln/a:solidFill`. `gradient`:
	 * `a:ln/a:gradFill`. `automatic`: no line fill.
	 */
	kind: 'none' | 'solid' | 'gradient' | 'automatic';
	color?: string;
	/** For `gradient`. */
	gradient?: ChartGradientFill;
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

/** `ST_BevelPresetType`. */
export type PptxChartBevelType =
	| 'relaxedInset'
	| 'circle'
	| 'slope'
	| 'cross'
	| 'angle'
	| 'softRound'
	| 'convex'
	| 'coolSlant'
	| 'divot'
	| 'riblet'
	| 'hardEdge'
	| 'artDeco';

/** `a:bevelT` / `a:bevelB`. */
export interface PptxChartBevel {
	type: PptxChartBevelType;
	/** Width in points (`@w`). */
	width: number;
	/** Height in points (`@h`). */
	height: number;
}

/**
 * 3-D Format section: `a:sp3d` (bevels, depth, contour, material) and the
 * lighting rig of `a:scene3d`. Each `null` removes the setting, `undefined`
 * leaves it as authored.
 */
export interface PptxChartElementThreeD {
	bevelTop?: PptxChartBevel | null;
	bevelBottom?: PptxChartBevel | null;
	/** Depth (`@extrusionH`, points) and its colour (`a:extrusionClr`). */
	depth?: { size: number; color?: string } | null;
	/** Contour (`@contourW`, points) and its colour (`a:contourClr`). */
	contour?: { size: number; color?: string } | null;
	/** `@prstMaterial` (`ST_PresetMaterialType`, e.g. `warmMatte`, `metal`, `plastic`). */
	material?: string | null;
	/** `a:scene3d/a:lightRig/@rig` (`ST_LightRigType`, e.g. `threePt`, `balanced`). */
	lighting?: string | null;
}

/** Text fill of the Text Options pane (`a:defRPr` / `a:rPr` fill). */
export interface PptxChartTextFill {
	kind: 'none' | 'solid' | 'gradient' | 'automatic';
	color?: string;
	transparency?: number;
	gradient?: ChartGradientFill;
}

/** Text outline (`a:defRPr/a:ln`). */
export interface PptxChartTextOutline {
	kind: 'none' | 'solid' | 'automatic';
	color?: string;
	transparency?: number;
	/** Width in points. */
	width?: number;
	dash?: string;
}

/**
 * Text Options of an element with text (title, legend, axis labels, data
 * labels, or the chart area for every chart text): `c:txPr` (and a title's
 * rich text runs). Each `null` removes the setting, `undefined` leaves it.
 */
export interface PptxChartElementText {
	fill?: PptxChartTextFill;
	outline?: PptxChartTextOutline;
	/** Text shadow and glow (`a:defRPr/a:effectLst`). */
	effects?: Pick<PptxChartElementEffects, 'shadow' | 'glow'>;
	/** Custom angle in degrees (`a:bodyPr/@rot`). */
	rotation?: number | null;
	/** Text direction (`a:bodyPr/@vert`). */
	direction?: 'horz' | 'vert' | 'vert270' | 'wordArtVert' | null;
	/** Vertical alignment (`a:bodyPr/@anchor`). */
	anchor?: 't' | 'ctr' | 'b' | null;
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
	threeD?: PptxChartElementThreeD;
	text?: PptxChartElementText;
	reset?: boolean;
}

/** Element formats by target key. */
export type PptxChartElementFormats = Record<PptxChartFormatTargetKey, PptxChartElementFormat>;
