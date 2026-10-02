import type { CellRange } from '../address.js';
import type { DifferentialStyle, PatternType } from '../model.js';

/** A font ready to paint. */
export interface FontView {
	family: string;
	/** CSS pixels at 100% zoom (points * 4/3). */
	sizePx: number;
	bold: boolean;
	italic: boolean;
	underline?: 'single' | 'double';
	strike: boolean;
	/** `#RRGGBB`. */
	color: string;
	vertAlign?: 'super' | 'sub';
}

export interface EdgeView {
	widthPx: number;
	style: 'solid' | 'dashed' | 'dotted' | 'double';
	color: string;
}

export type FillView =
	| { background: string }
	| { gradient: string /* css */ }
	| { pattern: PatternType; fg: string; bg: string };

export interface BordersView {
	top?: EdgeView;
	right?: EdgeView;
	bottom?: EdgeView;
	left?: EdgeView;
	diagonalUp?: EdgeView;
	diagonalDown?: EdgeView;
}

export type HAlignView =
	| 'left'
	| 'center'
	| 'right'
	| 'justify'
	| 'fill'
	| 'centerContinuous'
	| 'distributed';
export type VAlignView = 'top' | 'center' | 'bottom' | 'justify' | 'distributed';

export interface DataBarView {
	/** 0..1 bar length relative to the cell's width. */
	fraction: number;
	color: string;
	negative?: boolean;
}

export interface IconView {
	/** Icon set name (`3TrafficLights1`, `5Arrows`, ...). */
	set: string;
	/** 0 is the lowest bucket. */
	index: number;
}

export interface CellView {
	/** Display text (number formats applied; empty when a CF rule hides the value). */
	text: string;
	/** Tooltip: the hyperlink's tooltip or target. */
	title?: string;
	font: FontView;
	fill?: FillView;
	borders: BordersView;
	hAlign: HAlignView;
	vAlign: VAlignView;
	wrap: boolean;
	shrink: boolean;
	indentPx: number;
	/** Degrees counter-clockwise, -90..90 (0 for vertical stacked text, see `verticalText`). */
	rotation: number;
	/** Stacked vertical text (`textRotation` 255). */
	verticalText?: boolean;
	isNumber: boolean;
	/** Text may spill into empty neighbours (see `overflowExtent`). */
	overflow: boolean;
	dataBar?: DataBarView;
	icon?: IconView;
	hasComment: boolean;
	hasHyperlink: boolean;
	/** A list data validation with an in-cell drop-down covers the cell. */
	validationList: boolean;
	isError: boolean;
	/** Rich text runs, each with its fully resolved font. */
	rich?: { text: string; font: FontView }[];
}

/** What conditional formatting contributes to one cell. */
export interface ConditionalFormatResult {
	/** Merged differential format of every matching rule (higher priority wins per property). */
	style?: DifferentialStyle;
	dataBar?: DataBarView;
	/** Colour-scale background, `#RRGGBB`. */
	colorScale?: string;
	icon?: IconView;
	/** A data bar or icon set with `showValue="0"` hides the cell text. */
	hideValue?: boolean;
}

export interface ConditionalFormatEvaluator {
	at(row: number, col: number): ConditionalFormatResult | undefined;
}

export interface MergeView {
	/** The cell is the top-left cell of a merge. */
	anchor: boolean;
	/** The cell is covered by a merge it does not anchor (paint nothing). */
	hidden: boolean;
	range?: CellRange;
}
