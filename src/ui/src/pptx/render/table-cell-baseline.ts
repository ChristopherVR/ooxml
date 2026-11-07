/**
 * Puts a table cell's paragraph on PowerPoint's baselines.
 *
 * A cell lays its text out with the same line model as a text box (measured
 * on PowerPoint 16 exports: a table cell and a text box with the same font,
 * size and `a:lnSpc` put their baselines on the same pixel for 12 to 40pt
 * text at 50% to 300% and at 6 to 72pt exact), so the move
 * `paragraph-baseline` computes for shape text applies unchanged. Without it
 * the browser centres the leading above and below the line, which put a 150%
 * paragraph 3px high and a 72pt exact one 18px high.
 *
 * @module render/table-cell-baseline
 */
import {
	baselinePlacementStyle,
	type BaselinePlacement,
	type FontBoxMeasure,
} from './paragraph-baseline';
import { DEFAULT_FONT_FAMILY } from '../constants';
import type { ParagraphRun } from './paragraph-types';
import { cellRunStyle, type CellTextRun, type TableCellCss } from './table-style';
import { DEFAULT_LINE_HEIGHT } from './text-line-height';

/** CSS pixels in a point. */
const PX_PER_PT = 96 / 72;

/** A length in px from a CSS `<n>pt` / `<n>px` string or a bare number of px. */
function lengthPx(value: unknown): number | undefined {
	if (typeof value === 'number') {
		return value;
	}
	const match = typeof value === 'string' ? /^(-?[\d.]+)(pt|px)$/u.exec(value.trim()) : null;
	if (!match) {
		return undefined;
	}
	const amount = Number.parseFloat(match[1] ?? '');
	return match[2] === 'pt' ? amount * PX_PER_PT : amount;
}

/** What {@link cellBaselinePlacement} reads from one paragraph block of a cell. */
export interface CellBaselineInput {
	/** The block's runs. */
	runs: readonly CellTextRun[];
	/** The block's own `line-height` (a multiplier or a px string), else the cell's applies. */
	lineHeight: number | string | undefined;
	/** The block's own font size (pt), set when every run sets one. */
	strutPt: number | undefined;
	/** The cell's resolved style: its font and line height for runs and blocks that set none. */
	cellCss: TableCellCss | undefined;
	/** The table's default text size (pt), for a cell that sets none. */
	defaultFontPt?: number | undefined;
	measure?: FontBoxMeasure | undefined;
}

/**
 * The inline style that moves a cell paragraph onto PowerPoint's baselines,
 * or `undefined` when it needs none or cannot be modelled (no known font
 * size, several faces, several sizes other than at a roomy exact pitch,
 * super- and subscripts).
 */
export function cellBaselinePlacement(input: CellBaselineInput): BaselinePlacement | undefined {
	const { cellCss } = input;
	if (!cellCss) {
		// The cell's own font is what runs without one measure against.
		return undefined;
	}
	const cellSizePx =
		lengthPx(cellCss.fontSize) ??
		(input.defaultFontPt === undefined ? undefined : input.defaultFontPt * PX_PER_PT);
	const strutSizePx = input.strutPt === undefined ? cellSizePx : input.strutPt * PX_PER_PT;
	if (strutSizePx === undefined) {
		return undefined;
	}
	const runs: ParagraphRun[] = [];
	for (const run of input.runs) {
		if (!run.text || run.isLineBreak || run.isParagraphBreak) {
			continue;
		}
		const css = cellRunStyle(run, cellCss);
		const size = lengthPx(css.fontSize) ?? cellSizePx;
		if (size === undefined) {
			return undefined;
		}
		runs.push({
			text: run.text,
			style: {
				fontSize: size,
				...(typeof css.fontFamily === 'string' ? { fontFamily: css.fontFamily } : {}),
				...(css.fontWeight === undefined ? {} : { fontWeight: css.fontWeight }),
				...(css.fontStyle === 'italic' ? { fontStyle: 'italic' as const } : {}),
				...(css.verticalAlign === undefined ? {} : { verticalAlign: css.verticalAlign }),
				...(css.lineHeight === undefined ? {} : { lineHeight: css.lineHeight }),
			},
		});
	}
	return baselinePlacementStyle({
		runs,
		lineHeight: input.lineHeight ?? lineHeightOf(cellCss),
		strutSizePx,
		strutFamily: typeof cellCss.fontFamily === 'string' ? cellCss.fontFamily : DEFAULT_FONT_FAMILY,
		...(input.measure ? { measure: input.measure } : {}),
	});
}

/** The cell's `line-height`, single spacing when it sets none. */
function lineHeightOf(cellCss: TableCellCss | undefined): number | string {
	const value = cellCss?.lineHeight;
	return typeof value === 'number' || typeof value === 'string' ? value : DEFAULT_LINE_HEIGHT;
}
