/**
 * table-cell-padding.ts: CSS padding for a table cell's margins. PowerPoint
 * measures a margin from the grid line, so the half of each border that
 * `border-collapse` puts inside the cell is taken out of the padding.
 *
 * @module table-cell-padding
 */
import type { PptxTableCellStyle } from 'ooxml-core/pptx';

import type { TableCellCss } from './table-style';

function inset(margin: number, border: number | undefined): string {
	return `${Math.max(0, margin - (border ?? 0) / 2)}px`;
}

/**
 * Padding for the margins a cell sets. `!== undefined` rather than a truthy
 * check so an explicitly zeroed margin (`<a:marL w="0"/>`, common in dense or
 * image-filled tables) still renders as `0px` instead of falling through to
 * the browser's default cell padding.
 */
export function cellPaddingCss(style: PptxTableCellStyle): TableCellCss {
	const css: TableCellCss = {};
	if (style.marginLeft !== undefined) {
		css.paddingLeft = inset(style.marginLeft, style.borderLeftWidth);
	}
	if (style.marginRight !== undefined) {
		css.paddingRight = inset(style.marginRight, style.borderRightWidth);
	}
	if (style.marginTop !== undefined) {
		css.paddingTop = inset(style.marginTop, style.borderTopWidth);
	}
	if (style.marginBottom !== undefined) {
		css.paddingBottom = inset(style.marginBottom, style.borderBottomWidth);
	}
	return css;
}
