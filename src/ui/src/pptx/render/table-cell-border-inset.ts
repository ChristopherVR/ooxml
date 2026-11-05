/**
 * table-cell-border-inset.ts: PowerPoint measures a cell margin from the grid
 * line, so a row is `content + marT + marB` tall whatever the borders are.
 * `border-collapse` puts half of each border inside the cell box, which would
 * add the full border width (about 1px per row for a hairline grid) to every
 * row. This takes that half out of the padding, using the borders the cell
 * finally paints (its own and the ones its table style gives it), for the
 * default margins as well as explicit ones.
 *
 * @module table-cell-border-inset
 */
import type { PptxTableCellStyle } from 'ooxml-core/pptx';

import type { TableCellCss } from './table-style';

type Side = 'Top' | 'Bottom' | 'Left' | 'Right';

const SIDES: readonly Side[] = ['Top', 'Bottom', 'Left', 'Right'];

/** The width in px a resolved `border-*` value paints, 0 for none. */
export function borderCssWidthPx(value: unknown): number {
	if (typeof value !== 'string') {
		return 0;
	}
	const match = /^\s*(?:max\(\s*)?([\d.]+)px/u.exec(value);
	return match?.[1] ? Number.parseFloat(match[1]) : 0;
}

function pixels(value: unknown): number | undefined {
	if (typeof value === 'number') {
		return value;
	}
	const match = typeof value === 'string' ? /^\s*([\d.]+)px\s*$/u.exec(value) : null;
	return match?.[1] ? Number.parseFloat(match[1]) : undefined;
}

/**
 * Set each side's padding to its margin (the cell's own, else the default
 * already in `css`) less half of the border painted on that side.
 */
export function insetPaddingForBorders(
	css: TableCellCss,
	cellStyle: PptxTableCellStyle | undefined,
): void {
	for (const side of SIDES) {
		const padKey = `padding${side}` as const;
		const own = cellStyle?.[`margin${side}` as const];
		const margin = own ?? pixels(css[padKey]);
		if (margin === undefined) {
			continue;
		}
		const border = borderCssWidthPx(css[`border${side}` as const]);
		css[padKey] = `${Math.max(0, margin - border / 2)}px`;
	}
}
