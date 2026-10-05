/**
 * Excel's column and row size conversions. Column widths in SpreadsheetML (`<col width>`,
 * `defaultColWidth`) are "file widths": characters of the default font's maximum digit width
 * *including* the 5 px cell padding. The width Excel shows in its UI ("8.43") excludes the padding;
 * {@link charactersToColumnWidth} and {@link pixelsToCharacters} convert that one.
 */

/** Maximum digit width in pixels of 11pt Calibri at 96 dpi, Excel's default font. */
export const DEFAULT_MAX_DIGIT_WIDTH = 7;

/** Excel's default `baseColWidth` (characters) when a sheet records no `defaultColWidth`. */
export const DEFAULT_BASE_COL_WIDTH = 8;

/** Pixels of left plus right cell padding Excel adds to every column. */
export const COLUMN_PADDING_PX = 5;

/** Pixels per indent level at 100% zoom with the default font (Excel uses about three spaces). */
export const INDENT_PX_PER_LEVEL = 9;

const trunc = Math.trunc;

/**
 * Pixels of a column with file width `width`:
 * `Truncate(((256 * width + Truncate(128 / mdw)) / 256) * mdw)` (ECMA-376 Part 1, 18.3.1.13).
 */
export function columnWidthToPixels(width: number, mdw = DEFAULT_MAX_DIGIT_WIDTH): number {
	if (!(width > 0)) return 0;
	return trunc(((256 * width + trunc(128 / mdw)) / 256) * mdw);
}

/** The file width that renders as `px` pixels: `Truncate(px / mdw * 256) / 256`. */
export function pixelsToColumnWidth(px: number, mdw = DEFAULT_MAX_DIGIT_WIDTH): number {
	if (!(px > 0)) return 0;
	return trunc((px / mdw) * 256) / 256;
}

/**
 * The file width for a width typed in Excel's column-width dialog (characters without padding):
 * `Truncate((chars * mdw + 5) / mdw * 256) / 256`. `8.43` gives `9.140625`.
 */
export function charactersToColumnWidth(chars: number, mdw = DEFAULT_MAX_DIGIT_WIDTH): number {
	if (!(chars > 0)) return 0;
	return trunc(((chars * mdw + COLUMN_PADDING_PX) / mdw) * 256) / 256;
}

/** Pixels of a column `chars` characters wide as Excel's UI counts them (`8.43` -> 64 at mdw 7). */
export const charactersToPixels = (chars: number, mdw = DEFAULT_MAX_DIGIT_WIDTH): number =>
	columnWidthToPixels(charactersToColumnWidth(chars, mdw), mdw);

/** The width Excel's UI shows for a column of `px` pixels: `Truncate((px - 5) / mdw * 100 + 0.5) / 100`. */
export function pixelsToCharacters(px: number, mdw = DEFAULT_MAX_DIGIT_WIDTH): number {
	if (px <= COLUMN_PADDING_PX) return 0;
	return trunc(((px - COLUMN_PADDING_PX) / mdw) * 100 + 0.5) / 100;
}

/** The default column width in pixels for a sheet, from its `defaultColWidth` file width or from `baseColWidth`. */
export function defaultColumnPixels(
	defaultColWidth: number | undefined,
	mdw = DEFAULT_MAX_DIGIT_WIDTH,
	baseColWidth = DEFAULT_BASE_COL_WIDTH,
): number {
	if (defaultColWidth !== undefined && defaultColWidth > 0)
		return columnWidthToPixels(defaultColWidth, mdw);
	// baseColWidth characters plus padding, rounded up to a multiple of 8 pixels (ECMA 18.3.1.81).
	return Math.ceil((baseColWidth * mdw + COLUMN_PADDING_PX) / 8) * 8;
}

/** Points to whole pixels at 96 dpi (`15` -> `20`). */
export const pointsToPixels = (pt: number): number => (pt > 0 ? Math.round((pt * 96) / 72) : 0);

/** Pixels at 96 dpi to points (`20` -> `15`). */
export const pixelsToPoints = (px: number): number => (px * 72) / 96;
