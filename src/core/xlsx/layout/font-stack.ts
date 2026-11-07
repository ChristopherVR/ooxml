// The CSS font stack every place that paints or measures workbook text uses (cells, the cell
// editor, column autofit, automatic row heights, shapes, print and page layout). It is the shared
// workbook stack (`text/css-font-family`), which already lists the clone the webfont fallback
// downloads after the common Office fonts. For a family that stack does not know but that has a
// verified clone (`text/webfonts/metric-clones`: Helvetica, Arial Black, Aptos Serif...), the
// clone is placed right after the family, so the face downloaded for it is what the browser
// draws and measures when the family is missing.
import { cssFontFamily } from '../../text/css-font-family';
import { findMetricCompatibleGoogleFontsFamily } from '../../text/webfonts/metric-clones';

/** Excel's default font when a style names none. */
const DEFAULT_FAMILY = 'Calibri';

/** The end of the stack's leading entry: after its closing quote, or before the first comma. */
function leadEnd(stack: string): number {
	if (stack.startsWith('"')) {
		const close = stack.indexOf('"', 1);
		return close < 0 ? stack.length : close + 1;
	}
	const comma = stack.indexOf(',');
	return comma < 0 ? stack.length : comma;
}

/** The CSS `font-family` value for a workbook font name. */
export function workbookFontFamily(name: string | undefined): string {
	const family = name?.trim() || DEFAULT_FAMILY;
	const stack = cssFontFamily(family);
	const clone = findMetricCompatibleGoogleFontsFamily(family);
	if (!clone || stack.includes(`"${clone}"`)) return stack;
	const end = leadEnd(stack);
	return `${stack.slice(0, end)}, "${clone}"${stack.slice(end)}`;
}

/** Resolves a DrawingML theme font reference (`+mn-lt`, `+mj-ea`) through the workbook theme. */
export function resolveThemeTypeface(
	typeface: string | undefined,
	theme: { minorFont?: string; majorFont?: string } | undefined,
): string | undefined {
	if (!typeface) return undefined;
	if (typeface.startsWith('+mj-')) return theme?.majorFont || undefined;
	if (typeface.startsWith('+mn-')) return theme?.minorFont || undefined;
	return typeface;
}
