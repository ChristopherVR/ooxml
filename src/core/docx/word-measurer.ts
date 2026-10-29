// The headless text measurer the Word comparison tests share. Every probe and corpus document is set
// in Courier New, whose advance is 0.6 em, so line breaking is exact. Line heights are what Word
// reported for each font at 11 pt (`__fixtures__/word-spacing/`, `word-content-control/`); a
// font not listed here has the line height of Courier New.
import type { TextMeasurer } from './layout/measure';

/** Natural line height of a font in ems, as measured in desktop Word (points / 11 pt). */
export function wordLineHeightEm(family: string): number {
	if (family.includes('MS Gothic')) return 14.25 / 11;
	if (family.includes('Segoe UI Symbol')) return 15 / 11;
	if (family.includes('Symbol')) return 1.29;
	return 1.1328;
}

/** Monospace measurer with Word's line heights; tabs take no width (the layout positions them). */
export const wordProbeMeasurer: TextMeasurer = {
	widthOf: (text, font) => Array.from(text).filter((c) => c !== '\t').length * font.sizePx * 0.6,
	lineHeightOf: (font) => font.sizePx * wordLineHeightEm(font.family),
};
