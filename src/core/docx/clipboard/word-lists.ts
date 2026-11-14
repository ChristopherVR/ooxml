/** List levels from Word's `@list` rules, or from the marker text when the rules are missing. */
import type { NumberingLevelDefinition } from '../numbering-model';
import type { StNumberFormat } from '../generated/wml-simple-types';
import { signedTwips, twips } from '../units';
import { cssPoints, cssStringValue, cssTwips, firstFontFamily } from './css';

/** `mso-level-number-format` values (Office HTML) and the `ST_NumberFormat` they stand for. */
const NUMBER_FORMATS: Record<string, StNumberFormat> = {
	arabic: 'decimal',
	bullet: 'bullet',
	image: 'bullet',
	'alpha-lower': 'lowerLetter',
	'alpha-upper': 'upperLetter',
	'roman-lower': 'lowerRoman',
	'roman-upper': 'upperRoman',
	'arabic-leading-zero': 'decimalZero',
	none: 'none',
	ordinal: 'ordinal',
	'cardinal-text': 'cardinalText',
	'ordinal-text': 'ordinalText',
	hex: 'hex',
	chicago: 'chicago',
};

/** Fonts whose glyphs Word stores in the private-use area (`U+F0B7` is Symbol's bullet). */
const SYMBOL_FONTS = /^(symbol|wingdings[\s\d]*|webdings)$/i;

/** A bullet character as Word stores it in `w:lvlText` for its font. */
function bulletText(text: string, font: string | undefined): string {
	const code = text.codePointAt(0);
	return code !== undefined && code < 0x100 && font && SYMBOL_FONTS.test(font)
		? String.fromCodePoint(0xf000 + code)
		: text;
}

function markerFont(declarations: ReadonlyMap<string, string>): string | undefined {
	return firstFontFamily(
		declarations.get('font-family') ??
			declarations.get('mso-ascii-font-family') ??
			declarations.get('mso-hansi-font-family'),
	);
}

/**
 * One level of an `@list lN:levelM` rule. `inherited` is the paragraph style's own formatting,
 * which supplies the indent when the level does not set one (as Word's default bullet does).
 */
export function levelFromRule(
	level: number,
	rule: ReadonlyMap<string, string>,
	inherited?: ReadonlyMap<string, string>,
): NumberingLevelDefinition {
	const format = rule.get('mso-level-number-format')?.toLowerCase();
	const numFmt = (format && NUMBER_FORMATS[format]) || 'decimal';
	const font = markerFont(rule);
	const written = rule.get('mso-level-text');
	const text = written === undefined ? `%${level + 1}.` : cssStringValue(written);
	const start = Number(rule.get('mso-level-start-at'));
	const left = cssTwips(rule.get('margin-left') ?? inherited?.get('margin-left'));
	const indent = cssTwips(rule.get('text-indent'));
	const position = rule.get('mso-level-number-position')?.toLowerCase();
	const suffix = rule.get('mso-level-suffix')?.toLowerCase();
	const tab = cssPoints(rule.get('mso-level-tab-stop'));
	return {
		level,
		start: Number.isInteger(start) && start >= 0 ? start : 1,
		numFmt,
		lvlText: numFmt === 'bullet' ? bulletText(text, font) : text,
		...(position === 'center' || position === 'right' || position === 'left'
			? { lvlJc: position }
			: {}),
		...(left !== undefined ? { indentLeftTwips: signedTwips(left) } : {}),
		...(indent !== undefined && indent < 0 ? { hangingTwips: twips(-indent) } : {}),
		...(indent !== undefined && indent > 0 ? { firstLineTwips: twips(indent) } : {}),
		...(rule.get('mso-level-legal-format')?.toLowerCase() === 'yes' ? { isLgl: true } : {}),
		...(suffix === 'space' || suffix === 'none' ? { suffix } : { suffix: 'tab' as const }),
		...(tab !== undefined && tab >= 0 ? { tabStopTwips: twips(Math.round(tab * 20)) } : {}),
		...(font ? { markerFormat: { fontFamily: font } } : {}),
	};
}

const ROMAN = /^(?=[ivxlcdm]+$)m{0,3}(cm|cd|d?c{0,3})(xc|xl|l?x{0,3})(ix|iv|v?i{0,3})$/i;

function romanValue(text: string): number {
	const values: Record<string, number> = { i: 1, v: 5, x: 10, l: 50, c: 100, d: 500, m: 1000 };
	let total = 0;
	const lower = text.toLowerCase();
	for (let index = 0; index < lower.length; index++) {
		const value = values[lower[index]!]!;
		const next = values[lower[index + 1] ?? ''] ?? 0;
		total += value < next ? -value : value;
	}
	return total;
}

function letterValue(text: string): number {
	// Word repeats the letter past z: aa is 27.
	const code = text.toLowerCase().charCodeAt(0) - 96;
	return code + 26 * (text.length - 1);
}

/**
 * A level inferred from the marker Word wrote for browsers (`1)`, `a.`, `iv.`, `·`), for HTML
 * without its stylesheet. `font` is the marker span's font, which decides a bullet's glyph.
 */
export function levelFromMarker(
	level: number,
	marker: string,
	font: string | undefined,
): NumberingLevelDefinition {
	const text = marker.replace(/\s+$/g, '').trim();
	const numbered = /^([^\w]*)(\d+|[a-z]+|[A-Z]+)([^\w]*)$/.exec(text);
	const base = { level, suffix: 'tab' as const };
	if (!numbered || (font && SYMBOL_FONTS.test(font))) {
		return {
			...base,
			start: 1,
			numFmt: 'bullet',
			lvlText: bulletText(text || String.fromCodePoint(0x2022), font),
			...(font ? { markerFormat: { fontFamily: font } } : {}),
		};
	}
	const [, before = '', body = '', after = ''] = numbered;
	const placeholder = `${before}%${level + 1}${after}`;
	if (/^\d+$/.test(body))
		return {
			...base,
			start: Number(body),
			numFmt: body.length > 1 && body.startsWith('0') ? 'decimalZero' : 'decimal',
			lvlText: placeholder,
		};
	if (ROMAN.test(body) && (body.length > 1 || /^[iI]$/.test(body)))
		return {
			...base,
			start: romanValue(body),
			numFmt: body === body.toLowerCase() ? 'lowerRoman' : 'upperRoman',
			lvlText: placeholder,
		};
	return {
		...base,
		start: letterValue(body),
		numFmt: body === body.toLowerCase() ? 'lowerLetter' : 'upperLetter',
		lvlText: placeholder,
	};
}
