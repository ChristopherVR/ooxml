/**
 * The 12 standard PDF fonts (Helvetica, Times, Courier in four styles), WinAnsi text encoding
 * and approximate advance widths. Nothing is embedded: a viewer substitutes its own copy, so
 * text stays selectable and searchable but glyph shapes and widths follow the standard font,
 * not the document's typeface. DOM-free.
 */

/** The standard-14 BaseFont name for a CSS-like family list and style. */
export function pdfBaseFont(family: string, bold: boolean, italic: boolean): string {
	const lower = family.toLowerCase();
	const sans = /sans|arial|helvetica|calibri|segoe|verdana|tahoma/.test(lower);
	const kind = /courier|consolas|mono|menlo|lucida console/.test(lower)
		? 'Courier'
		: !sans && /times|serif|georgia|cambria|garamond|palatino|book|roman|minion/.test(lower)
			? 'Times'
			: 'Helvetica';
	if (kind === 'Times')
		return bold
			? italic
				? 'Times-BoldItalic'
				: 'Times-Bold'
			: italic
				? 'Times-Italic'
				: 'Times-Roman';
	const style = bold ? (italic ? '-BoldOblique' : '-Bold') : italic ? '-Oblique' : '';
	return `${kind}${style}`;
}

// Helvetica advance widths (1000 units/em) for U+0020..U+007E.
const HELVETICA = [
	278, 278, 355, 556, 556, 889, 667, 191, 333, 333, 389, 584, 278, 333, 278, 278, 556, 556, 556,
	556, 556, 556, 556, 556, 556, 556, 278, 278, 584, 584, 584, 556, 1015, 667, 667, 722, 722, 667,
	611, 778, 722, 278, 500, 667, 556, 833, 722, 778, 667, 778, 722, 667, 611, 722, 667, 944, 667,
	667, 611, 278, 278, 278, 469, 556, 333, 556, 556, 500, 556, 556, 278, 556, 556, 222, 222, 500,
	222, 833, 556, 556, 556, 556, 333, 500, 278, 556, 500, 722, 500, 500, 500, 334, 260, 334, 584,
];

// Helvetica-Bold advance widths (1000 units/em) for U+0020..U+007E.
const HELVETICA_BOLD = [
	278, 333, 474, 556, 556, 889, 722, 238, 333, 333, 389, 584, 278, 333, 278, 278, 556, 556, 556,
	556, 556, 556, 556, 556, 556, 556, 333, 333, 584, 584, 584, 611, 975, 722, 722, 722, 722, 667,
	611, 778, 722, 278, 556, 722, 611, 833, 722, 778, 667, 778, 722, 667, 611, 722, 667, 944, 667,
	667, 611, 333, 278, 333, 584, 556, 333, 556, 611, 556, 611, 556, 333, 611, 611, 278, 278, 556,
	278, 889, 611, 611, 611, 611, 389, 556, 333, 611, 556, 778, 556, 556, 500, 389, 280, 389, 584,
];

// Times-Roman advance widths (1000 units/em) for U+0020..U+007E.
const TIMES = [
	250, 333, 408, 500, 500, 833, 778, 180, 333, 333, 500, 564, 250, 333, 250, 278, 500, 500, 500,
	500, 500, 500, 500, 500, 500, 500, 278, 278, 564, 564, 564, 444, 921, 722, 667, 667, 722, 611,
	556, 722, 722, 333, 389, 722, 611, 889, 722, 722, 556, 722, 667, 556, 611, 722, 722, 944, 722,
	722, 611, 333, 278, 333, 469, 500, 333, 444, 500, 444, 500, 444, 333, 500, 500, 278, 278, 500,
	278, 778, 500, 500, 500, 500, 333, 389, 278, 500, 500, 722, 500, 500, 444, 480, 200, 480, 541,
];

/**
 * Approximate width of `text` in em units: the standard fonts' own ASCII advances (Helvetica,
 * Helvetica-Bold, Times-Roman; bold Times scaled from roman), an average for other characters.
 */
export function pdfTextWidth(text: string, baseFont: string): number {
	if (baseFont.startsWith('Courier')) return text.length * 0.6;
	const times = baseFont.startsWith('Times');
	const bold = /Bold/.test(baseFont);
	const table = times ? TIMES : bold ? HELVETICA_BOLD : HELVETICA;
	let total = 0;
	for (let index = 0; index < text.length; index++) {
		const code = text.charCodeAt(index);
		total += code >= 32 && code <= 126 ? table[code - 32]! : 556;
	}
	return (total / 1000) * (times && bold ? 1.05 : 1);
}

// Windows-1252 code points 0x80..0x9F (0 where undefined).
const CP1252_HIGH = [
	0x20ac, 0, 0x201a, 0x0192, 0x201e, 0x2026, 0x2020, 0x2021, 0x02c6, 0x2030, 0x0160, 0x2039, 0x0152,
	0, 0x017d, 0, 0, 0x2018, 0x2019, 0x201c, 0x201d, 0x2022, 0x2013, 0x2014, 0x02dc, 0x2122, 0x0161,
	0x203a, 0x0153, 0, 0x017e, 0x0178,
];

/** `text` as WinAnsi bytes held in a Latin-1 string, or `undefined` if a character has no code. */
export function encodeWinAnsi(text: string): string | undefined {
	let out = '';
	for (const char of text) {
		const code = char.codePointAt(0)!;
		if (code === 9 || code === 10 || code === 13 || code === 0xa0) out += ' ';
		else if (code >= 32 && code < 127) out += char;
		else if (code >= 0xa1 && code <= 0xff) out += char;
		else {
			const index = CP1252_HIGH.indexOf(code);
			if (index < 0 || code === 0) return undefined;
			out += String.fromCharCode(0x80 + index);
		}
	}
	return out;
}

/** The standard Symbol font, written with its own built-in encoding rather than WinAnsi. */
export const PDF_SYMBOL_FONT = 'Symbol';

/** Whether a standard font is written with its built-in encoding (Symbol, ZapfDingbats). */
export const usesBuiltInEncoding = (baseFont: string): boolean =>
	baseFont === PDF_SYMBOL_FONT || baseFont === 'ZapfDingbats';

/** A PDF literal string for a WinAnsi string. */
export const pdfLiteral = (text: string): string =>
	`(${text.replace(/[\\()]/g, '\\$&').replace(/[\r\n]/g, ' ')})`;
