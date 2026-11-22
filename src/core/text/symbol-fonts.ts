import { DINGBAT_CODE_POINTS } from './symbol-dingbats';

/**
 * Unicode equivalents for text written in Office's symbol-encoded fonts.
 *
 * Symbol, Wingdings and Webdings are symbol-encoded: a file stores the byte
 * that selects a glyph (`a:buChar char="§"` with `a:buFont typeface="Wingdings"`
 * is a small black square), either as Latin-1 or in the private-use block
 * U+F020..U+F0FF. Where the font is installed (Windows, Office on macOS) the
 * browser paints the intended glyph; everywhere else (iPadOS, Android, Linux)
 * the character falls back to an ordinary font and shows as `§`. These tables
 * give each glyph's Unicode counterpart so a renderer can show the same
 * symbol without the font.
 *
 * Symbol follows Adobe's published Symbol encoding. The Wingdings entries cover
 * the glyphs Office's bullet and symbol pickers offer and the ones decks use as
 * bullets; a glyph with no listed equivalent maps to `undefined`.
 */

const SYMBOL_GREEK_UPPER = 'ΑΒΧΔΕΦΓΗΙϑΚΛΜΝΟΠΘΡΣΤΥςΩΞΨΖ';
const SYMBOL_GREEK_LOWER = 'αβχδεφγηιϕκλμνοπθρστυϖωξψζ';

function symbolTable(): Map<number, string> {
	const table = new Map<number, string>();
	[...SYMBOL_GREEK_UPPER].forEach((char, index) => table.set(0x41 + index, char));
	[...SYMBOL_GREEK_LOWER].forEach((char, index) => table.set(0x61 + index, char));
	const pairs: Array<[number, string]> = [
		[0x22, '∀'],
		[0x24, '∃'],
		[0x27, '∋'],
		[0x2a, '∗'],
		[0x2d, '−'],
		[0x40, '≅'],
		[0x5c, '∴'],
		[0x5e, '⊥'],
		[0x7e, '∼'],
		[0xa1, 'ϒ'],
		[0xa2, '′'],
		[0xa3, '≤'],
		[0xa4, '⁄'],
		[0xa5, '∞'],
		[0xa6, 'ƒ'],
		[0xa7, '♣'],
		[0xa8, '♦'],
		[0xa9, '♥'],
		[0xaa, '♠'],
		[0xab, '↔'],
		[0xac, '←'],
		[0xad, '↑'],
		[0xae, '→'],
		[0xaf, '↓'],
		[0xb0, '°'],
		[0xb1, '±'],
		[0xb2, '″'],
		[0xb3, '≥'],
		[0xb4, '×'],
		[0xb5, '∝'],
		[0xb6, '∂'],
		[0xb7, '•'],
		[0xb8, '÷'],
		[0xb9, '≠'],
		[0xba, '≡'],
		[0xbb, '≈'],
		[0xbc, '…'],
		[0xc0, 'ℵ'],
		[0xc1, 'ℑ'],
		[0xc2, 'ℜ'],
		[0xc3, '℘'],
		[0xc4, '⊗'],
		[0xc5, '⊕'],
		[0xc6, '∅'],
		[0xc7, '∩'],
		[0xc8, '∪'],
		[0xc9, '⊃'],
		[0xca, '⊇'],
		[0xcb, '⊄'],
		[0xcc, '⊂'],
		[0xcd, '⊆'],
		[0xce, '∈'],
		[0xcf, '∉'],
		[0xd0, '∠'],
		[0xd1, '∇'],
		[0xd2, '®'],
		[0xd3, '©'],
		[0xd4, '™'],
		[0xd5, '∏'],
		[0xd6, '√'],
		[0xd7, '⋅'],
		[0xd8, '¬'],
		[0xd9, '∧'],
		[0xda, '∨'],
		[0xdb, '⇔'],
		[0xdc, '⇐'],
		[0xdd, '⇑'],
		[0xde, '⇒'],
		[0xdf, '⇓'],
		[0xe0, '◊'],
		[0xe1, '〈'],
		[0xe5, '∑'],
		[0xf1, '〉'],
		[0xf2, '∫'],
	];
	for (const [code, char] of pairs) table.set(code, char);
	return table;
}

function wingdingsTable(): Map<number, string> {
	const table = new Map<number, string>([
		[0x21, '✏'],
		[0x22, '✂'],
		[0x23, '✁'],
		[0x28, '☎'],
		[0x2a, '✉'],
		[0x36, '⌛'],
		[0x37, '⌨'],
		[0x3f, '✍'],
		[0x41, '✌'],
		[0x45, '☜'],
		[0x46, '☞'],
		[0x47, '☝'],
		[0x48, '☟'],
		[0x4a, '☺'],
		[0x4c, '☹'],
		[0x4e, '☠'],
		[0x51, '✈'],
		[0x52, '☼'],
		[0x54, '❄'],
		[0x56, '✞'],
		[0x58, '✠'],
		[0x59, '✡'],
		[0x5a, '☪'],
		[0x5b, '☯'],
		[0x5c, 'ॐ'],
		[0x5d, '☸'],
		[0x6c, '●'],
		[0x6d, '❍'],
		[0x6e, '■'],
		[0x6f, '□'],
		[0x71, '❑'],
		[0x72, '❒'],
		[0x73, '⬧'],
		[0x74, '⧫'],
		[0x75, '◆'],
		[0x76, '❖'],
		[0x77, '⬥'],
		[0x78, '⌧'],
		[0x79, '⍓'],
		[0x7a, '⌘'],
		[0x7b, '❀'],
		[0x7c, '✿'],
		[0x7d, '❝'],
		[0x7e, '❞'],
		[0x80, '⓪'],
		[0x8b, '⓿'],
		[0x9e, '·'],
		[0x9f, '•'],
		[0xa1, '○'],
		[0xa4, '◉'],
		[0xa5, '◎'],
		[0xa7, '▪'],
		[0xa8, '◻'],
		[0xaa, '✦'],
		[0xab, '★'],
		[0xac, '✶'],
		[0xad, '✴'],
		[0xae, '✹'],
		[0xaf, '✵'],
		[0xb1, '⌖'],
		[0xd5, '⌫'],
		[0xd6, '⌦'],
		[0xd8, '➢'],
		[0xdf, '←'],
		[0xe0, '→'],
		[0xe1, '↑'],
		[0xe2, '↓'],
		[0xe8, '➔'],
		[0xef, '⇦'],
		[0xf0, '⇨'],
		[0xf1, '⇧'],
		[0xf2, '⇩'],
		[0xfb, '✗'],
		[0xfc, '✔'],
		[0xfd, '☒'],
		[0xfe, '☑'],
	]);
	// The zodiac signs (^ to i) and the circled digits 1-10, white then black.
	for (let i = 0; i < 12; i++) table.set(0x5e + i, String.fromCodePoint(0x2648 + i));
	for (let i = 0; i < 10; i++) {
		table.set(0x81 + i, String.fromCodePoint(0x2460 + i));
		table.set(0x8c + i, String.fromCodePoint(0x2776 + i));
	}
	return table;
}

const TABLES: ReadonlyMap<string, ReadonlyMap<number, string>> = new Map([
	['symbol', symbolTable()],
	['wingdings', wingdingsTable()],
]);

/** Fonts whose characters select glyphs by code rather than by Unicode meaning. */
export const OFFICE_SYMBOL_FONTS = [
	'Symbol',
	'Wingdings',
	'Wingdings 2',
	'Wingdings 3',
	'Webdings',
] as const;

const SYMBOL_ENCODED = new Set(OFFICE_SYMBOL_FONTS.map((font) => font.toLowerCase()));

function normaliseFamily(family: string | undefined): string {
	return (family ?? '')
		.split(',')[0]!
		.trim()
		.replace(/^["']|["']$/g, '')
		.toLowerCase();
}

/** Whether `family` is one of Office's symbol-encoded fonts (Symbol, Wingdings 1-3, Webdings). */
export function isSymbolEncodedFont(family: string | undefined): boolean {
	return SYMBOL_ENCODED.has(normaliseFamily(family));
}

/**
 * The Unicode equivalent of one character painted in a symbol-encoded font, or
 * `undefined` when the font is not symbol-encoded or the glyph has no listed
 * equivalent. Private-use code points (U+F020..U+F0FF) are read as their byte.
 */
export function symbolCharToUnicode(family: string | undefined, char: string): string | undefined {
	const name = normaliseFamily(family);
	if (!SYMBOL_ENCODED.has(name)) return undefined;
	const code = char.codePointAt(0);
	if (code === undefined) return undefined;
	const byte = code >= 0xf020 && code <= 0xf0ff ? code - 0xf000 : code;
	const known = TABLES.get(name)?.get(byte);
	const standard = DINGBAT_CODE_POINTS[name]?.[byte - 0x21];
	return known ?? (standard ? String.fromCodePoint(standard) : undefined);
}

/** Portable Unicode characters corresponding to the glyphs of a legacy symbol font. */
export function officeSymbolFontGlyphs(family: string): string[] {
	const glyphs = new Set<string>();
	for (let byte = 0x21; byte <= 0xff; byte++) {
		const symbol = symbolCharToUnicode(family, String.fromCharCode(byte));
		if (symbol) glyphs.add(symbol);
	}
	return [...glyphs];
}

/**
 * `text` painted in `family`, rewritten in Unicode: `undefined` unless the font
 * is symbol-encoded and every character has an equivalent (spaces stay).
 */
export function symbolTextToUnicode(family: string | undefined, text: string): string | undefined {
	if (!isSymbolEncodedFont(family) || text.length === 0) return undefined;
	let out = '';
	for (const char of text) {
		if (char === ' ' || char === ' ') {
			out += char;
			continue;
		}
		const mapped = symbolCharToUnicode(family, char);
		if (mapped === undefined) return undefined;
		out += mapped;
	}
	return out;
}

/** A typed code point (`00A9`, `U+00A9`) as its character; undefined when not printable text. */
export function parseOfficeSymbolCode(text: string): string | undefined {
	const match = /^\s*(?:u\+|0x)?([0-9a-f]{1,6})\s*$/i.exec(text);
	if (!match) return undefined;
	const code = parseInt(match[1]!, 16);
	if (
		code < 0x20 ||
		(code >= 0x7f && code <= 0x9f) ||
		(code >= 0xd800 && code <= 0xdfff) ||
		(code & 0xfffe) === 0xfffe ||
		(code >= 0xfdd0 && code <= 0xfdef) ||
		code > 0x10ffff
	)
		return undefined;
	return String.fromCodePoint(code);
}
