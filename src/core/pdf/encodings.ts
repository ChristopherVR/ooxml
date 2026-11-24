/**
 * Character encodings of simple PDF fonts (ISO 32000-1 annex D) and glyph names to Unicode (the
 * Adobe Glyph List's naming rules: `uniXXXX`, `uXXXX`, ligature components and the common names).
 */

/** Windows-1252 code points for 0x80..0x9f (0 where undefined). */
const CP1252: number[] = [
	0x20ac, 0, 0x201a, 0x0192, 0x201e, 0x2026, 0x2020, 0x2021, 0x02c6, 0x2030, 0x0160, 0x2039, 0x0152,
	0, 0x017d, 0, 0, 0x2018, 0x2019, 0x201c, 0x201d, 0x2022, 0x2013, 0x2014, 0x02dc, 0x2122, 0x0161,
	0x203a, 0x0153, 0, 0x017e, 0x0178,
];

/** Mac OS Roman code points for 0x80..0xff. */
const MAC_ROMAN_HIGH: number[] = [
	0xc4, 0xc5, 0xc7, 0xc9, 0xd1, 0xd6, 0xdc, 0xe1, 0xe0, 0xe2, 0xe4, 0xe3, 0xe5, 0xe7, 0xe9, 0xe8,
	0xea, 0xeb, 0xed, 0xec, 0xee, 0xef, 0xf1, 0xf3, 0xf2, 0xf4, 0xf6, 0xf5, 0xfa, 0xf9, 0xfb, 0xfc,
	0x2020, 0xb0, 0xa2, 0xa3, 0xa7, 0x2022, 0xb6, 0xdf, 0xae, 0xa9, 0x2122, 0xb4, 0xa8, 0x2260, 0xc6,
	0xd8, 0x221e, 0xb1, 0x2264, 0x2265, 0xa5, 0xb5, 0x2202, 0x2211, 0x220f, 0x3c0, 0x222b, 0xaa, 0xba,
	0x3a9, 0xe6, 0xf8, 0xbf, 0xa1, 0xac, 0x221a, 0x192, 0x2248, 0x2206, 0xab, 0xbb, 0x2026, 0xa0,
	0xc0, 0xc3, 0xd5, 0x152, 0x153, 0x2013, 0x2014, 0x201c, 0x201d, 0x2018, 0x2019, 0xf7, 0x25ca,
	0xff, 0x178, 0x2044, 0x20ac, 0x2039, 0x203a, 0xfb01, 0xfb02, 0x2021, 0xb7, 0x201a, 0x201e, 0x2030,
	0xc2, 0xca, 0xc1, 0xcb, 0xc8, 0xcd, 0xce, 0xcf, 0xcc, 0xd3, 0xd4, 0xf8ff, 0xd2, 0xda, 0xdb, 0xd9,
	0x131, 0x2c6, 0x2dc, 0xaf, 0x2d8, 0x2d9, 0x2da, 0xb8, 0x2dd, 0x2db, 0x2c7,
];

/** StandardEncoding where it differs from Latin-1 (the rest of 0xa1..0xff is unassigned). */
const STANDARD_HIGH: Record<number, number> = {
	0x27: 0x2019,
	0x60: 0x2018,
	0xa1: 0xa1,
	0xa2: 0xa2,
	0xa3: 0xa3,
	0xa4: 0x2044,
	0xa5: 0xa5,
	0xa6: 0x192,
	0xa7: 0xa7,
	0xa8: 0xa4,
	0xa9: 0x27,
	0xaa: 0x201c,
	0xab: 0xab,
	0xac: 0x2039,
	0xad: 0x203a,
	0xae: 0xfb01,
	0xaf: 0xfb02,
	0xb1: 0x2013,
	0xb2: 0x2020,
	0xb3: 0x2021,
	0xb4: 0xb7,
	0xb6: 0xb6,
	0xb7: 0x2022,
	0xb8: 0x201a,
	0xb9: 0x201e,
	0xba: 0x201d,
	0xbb: 0xbb,
	0xbc: 0x2026,
	0xbd: 0x2030,
	0xbf: 0xbf,
	0xc1: 0x60,
	0xc2: 0xb4,
	0xc3: 0x2c6,
	0xc4: 0x2dc,
	0xc5: 0xaf,
	0xc6: 0x2d8,
	0xc7: 0x2d9,
	0xc8: 0xa8,
	0xca: 0x2da,
	0xcb: 0xb8,
	0xcd: 0x2dd,
	0xce: 0x2db,
	0xcf: 0x2c7,
	0xd0: 0x2014,
	0xe1: 0xc6,
	0xe3: 0xaa,
	0xe8: 0x141,
	0xe9: 0xd8,
	0xea: 0x152,
	0xeb: 0xba,
	0xf1: 0xe6,
	0xf5: 0x131,
	0xf8: 0x142,
	0xf9: 0xf8,
	0xfa: 0x153,
	0xfb: 0xdf,
};

export type BaseEncoding =
	| 'WinAnsiEncoding'
	| 'MacRomanEncoding'
	| 'StandardEncoding'
	| 'PDFDocEncoding';

/** The Unicode text of `code` in a base encoding, or `undefined` when unassigned. */
export function baseEncodingChar(encoding: BaseEncoding, code: number): string | undefined {
	if (encoding === 'StandardEncoding') {
		if (STANDARD_HIGH[code] !== undefined) return String.fromCodePoint(STANDARD_HIGH[code]!);
		return code >= 0x20 && code < 0x7f ? String.fromCharCode(code) : undefined;
	}
	if (code < 0x80)
		return code >= 0x20 || code === 9 || code === 10 || code === 13
			? String.fromCharCode(code)
			: undefined;
	if (encoding === 'MacRomanEncoding') return String.fromCodePoint(MAC_ROMAN_HIGH[code - 0x80]!);
	if (code < 0xa0) {
		const mapped = CP1252[code - 0x80]!;
		return mapped ? String.fromCodePoint(mapped) : undefined;
	}
	return String.fromCharCode(code);
}

const NAMES: Record<string, string> = {
	space: ' ',
	exclam: '!',
	quotedbl: '"',
	numbersign: '#',
	dollar: '$',
	percent: '%',
	ampersand: '&',
	quotesingle: "'",
	quoteright: '\u2019',
	parenleft: '(',
	parenright: ')',
	asterisk: '*',
	plus: '+',
	comma: ',',
	hyphen: '-',
	minus: '\u2212',
	period: '.',
	slash: '/',
	zero: '0',
	one: '1',
	two: '2',
	three: '3',
	four: '4',
	five: '5',
	six: '6',
	seven: '7',
	eight: '8',
	nine: '9',
	colon: ':',
	semicolon: ';',
	less: '<',
	equal: '=',
	greater: '>',
	question: '?',
	at: '@',
	bracketleft: '[',
	backslash: '\\',
	bracketright: ']',
	asciicircum: '^',
	underscore: '_',
	grave: '`',
	quoteleft: '\u2018',
	braceleft: '{',
	bar: '|',
	braceright: '}',
	asciitilde: '~',
	exclamdown: '\u00a1',
	cent: '\u00a2',
	sterling: '\u00a3',
	currency: '\u00a4',
	yen: '\u00a5',
	brokenbar: '\u00a6',
	section: '\u00a7',
	dieresis: '\u00a8',
	copyright: '\u00a9',
	ordfeminine: '\u00aa',
	guillemotleft: '\u00ab',
	logicalnot: '\u00ac',
	registered: '\u00ae',
	macron: '\u00af',
	degree: '\u00b0',
	plusminus: '\u00b1',
	acute: '\u00b4',
	mu: '\u00b5',
	paragraph: '\u00b6',
	periodcentered: '\u00b7',
	cedilla: '\u00b8',
	ordmasculine: '\u00ba',
	guillemotright: '\u00bb',
	onequarter: '\u00bc',
	onehalf: '\u00bd',
	threequarters: '\u00be',
	questiondown: '\u00bf',
	multiply: '\u00d7',
	divide: '\u00f7',
	germandbls: '\u00df',
	AE: '\u00c6',
	ae: '\u00e6',
	OE: '\u0152',
	oe: '\u0153',
	Oslash: '\u00d8',
	oslash: '\u00f8',
	Lslash: '\u0141',
	lslash: '\u0142',
	dotlessi: '\u0131',
	Eth: '\u00d0',
	eth: '\u00f0',
	Thorn: '\u00de',
	thorn: '\u00fe',
	florin: '\u0192',
	bullet: '\u2022',
	endash: '\u2013',
	emdash: '\u2014',
	ellipsis: '\u2026',
	dagger: '\u2020',
	daggerdbl: '\u2021',
	perthousand: '\u2030',
	quotedblleft: '\u201c',
	quotedblright: '\u201d',
	quotesinglbase: '\u201a',
	quotedblbase: '\u201e',
	guilsinglleft: '\u2039',
	guilsinglright: '\u203a',
	trademark: '\u2122',
	Euro: '\u20ac',
	fraction: '\u2044',
	circumflex: '\u02c6',
	tilde: '\u02dc',
	breve: '\u02d8',
	dotaccent: '\u02d9',
	ring: '\u02da',
	hungarumlaut: '\u02dd',
	ogonek: '\u02db',
	caron: '\u02c7',
	nbspace: '\u00a0',
	sfthyphen: '\u00ad',
	fi: 'fi',
	fl: 'fl',
	ff: 'ff',
	ffi: 'ffi',
	ffl: 'ffl',
	arrowright: '\u2192',
	arrowleft: '\u2190',
	checkmark: '\u2713',
	square: '\u25a0',
	H22073: '\u25a1',
	H18543: '\u25aa',
	H18551: '\u25ab',
	circle: '\u25cb',
	a71: '\u25cf',
	onesuperior: '\u00b9',
	twosuperior: '\u00b2',
	threesuperior: '\u00b3',
};

const ACCENTS: Record<string, string> = {
	acute: '\u0301',
	grave: '\u0300',
	circumflex: '\u0302',
	dieresis: '\u0308',
	tilde: '\u0303',
	ring: '\u030a',
	cedilla: '\u0327',
	caron: '\u030c',
	macron: '\u0304',
	breve: '\u0306',
	dotaccent: '\u0307',
	ogonek: '\u0328',
	hungarumlaut: '\u030b',
	commaaccent: '\u0326',
};

/** The Unicode text a glyph name stands for, or `undefined` for an unknown name. */
export function glyphNameToUnicode(name: string): string | undefined {
	const base = name.split('.')[0]!;
	if (!base) return undefined;
	if (base.includes('_')) {
		const parts = base.split('_').map(glyphNameToUnicode);
		return parts.every((part) => part !== undefined) ? parts.join('') : undefined;
	}
	if (NAMES[base] !== undefined) return NAMES[base];
	if (/^[A-Za-z]$/.test(base)) return base;
	let match = /^uni((?:[0-9A-F]{4})+)$/.exec(base);
	if (match)
		return match[1]!
			.match(/.{4}/g)!
			.map((hex) => String.fromCharCode(Number.parseInt(hex, 16)))
			.join('');
	match = /^u([0-9A-F]{4,6})$/.exec(base);
	if (match) {
		const code = Number.parseInt(match[1]!, 16);
		return code <= 0x10ffff ? String.fromCodePoint(code) : undefined;
	}
	match =
		/^([A-Za-z])(acute|grave|circumflex|dieresis|tilde|ring|cedilla|caron|macron|breve|dotaccent|ogonek|hungarumlaut|commaaccent)$/.exec(
			base,
		);
	if (match) return (match[1]! + ACCENTS[match[2]!]!).normalize('NFC');
	// Producers that name glyphs by code (`g123`, `cid45`, `G4F`) give no text: the caller falls back.
	return undefined;
}
