// List labels in Word's symbol fonts as Unicode text. A small table of the codes used as bullets,
// so the model and layout need not load the full glyph tables of `text/symbol-fonts.ts` (which
// these entries agree with, `symbol-bullets.test.ts`).
import { normaliseFamily } from '../text/symbol-font-names';

/** Bullet codes (the byte of U+F0xx) of each symbol font and the Unicode picture they draw. */
const BULLETS: Readonly<Record<string, Readonly<Record<number, string>>>> = {
	symbol: {
		0x2a: '∗',
		0x2d: '−',
		0xa7: '♣',
		0xa8: '♦',
		0xa9: '♥',
		0xaa: '♠',
		0xab: '↔',
		0xac: '←',
		0xad: '↑',
		0xae: '→',
		0xaf: '↓',
		0xb0: '°',
		0xb1: '±',
		0xb3: '≥',
		0xb4: '×',
		0xb7: '•',
	},
	wingdings: {
		0x6c: '●',
		0x6e: '■',
		0x71: '❑',
		0x76: '❖',
		0x9e: '·',
		0x9f: '•',
		0xa1: '○',
		0xa7: '▪',
		0xa8: '◻',
		0xaa: '✦',
		0xab: '★',
		0xd8: '➢',
		0xe0: '→',
		0xe8: '➔',
		0xf0: '⇨',
		0xfb: '✗',
		// Word 16's PDF names it U+2713 (the shared table, for PowerPoint, has U+2714).
		0xfc: '✓',
		0xfd: '☒',
		0xfe: '☑',
	},
};

/** Bullets of Word's bullet library by code, for a label whose font is not a symbol font. */
const LIBRARY: Readonly<Record<number, string>> = {
	0xb7: '•',
	0xa7: '▪',
	0xd8: '➢',
	0x76: '❖',
	0xfc: '✓',
	0x6e: '■',
	0x71: '❑',
	0xe0: '➩',
	0xb2: '○',
};

/** The table `displayListLabel` reads for `font`, for tests. */
export function symbolBulletTable(font: string): Readonly<Record<number, string>> | undefined {
	return BULLETS[normaliseFamily(font)];
}

/**
 * A list label as Unicode text. Bullets in Word's symbol fonts are codes that select the font's
 * pictures (U+F0xx, or a byte); without those fonts they render as missing glyphs, so each is
 * shown as its Unicode equivalent in `font`: Symbol F02D is a minus sign, F0AE an arrow,
 * Wingdings F0FC a check mark. An unknown private-use bullet shows as `•`.
 */
export function displayListLabel(text: string, font?: string): string {
	const table = font === undefined ? undefined : BULLETS[normaliseFamily(font)];
	let out = '';
	for (const char of text) {
		const code = char.codePointAt(0)!;
		const privateUse = code >= 0xf000 && code <= 0xf0ff;
		const byte = privateUse ? code - 0xf000 : code;
		const named = table && byte > 0x20 && byte <= 0xff ? table[byte] : undefined;
		out += named ?? (privateUse ? (LIBRARY[byte] ?? '•') : char);
	}
	return out;
}
