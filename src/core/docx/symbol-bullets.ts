// List labels in Word's symbol fonts as Unicode text. A small table of the codes used as bullets,
// so the model and layout need not load the full glyph tables of `text/symbol-fonts.ts` (which
// these entries agree with, `symbol-bullets.test.ts`).
import { normaliseFamily } from '../text/symbol-font-names';

/**
 * Bullet codes (the byte of U+F0xx) of each symbol font, each followed by the Unicode picture it
 * draws. Word 16's PDF names Wingdings F0FC U+2713 (the shared table, for PowerPoint, has U+2714).
 */
const BULLETS: Readonly<Record<string, string>> = {
	symbol:
		'*\u2217-\u2212\u00a7\u2663\u00a8\u2666\u00a9\u2665\u00aa\u2660\u00ab\u2194\u00ac\u2190\u00ad\u2191\u00ae\u2192\u00af\u2193\u00b7\u2022',
	wingdings:
		'l\u25cfn\u25a0q\u2751v\u2756\u009f\u2022\u00a1\u25cb\u00a7\u25aa\u00a8\u25fb\u00aa\u2726\u00ab\u2605\u00d8\u27a2\u00e0\u2192\u00e8\u2794\u00f0\u21e8\u00fb\u2717\u00fc\u2713\u00fd\u2612\u00fe\u2611',
};

/** The bullets of `font` by code, or undefined for a font without a table. */
export function symbolBullets(font: string | undefined): Map<number, string> | undefined {
	const pairs = BULLETS[normaliseFamily(font)];
	if (pairs === undefined) return undefined;
	const table = new Map<number, string>();
	for (let at = 0; at < pairs.length; at += 2) table.set(pairs.charCodeAt(at), pairs[at + 1]!);
	return table;
}

/**
 * A list label as Unicode text. Bullets in Word's symbol fonts are codes that select the font's
 * pictures (U+F0xx, or a byte); without those fonts they render as missing glyphs, so each is
 * shown as its Unicode equivalent in `font`: Symbol F02D is a minus sign, F0AE an arrow,
 * Wingdings F0FC a check mark. A private-use bullet in another font is read as Word's bullet
 * library has it (Wingdings, then Symbol), else shown as `•`.
 */
export function displayListLabel(text: string, font?: string): string {
	const own = symbolBullets(font);
	let out = '';
	for (const char of text) {
		const code = char.charCodeAt(0);
		const privateUse = code >= 0xf000 && code <= 0xf0ff;
		const byte = privateUse ? code - 0xf000 : code;
		out +=
			own?.get(byte) ??
			(privateUse
				? (symbolBullets('Wingdings')!.get(byte) ?? symbolBullets('Symbol')!.get(byte) ?? '\u2022')
				: char);
	}
	return out;
}
