/**
 * The font of a PDF's invisible text layer: text drawn as a picture (characters no supplied font
 * has, such as Chinese, Japanese or Korean in a deck that embeds no font for them, scripts that
 * need shaping, whole-slide pictures) is written over the picture in render mode 3 with this
 * font, so it stays selectable, searchable and readable by assistive tools. Its glyphs are empty
 * (`glyphlessTrueType` from `ooxml-core/text`), so it carries no outlines and no third-party
 * font data, and it never needs a font download. Each distinct character gets the next CID, with
 * an East Asian character one em wide and any other half an em, and the ToUnicode map gives each
 * CID back its character. DOM-free.
 */
import { glyphlessTrueType, isEastAsianChar, MAX_GLYPHLESS_GLYPHS } from '../../text/index';

import { toUnicodeCMap, type PdfCidFont, type PdfFontParts } from './pdf-cid-font';

/** Advance of `char` in 1/1000 em. */
const advanceOf = (char: string): number => (isEastAsianChar(char) ? 1000 : 500);

export class PdfTextLayerFont implements PdfCidFont {
	readonly family = 'TextLayer';
	readonly bold = false;
	readonly italic = false;
	readonly #cids = new Map<number, number>();
	readonly #codes: number[] = [];
	/** Advance width of `text` in em units. */
	width(text: string): number {
		let total = 0;
		for (const char of text) total += advanceOf(char) / 1000;
		return total;
	}
	/** `text` as the hex digits of 2-byte CIDs; each new character takes the next CID. */
	encode(text: string): string {
		let out = '';
		for (const char of text) {
			const code = char.codePointAt(0)!;
			let cid = this.#cids.get(code);
			if (cid === undefined) {
				// Past the last glyph a font can hold, characters share CID 0 (they stay invisible).
				cid = this.#codes.length + 1 < MAX_GLYPHLESS_GLYPHS ? this.#codes.length + 1 : 0;
				if (cid) {
					this.#cids.set(code, cid);
					this.#codes.push(code);
				}
			}
			out += cid.toString(16).padStart(4, '0');
		}
		return out;
	}
	get used(): boolean {
		return this.#codes.length > 0;
	}
	build(): PdfFontParts {
		const widths = [0, ...this.#codes.map((code) => advanceOf(String.fromCodePoint(code)))];
		return {
			fontFile: glyphlessTrueType(widths),
			widths,
			toUnicode: toUnicodeCMap(this.#codes),
			descriptor: {
				flags: 4,
				bbox: [0, -200, 1000, 800],
				italicAngle: 0,
				ascent: 800,
				descent: -200,
				capHeight: 700,
				stemV: 80,
			},
		};
	}
}
