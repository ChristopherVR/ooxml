/**
 * A TrueType font embedded in a PDF as a subset Type0 / CIDFontType2 font with a ToUnicode CMap,
 * so text outside WinAnsi (any script, a minus sign) stays real, selectable text. Characters get
 * sequential CIDs in the order they are first used; the subset is cut when the PDF is built.
 * DOM-free.
 */
import {
	embeddingPermission,
	mayEmbedFontSubset,
	parseTrueType,
	subsetTrueType,
	type TrueTypeFont,
	type TrueTypeSubset,
} from 'ooxml-core/text';
import { toUnicodeCMap, type PdfCidFont, type PdfFontParts } from './pdf-cid-font';
import { PdfTextLayerFont } from './pdf-text-layer-font';

export type { PdfCidFont, PdfFontParts } from './pdf-cid-font';

/** A font file the host supplies for embedding. */
export interface PdfFontFile {
	/** CSS family name the document's text asks for (case-insensitive). */
	family: string;
	bold?: boolean;
	italic?: boolean;
	/** TrueType bytes (`.ttf`, or a `.ttc` whose first font is used). */
	data: Uint8Array;
}

export class PdfEmbeddedFont implements PdfCidFont {
	readonly #cids = new Map<number, number>();
	readonly #used: number[] = [];
	readonly #codes: number[] = [];
	readonly family: string;
	readonly bold: boolean;
	readonly italic: boolean;
	private constructor(
		file: PdfFontFile,
		readonly font: TrueTypeFont,
	) {
		this.family = file.family;
		this.bold = !!file.bold;
		this.italic = !!file.italic;
	}
	/** `undefined` when the bytes are not TrueType or the font's licence forbids embedding. */
	static create(file: PdfFontFile): PdfEmbeddedFont | undefined {
		const font = parseTrueType(file.data);
		// A PDF carries a subset of the outlines: restricted, no-subsetting and bitmap-only faces
		// may not be embedded that way.
		if (!font || !mayEmbedFontSubset(embeddingPermission(font.embeddingFlags))) return undefined;
		return new PdfEmbeddedFont(file, font);
	}
	/** Whether the font has a glyph for every character of `text` (controls excepted). */
	covers(text: string): boolean {
		for (const char of text) {
			const code = char.codePointAt(0)!;
			if (code < 32 || code === 0xa0) continue;
			if (!this.font.glyphFor(code)) return false;
		}
		return true;
	}
	/** Advance width of `text` in em units. */
	width(text: string): number {
		let total = 0;
		for (const char of text) total += this.font.advance(this.#glyph(char)) / this.font.unitsPerEm;
		return total;
	}
	#glyph(char: string): number {
		const code = char.codePointAt(0)!;
		return this.font.glyphFor(code === 0xa0 || code < 32 ? 32 : code);
	}
	/** `text` as the hex digits of 2-byte CIDs (no angle brackets); marks the glyphs as used. */
	encode(text: string): string {
		let out = '';
		for (const char of text) {
			const glyph = this.#glyph(char);
			let cid = this.#cids.get(glyph);
			if (cid === undefined) {
				cid = this.#used.length + 1;
				this.#cids.set(glyph, cid);
				this.#used.push(glyph);
				this.#codes.push(char === ' ' ? 32 : char.codePointAt(0)!);
			}
			out += cid.toString(16).padStart(4, '0');
		}
		return out;
	}
	get used(): boolean {
		return this.#used.length > 0;
	}
	/** The hex code of one space glyph, for word spacing. */
	get spaceCode(): string {
		return this.encode(' ');
	}
	/** Everything the PDF needs to write the font objects. */
	build(): PdfFontParts {
		const subset: TrueTypeSubset = subsetTrueType(this.font, this.#used);
		const scale = 1000 / this.font.unitsPerEm;
		const widths = subset.glyphs.map((glyph) => Math.round(this.font.advance(glyph) * scale));
		const units = (value: number) => Math.round(value * scale);
		const [x0, y0, x1, y1] = this.font.bbox;
		const ascent = units(this.font.ascent);
		return {
			fontFile: subset.data,
			widths,
			toUnicode: toUnicodeCMap(this.#codes),
			descriptor: {
				flags: 4 | (this.italic ? 64 : 0) | (this.bold ? 262144 : 0),
				bbox: [units(x0), units(y0), units(x1), units(y1)],
				italicAngle: this.italic && !this.font.italicAngle ? -12 : this.font.italicAngle,
				ascent,
				descent: units(this.font.descent),
				capHeight: this.font.capHeight ? units(this.font.capHeight) : Math.round(ascent * 0.7),
				stemV: this.bold ? 140 : 80,
			},
		};
	}
}

/**
 * The embedded fonts for a document: matched by family and style, with every other font as a
 * fallback for characters the match lacks.
 */
export class PdfFontSet {
	readonly fonts: PdfEmbeddedFont[];
	#textLayer: PdfTextLayerFont | undefined;
	/**
	 * Write text in a family no supplied font is named for with a supplied font that covers it,
	 * rather than an unembedded standard PDF font (PDF/A needs every font embedded).
	 */
	readonly embedAll: boolean;
	constructor(files: readonly PdfFontFile[] = [], options: { embedAll?: boolean } = {}) {
		this.embedAll = !!options.embedAll;
		this.fonts = files.flatMap((file) => {
			const font = PdfEmbeddedFont.create(file);
			return font ? [font] : [];
		});
	}
	get size(): number {
		return this.fonts.length;
	}
	/**
	 * The glyphless font of this document's invisible text layer, for text over a picture of it
	 * that no supplied font can write (one font for every page, created on first use).
	 */
	get textLayer(): PdfTextLayerFont {
		this.#textLayer ??= new PdfTextLayerFont();
		return this.#textLayer;
	}
	/** The best font that has every character of `text`, or `undefined`. */
	pick(family: string, bold: boolean, italic: boolean, text: string): PdfEmbeddedFont | undefined {
		const names = family.split(',').map((name) =>
			name
				.trim()
				.replace(/^['"]|['"]$/g, '')
				.toLowerCase(),
		);
		const score = (font: PdfEmbeddedFont) =>
			(names.includes(font.family.toLowerCase()) ? 4 : 0) +
			(font.bold === bold ? 2 : 0) +
			(font.italic === italic ? 1 : 0);
		return [...this.fonts].sort((a, b) => score(b) - score(a)).find((font) => font.covers(text));
	}
	/** Whether text in `family` should be written with a supplied font when one covers it. */
	prefers(family: string): boolean {
		return this.embedAll ? this.fonts.length > 0 : this.has(family);
	}
	/** Whether a font was asked for by name (so a standard font would be a substitution). */
	has(family: string): boolean {
		const names = family.split(',').map((name) =>
			name
				.trim()
				.replace(/^['"]|['"]$/g, '')
				.toLowerCase(),
		);
		return this.fonts.some((font) => names.includes(font.family.toLowerCase()));
	}
}
