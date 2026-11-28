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
} from '../../text/index';
import { pairKerning, type PairKerning } from '../../text/truetype/kerning';
import { toUnicodeCMap, type PdfCidFont, type PdfFontParts } from './pdf-cid-font';
import { PdfTextLayerFont } from './pdf-text-layer-font';

export type { PdfCidFont, PdfFontParts } from './pdf-cid-font';

/**
 * The glyph a symbol font's file draws for a code, its own or the picture a Unicode equivalent
 * stands for (`symbolFontGlyph` of `text/symbol-fonts`, passed in so the glyph tables load only
 * where they are used).
 */
export type PdfSymbolGlyph = (
	font: Pick<TrueTypeFont, 'glyphFor' | 'symbolEncoded'>,
	family: string,
	code: number,
) => number;

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
		readonly symbolGlyph?: PdfSymbolGlyph,
	) {
		this.family = file.family;
		// The host's flag, or the font's own weight (a bold face the host did not flag).
		this.bold = !!file.bold || (font.weightClass ?? 400) >= 600;
		// The host's flag, or the font's own slant (a non-zero italic angle in its `post` table).
		this.italic = !!file.italic || font.italicAngle !== 0;
	}
	/** `undefined` when the bytes are not TrueType or the font's licence forbids embedding. */
	static create(file: PdfFontFile, symbolGlyph?: PdfSymbolGlyph): PdfEmbeddedFont | undefined {
		const font = parseTrueType(file.data);
		// A PDF carries a subset of the outlines: restricted, no-subsetting and bitmap-only faces
		// may not be embedded that way.
		if (!font || !mayEmbedFontSubset(embeddingPermission(font.embeddingFlags))) return undefined;
		return new PdfEmbeddedFont(file, font, symbolGlyph);
	}
	/** Whether the font has a glyph for every character of `text` (controls excepted). */
	covers(text: string): boolean {
		for (const char of text) {
			const code = char.codePointAt(0)!;
			if (code < 32 || code === 0xa0) continue;
			if (!this.glyphOf(code)) return false;
		}
		return true;
	}
	/** Advance width of `text` in em units. */
	width(text: string): number {
		let total = 0;
		for (const char of text) total += this.font.advance(this.#glyph(char)) / this.font.unitsPerEm;
		return total;
	}
	#pairs: PairKerning | undefined | null = null;
	/** The pair adjustment between two characters in em units (negative pulls them together). */
	kern(left: string, right: string): number {
		if (this.#pairs === null) this.#pairs = pairKerning(this.font);
		const first = this.#glyph(left);
		const second = this.#glyph(right);
		if (!this.#pairs || !first || !second) return 0;
		return this.#pairs(first, second) / this.font.unitsPerEm;
	}
	#glyph(char: string): number {
		const code = char.codePointAt(0)!;
		return this.glyphOf(code === 0xa0 || code < 32 ? 32 : code);
	}
	/** A symbol font also draws the Unicode equivalents of its pictures (a list's check mark). */
	glyphOf(code: number): number {
		return this.symbolGlyph
			? this.symbolGlyph(this.font, this.family, code)
			: this.font.glyphFor(code);
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
	/** The supplied family that stands in for a family no file is named for, if any. */
	readonly #substitute: ((family: string) => string | undefined) | undefined;
	constructor(
		files: readonly PdfFontFile[] = [],
		options: {
			embedAll?: boolean;
			substitute?: (family: string) => string | undefined;
			symbolGlyph?: PdfSymbolGlyph;
		} = {},
	) {
		this.embedAll = !!options.embedAll;
		this.#substitute = options.substitute;
		this.fonts = files.flatMap((file) => {
			const font = PdfEmbeddedFont.create(file, options.symbolGlyph);
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
		const names = this.#names(family);
		const score = (font: PdfEmbeddedFont) =>
			(names.includes(font.family.toLowerCase()) ? 4 : 0) +
			(font.bold === bold ? 2 : 0) +
			(font.italic === italic ? 1 : 0);
		// A symbol font "covers" Latin text with its pictures: it writes only text asked of it.
		return [...this.fonts]
			.filter((font) => !font.font.symbolEncoded || names.includes(font.family.toLowerCase()))
			.sort((a, b) => score(b) - score(a))
			.find((font) => font.covers(text));
	}
	/** Whether text in `family` should be written with a supplied font when one covers it. */
	prefers(family: string): boolean {
		return this.embedAll ? this.fonts.length > 0 : this.has(family);
	}
	/** Whether a font was asked for by name (so a standard font would be a substitution). */
	has(family: string): boolean {
		const names = this.#names(family);
		return this.fonts.some((font) => names.includes(font.family.toLowerCase()));
	}
	/** The lower-case names `family` lists, and the supplied family substituted for it. */
	#names(family: string): string[] {
		const names = family.split(',').map((name) =>
			name
				.trim()
				.replace(/^['"]|['"]$/g, '')
				.toLowerCase(),
		);
		const substitute = this.#substitute?.(family);
		return substitute ? [...names, substitute.toLowerCase()] : names;
	}
}
