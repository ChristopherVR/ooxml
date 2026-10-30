/**
 * A TrueType font embedded in a PDF as a subset Type0 / CIDFontType2 font with a ToUnicode CMap,
 * so text outside WinAnsi (any script, a minus sign) stays real, selectable text. Characters get
 * sequential CIDs in the order they are first used; the subset is cut when the PDF is built.
 * DOM-free.
 */
import {
	parseTrueType,
	subsetTrueType,
	type TrueTypeFont,
	type TrueTypeSubset,
} from 'ooxml-core/text';

/** A font file the host supplies for embedding. */
export interface PdfFontFile {
	/** CSS family name the document's text asks for (case-insensitive). */
	family: string;
	bold?: boolean;
	italic?: boolean;
	/** TrueType bytes (`.ttf`, or a `.ttc` whose first font is used). */
	data: Uint8Array;
}

export class PdfEmbeddedFont {
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
		// fsType bit 1 is restricted-licence embedding, bit 8 forbids subsetting.
		if (!font || font.embeddingFlags & 0x0002 || font.embeddingFlags & 0x0100) return undefined;
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

export interface PdfFontParts {
	fontFile: Uint8Array;
	/** Advance per CID in 1/1000 em, starting at CID 0. */
	widths: number[];
	toUnicode: string;
	descriptor: {
		flags: number;
		bbox: [number, number, number, number];
		italicAngle: number;
		ascent: number;
		descent: number;
		capHeight: number;
		stemV: number;
	};
}

const hex4 = (value: number) => value.toString(16).padStart(4, '0');
const utf16 = (code: number) => {
	if (code < 0x10000) return hex4(code);
	const v = code - 0x10000;
	return hex4(0xd800 + (v >> 10)) + hex4(0xdc00 + (v & 0x3ff));
};

/** A ToUnicode CMap sending CID `index + 1` to `codes[index]`. */
function toUnicodeCMap(codes: readonly number[]): string {
	const lines = [
		'/CIDInit /ProcSet findresource begin',
		'12 dict begin',
		'begincmap',
		'/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def',
		'/CMapName /Adobe-Identity-UCS def',
		'/CMapType 2 def',
		'1 begincodespacerange',
		'<0000> <FFFF>',
		'endcodespacerange',
	];
	for (let start = 0; start < codes.length; start += 100) {
		const chunk = codes.slice(start, start + 100);
		lines.push(`${chunk.length} beginbfchar`);
		chunk.forEach((code, index) => lines.push(`<${hex4(start + index + 1)}> <${utf16(code)}>`));
		lines.push('endbfchar');
	}
	lines.push('endcmap', 'CMapName currentdict /CMap defineresource pop', 'end', 'end');
	return lines.join('\n');
}

/**
 * The embedded fonts for a document: matched by family and style, with every other font as a
 * fallback for characters the match lacks.
 */
export class PdfFontSet {
	readonly fonts: PdfEmbeddedFont[];
	constructor(files: readonly PdfFontFile[] = []) {
		this.fonts = files.flatMap((file) => {
			const font = PdfEmbeddedFont.create(file);
			return font ? [font] : [];
		});
	}
	get size(): number {
		return this.fonts.length;
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
