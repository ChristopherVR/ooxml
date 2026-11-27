/**
 * What the PDF writer needs from a font it writes as Type0 / CIDFontType2 with 2-byte CIDs and a
 * ToUnicode CMap: a subset of a real TrueType font (`pdf-embedded-font.ts`) or the glyphless
 * font of the invisible text layer (`pdf-text-layer-font.ts`). DOM-free.
 */

/** A font the PDF writes as Type0 / CIDFontType2 with 2-byte CIDs (`PdfVectorPage.embeddedText`). */
export interface PdfCidFont {
	readonly family: string;
	readonly bold: boolean;
	readonly italic: boolean;
	/** Advance width of `text` in em units. */
	width(text: string): number;
	/** `text` as the hex digits of 2-byte CIDs (no angle brackets); marks the glyphs as used. */
	encode(text: string): string;
	/** Everything the PDF needs to write the font objects. */
	build(): PdfFontParts;
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
export function toUnicodeCMap(codes: readonly number[]): string {
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
