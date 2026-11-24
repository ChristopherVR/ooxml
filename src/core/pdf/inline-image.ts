/**
 * Inline images (ISO 32000-1 section 8.9.7): `BI` key-value pairs `ID` data `EI` inside a content
 * stream. The abbreviated keys are expanded so the image reads like an XObject; the data ends
 * where its known length says (unfiltered samples or /L), else at the first `EI` that stands
 * alone between whitespace.
 */
import { decodeFilters } from './filters';
import { EOF, isWhite, PdfKeyword, type PdfLexer } from './lexer';
import { isName, isNumber, PdfName, type PdfDict, type PdfValue } from './objects';
import { indexOf } from './xref';

const KEYS: Record<string, string> = {
	BPC: 'BitsPerComponent',
	CS: 'ColorSpace',
	D: 'Decode',
	DP: 'DecodeParms',
	F: 'Filter',
	H: 'Height',
	IM: 'ImageMask',
	I: 'Interpolate',
	W: 'Width',
	L: 'Length',
};
const COMPONENTS: Record<string, number> = {
	G: 1,
	DeviceGray: 1,
	RGB: 3,
	DeviceRGB: 3,
	CMYK: 4,
	DeviceCMYK: 4,
};

export interface InlineImage {
	dict: PdfDict;
	/** The encoded data between `ID` and `EI`. */
	raw: Uint8Array;
}

/** The byte length of unfiltered samples, when the dictionary tells it. */
function sampleLength(dict: PdfDict): number | undefined {
	const width = dict.get('Width');
	const height = dict.get('Height');
	if (!isNumber(width) || !isNumber(height) || dict.has('Filter')) return undefined;
	const mask = dict.get('ImageMask') === true;
	const space = dict.get('ColorSpace');
	const components = mask
		? 1
		: isName(space)
			? COMPONENTS[space.name]
			: Array.isArray(space) && (isName(space[0], 'I') || isName(space[0], 'Indexed'))
				? 1
				: undefined;
	const bpc = mask ? 1 : dict.get('BitsPerComponent');
	if (components === undefined || !isNumber(bpc)) return undefined;
	return Math.ceil((width * components * bpc) / 8) * height;
}

/** Whether `EI` stands alone at `at` (after optional whitespace). */
function endsAt(bytes: Uint8Array, at: number): number | undefined {
	let pos = at;
	while (isWhite(bytes[pos])) pos++;
	if (bytes[pos] !== 0x45 || bytes[pos + 1] !== 0x49) return undefined;
	const after = bytes[pos + 2];
	return after === undefined || isWhite(after) || after === 0x51 /* Q */ ? pos : undefined;
}

/** Reads an inline image after its `BI` operator, leaving the lexer after `EI`. */
export function readInlineImage(lexer: PdfLexer): InlineImage | undefined {
	const dict: PdfDict = new Map();
	for (let guard = 0; ; guard++) {
		const key = lexer.object();
		if (key === EOF || guard > 200) return undefined;
		if (key instanceof PdfKeyword) {
			if (key.word === 'ID') break;
			continue;
		}
		if (key instanceof PdfName) {
			const value = lexer.object();
			if (value !== EOF && !(value instanceof PdfKeyword))
				dict.set(KEYS[key.name] ?? key.name, value as PdfValue);
		}
	}
	const { bytes } = lexer;
	const start = Math.min(bytes.length, lexer.pos + 1);
	const declared = dict.get('Length');
	const known = isNumber(declared) && declared >= 0 ? declared : sampleLength(dict);
	if (known !== undefined && start + known <= bytes.length) {
		const at = endsAt(bytes, start + known);
		if (at !== undefined) {
			lexer.pos = at + 2;
			return { dict, raw: bytes.subarray(start, start + known) };
		}
	}
	for (let end = start; ; end += 2) {
		end = indexOf(bytes, 'EI', end);
		if (end < 0) {
			lexer.pos = bytes.length;
			return undefined;
		}
		if (isWhite(bytes[end - 1]) && (end + 2 >= bytes.length || isWhite(bytes[end + 2]))) {
			lexer.pos = end + 2;
			return { dict, raw: bytes.subarray(start, end - 1) };
		}
	}
}

/** The inline image's data with its non-image filters undone. */
export function inlineData(image: InlineImage, maxBytes: number) {
	return decodeFilters(image.raw, image.dict, maxBytes);
}
