/**
 * PDF fonts as the text reader needs them (ISO 32000-1 section 9.6 to 9.10): how a string splits
 * into character codes, each code's Unicode text and advance width, and what the font is called
 * (family, bold, italic). Font programs themselves are never read.
 */
import { PdfCMap } from './cmap';
import type { PdfDocument } from './document';
import { baseEncodingChar, glyphNameToUnicode, type BaseEncoding } from './encodings';
import { standardWidth } from './standard-metrics';
import {
	decodeTextString,
	isDict,
	isName,
	isNumber,
	PdfName,
	PdfStream,
	type PdfDict,
	type PdfValue,
} from './objects';

export interface PdfFontInfo {
	/** The font's name without a subset tag, e.g. `Arial-BoldMT`. */
	name: string;
	/** A family name as Word would show it, e.g. `Arial`. */
	family: string;
	bold: boolean;
	italic: boolean;
	serif: boolean;
	monospace: boolean;
}

export interface PdfGlyph {
	code: number;
	/** Byte width of the code (for `Tw`, which applies only to the single-byte code 32). */
	bytes: number;
	text: string;
	/** Advance in text space units (thousandths of the font size for most fonts). */
	width: number;
}

const KNOWN_FAMILIES: Record<string, string> = {
	arialmt: 'Arial',
	arial: 'Arial',
	timesnewromanpsmt: 'Times New Roman',
	timesnewroman: 'Times New Roman',
	'times-roman': 'Times New Roman',
	times: 'Times New Roman',
	helvetica: 'Arial',
	couriernewpsmt: 'Courier New',
	couriernew: 'Courier New',
	courier: 'Courier New',
	symbolmt: 'Symbol',
	cambriamath: 'Cambria Math',
};

/** A family name from a PostScript name: `TimesNewRomanPS-BoldItalicMT` gives `Times New Roman`. */
export function familyName(name: string): string {
	const base = name.replace(/^[A-Z]{6}\+/, '').split(/[-,]/)[0]!;
	const known = KNOWN_FAMILIES[base.toLowerCase()];
	if (known) return known;
	const cleaned = base.replace(/(PS)?MT$/, '').replace(/PS$/, '');
	return (
		cleaned
			.replace(/([a-z])([A-Z])/g, '$1 $2')
			.replace(/([A-Z]+)([A-Z][a-z])/g, '$1 $2')
			.trim() || 'Calibri'
	);
}

export class PdfFont {
	readonly #widths = new Map<number, number>();
	#defaultWidth = 500;
	#encoding = new Map<number, string>();
	#base: BaseEncoding | undefined;
	#toUnicode: PdfCMap | undefined;
	#cmap: PdfCMap | undefined;
	#symbolic = false;
	/** Type 3 glyph space to text space. */
	scale = 0.001;
	readonly vertical: boolean;

	private constructor(
		readonly info: PdfFontInfo,
		readonly composite: boolean,
		vertical: boolean,
	) {
		this.vertical = vertical;
	}

	static async load(document: PdfDocument, dict: PdfDict): Promise<PdfFont> {
		const subtype = (await document.resolve(dict.get('Subtype'))) as PdfValue;
		const composite = isName(subtype, 'Type0');
		const baseName = await document.resolve(dict.get('BaseFont'));
		const name = (baseName instanceof PdfName ? baseName.name : 'Unknown').replace(
			/^[A-Z]{6}\+/,
			'',
		);
		let descriptorSource = dict;
		let descendant: PdfDict | undefined;
		if (composite) {
			const kids = await document.resolve(dict.get('DescendantFonts'));
			descendant = Array.isArray(kids) ? await document.dict(kids[0]) : undefined;
			if (descendant) descriptorSource = descendant;
		}
		const descriptor = await document.dict(descriptorSource.get('FontDescriptor'));
		const flagsValue = await document.resolve(descriptor?.get('Flags'));
		const flags = isNumber(flagsValue) ? flagsValue : 0;
		const weight = await document.resolve(descriptor?.get('FontWeight'));
		const angle = await document.resolve(descriptor?.get('ItalicAngle'));
		const declaredFamily = decodeTextString(
			await document.resolve(descriptor?.get('FontFamily')),
		)?.trim();
		const lower = name.toLowerCase();
		const info: PdfFontInfo = {
			name,
			family: declaredFamily || familyName(name),
			bold:
				/bold|black|heavy|semibold|demi/.test(lower) ||
				(isNumber(weight) && weight >= 600) ||
				(flags & 0x40000) !== 0,
			italic:
				/italic|oblique/.test(lower) || (isNumber(angle) && angle !== 0) || (flags & 0x40) !== 0,
			serif: (flags & 2) !== 0 || /times|serif|georgia|garamond|cambria|book/.test(lower),
			monospace: (flags & 1) !== 0 || /courier|mono|consolas/.test(lower),
		};
		const encodingValue = await document.resolve(dict.get('Encoding'));
		const font = new PdfFont(info, composite, isName(encodingValue, 'Identity-V'));
		font.#symbolic = (flags & 4) !== 0 && (flags & 32) === 0;
		const toUnicode = await document.resolve(dict.get('ToUnicode'));
		if (toUnicode instanceof PdfStream) {
			try {
				font.#toUnicode = PdfCMap.parse(await document.streamBytes(toUnicode));
			} catch {
				font.#toUnicode = undefined;
			}
		}
		if (composite) await font.#loadComposite(document, encodingValue, descendant);
		else await font.#loadSimple(document, dict, subtype, encodingValue, name);
		return font;
	}

	async #loadComposite(
		document: PdfDocument,
		encoding: PdfValue | undefined,
		descendant: PdfDict | undefined,
	) {
		if (encoding instanceof PdfStream) {
			try {
				this.#cmap = PdfCMap.parse(await document.streamBytes(encoding));
			} catch {
				this.#cmap = PdfCMap.identityH();
			}
		} else this.#cmap = PdfCMap.identityH();
		const dw = await document.resolve(descendant?.get('DW'));
		this.#defaultWidth = isNumber(dw) ? dw : 1000;
		const w = await document.resolve(descendant?.get('W'));
		if (!Array.isArray(w)) return;
		for (let at = 0; at < w.length;) {
			const first = await document.resolve(w[at]);
			const second = await document.resolve(w[at + 1]);
			if (!isNumber(first)) break;
			if (Array.isArray(second)) {
				for (let index = 0; index < second.length; index++) {
					const width = await document.resolve(second[index]);
					if (isNumber(width)) this.#widths.set(first + index, width);
				}
				at += 2;
			} else {
				const width = await document.resolve(w[at + 2]);
				if (isNumber(second) && isNumber(width))
					for (let cid = first; cid <= second && cid - first < 0x10000; cid++)
						this.#widths.set(cid, width);
				at += 3;
			}
		}
	}

	async #loadSimple(
		document: PdfDocument,
		dict: PdfDict,
		subtype: PdfValue,
		encoding: PdfValue | undefined,
		name: string,
	) {
		const isType3 = isName(subtype, 'Type3');
		if (isType3) {
			const matrix = await document.resolve(dict.get('FontMatrix'));
			if (Array.isArray(matrix) && isNumber(matrix[0])) this.scale = matrix[0];
		}
		const first = await document.resolve(dict.get('FirstChar'));
		const widths = await document.resolve(dict.get('Widths'));
		if (isNumber(first) && Array.isArray(widths))
			for (let index = 0; index < widths.length; index++) {
				const width = await document.resolve(widths[index]);
				if (isNumber(width)) this.#widths.set(first + index, width);
			}
		const descriptor = await document.dict(dict.get('FontDescriptor'));
		const missing = await document.resolve(descriptor?.get('MissingWidth'));
		this.#defaultWidth = isNumber(missing) && missing > 0 ? missing : this.#widths.size ? 0 : 500;
		this.standard = !this.#widths.size ? name : undefined;
		let base: BaseEncoding | undefined;
		let differences: PdfValue[] = [];
		if (isName(encoding)) base = encoding.name as BaseEncoding;
		else if (isDict(encoding)) {
			const baseName = await document.resolve(encoding.get('BaseEncoding'));
			if (isName(baseName)) base = baseName.name as BaseEncoding;
			const diff = await document.resolve(encoding.get('Differences'));
			if (Array.isArray(diff)) differences = diff;
		}
		const symbolFont = /symbol|dingbat|wingding/i.test(name);
		this.#base =
			base ??
			(symbolFont || this.#symbolic
				? undefined
				: isName(subtype, 'TrueType')
					? 'WinAnsiEncoding'
					: 'StandardEncoding');
		let code = 0;
		for (const item of differences) {
			if (isNumber(item)) code = item;
			else if (item instanceof PdfName) {
				const text = glyphNameToUnicode(item.name);
				if (text !== undefined) this.#encoding.set(code, text);
				code++;
			}
		}
	}

	/** Set for a standard 14 font without widths: advances come from its metrics. */
	standard: string | undefined;

	/** The glyphs of a string. */
	glyphs(bytes: Uint8Array): PdfGlyph[] {
		const codes = this.composite
			? this.#cmap!.codes(bytes)
			: Array.from(bytes, (code) => ({ code, bytes: 1 }));
		return codes.map(({ code, bytes: size }) => {
			const key = this.composite ? this.#cmap!.cidOf(code) : code;
			const text = this.#text(code);
			let width = this.#widths.get(key);
			if (width === undefined)
				width = this.standard ? standardWidth(this.standard, text) : this.#defaultWidth;
			return { code, bytes: size, text, width: width * this.scale };
		});
	}

	#text(code: number): string {
		const mapped = this.#toUnicode?.unicode.get(code);
		if (mapped !== undefined) return mapped;
		if (this.composite) return '';
		const differ = this.#encoding.get(code);
		if (differ !== undefined) return differ;
		if (this.#base) return baseEncodingChar(this.#base, code) ?? '';
		// A symbolic TrueType font maps its codes straight to the font's (often private) cmap.
		if (code >= 0x20 && code < 0x7f) return String.fromCharCode(code);
		return '';
	}
}
