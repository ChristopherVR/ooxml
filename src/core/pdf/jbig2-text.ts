/**
 * The JBIG2 text region decoding procedure (ITU-T T.88 6.4): symbol instances placed in strips,
 * arithmetic or Huffman coded, optionally refined. Symbol dictionaries reuse it for refinement
 * and aggregate coding (6.5.8.2).
 */
import {
	ArithmeticDecoder,
	decodeId,
	decodeInteger,
	idContexts,
	integerContexts,
} from './jbig2-arith';
import { Bitmap, compose, Jbig2Error, type Jbig2Budget } from './jbig2-bitmap';
import { decodeRefinement, refinementContexts, type At } from './jbig2-generic';
import { BitReader, decodeHuffman, type HuffmanTable } from './jbig2-huffman';

/** The arithmetic integer and symbol ID decoders a text region uses (shared within a dictionary). */
export class TextContexts {
	readonly dt = integerContexts();
	readonly fs = integerContexts();
	readonly ds = integerContexts();
	readonly it = integerContexts();
	readonly ri = integerContexts();
	readonly rdw = integerContexts();
	readonly rdh = integerContexts();
	readonly rdx = integerContexts();
	readonly rdy = integerContexts();
	readonly id: Uint8Array;
	readonly refinement: Uint8Array;
	constructor(
		readonly codeLength: number,
		refinementTemplate: number,
	) {
		if (codeLength > 24) throw new Jbig2Error('A JBIG2 symbol code is too long.');
		this.id = idContexts(codeLength);
		this.refinement = refinementContexts(refinementTemplate);
	}
}

export interface HuffmanTextTables {
	fs: HuffmanTable;
	ds: HuffmanTable;
	dt: HuffmanTable;
	rdw: HuffmanTable;
	rdh: HuffmanTable;
	rdx: HuffmanTable;
	rdy: HuffmanTable;
	rsize: HuffmanTable;
	/** The symbol ID table, or the bit length of fixed-length IDs (symbol dictionaries). */
	id: HuffmanTable | number;
}

export interface TextRegion {
	width: number;
	height: number;
	instances: number;
	strips: number;
	symbols: Bitmap[];
	defaultPixel: number;
	operator: number;
	transposed: boolean;
	/** 0 bottom left, 1 top left, 2 bottom right, 3 top right. */
	corner: number;
	dsOffset: number;
	refine: boolean;
	refinementTemplate: number;
	refinementAt: At[];
}

/** Either the arithmetic decoder and its contexts or the Huffman reader and its tables. */
export type TextCoding =
	| { kind: 'arith'; decoder: ArithmeticDecoder; contexts: TextContexts }
	| { kind: 'huffman'; reader: BitReader; tables: HuffmanTextTables; refinement: Uint8Array };

const value = (decoded: number | null): number => {
	if (decoded === null) throw new Jbig2Error('The JBIG2 text region has an unexpected OOB.');
	return decoded;
};

/** Decodes a text region (6.4.5). */
export function decodeText(region: TextRegion, coding: TextCoding, budget: Jbig2Budget): Bitmap {
	const out = new Bitmap(region.width, region.height, budget, region.defaultPixel);
	const { strips, symbols, transposed, corner } = region;
	const logStrips = Math.log2(strips);
	const arith = coding.kind === 'arith' ? coding : undefined;
	const huff = coding.kind === 'huffman' ? coding : undefined;
	const integer = (key: 'dt' | 'fs' | 'ds') =>
		arith
			? decodeInteger(arith.decoder, arith.contexts[key])
			: decodeHuffman(huff!.reader, huff!.tables[key]);
	let stripT = -value(integer('dt')) * strips;
	let firstS = 0;
	let instances = 0;
	while (instances < region.instances) {
		stripT += value(integer('dt')) * strips;
		let curS = 0;
		// A strip ends with an OOB, coded even after the last instance (6.4.5 step 3c ii).
		for (let first = true; ; first = false) {
			if (first) {
				firstS += value(integer('fs'));
				curS = firstS;
			} else {
				const ds = integer('ds');
				if (ds === null) break;
				if (instances >= region.instances) return out;
				curS += ds + region.dsOffset;
			}
			const curT =
				strips === 1
					? 0
					: arith
						? value(decodeInteger(arith.decoder, arith.contexts.it))
						: huff!.reader.bits(logStrips);
			const t = stripT + curT;
			const id = arith
				? decodeId(arith.decoder, arith.contexts.id, arith.contexts.codeLength)
				: typeof huff!.tables.id === 'number'
					? huff!.reader.bits(huff!.tables.id)
					: value(decodeHuffman(huff!.reader, huff!.tables.id));
			let symbol = symbols[id];
			if (!symbol) throw new Jbig2Error('A JBIG2 text region refers to a missing symbol.');
			const refine = region.refine
				? arith
					? value(decodeInteger(arith.decoder, arith.contexts.ri))
					: huff!.reader.bit()
				: 0;
			if (refine) symbol = refined(region, coding, symbol, budget);
			const wi = symbol.width;
			const hi = symbol.height;
			if (!transposed && (corner === 2 || corner === 3)) curS += wi - 1;
			else if (transposed && (corner === 0 || corner === 2)) curS += hi - 1;
			// The reference corner of the instance is at (S, T), or (T, S) when transposed.
			let x = transposed ? t : curS;
			let y = transposed ? curS : t;
			if (corner === 2 || corner === 3) x -= wi - 1;
			if (corner === 0 || corner === 2) y -= hi - 1;
			compose(out, symbol, x, y, region.operator, budget);
			if (!transposed && (corner === 0 || corner === 1)) curS += wi - 1;
			else if (transposed && (corner === 1 || corner === 3)) curS += hi - 1;
			instances++;
		}
	}
	return out;
}

/** Decodes the refinement of one symbol instance (6.4.11). */
function refined(
	region: TextRegion,
	coding: TextCoding,
	symbol: Bitmap,
	budget: Jbig2Budget,
): Bitmap {
	const read = (key: 'rdw' | 'rdh' | 'rdx' | 'rdy') =>
		value(
			coding.kind === 'arith'
				? decodeInteger(coding.decoder, coding.contexts[key])
				: decodeHuffman(coding.reader, coding.tables[key]),
		);
	const rdw = read('rdw');
	const rdh = read('rdh');
	const rdx = read('rdx');
	const rdy = read('rdy');
	const options = {
		width: symbol.width + rdw,
		height: symbol.height + rdh,
		template: region.refinementTemplate,
		reference: symbol,
		dx: Math.floor(rdw / 2) + rdx,
		dy: Math.floor(rdh / 2) + rdy,
		at: region.refinementAt,
		tpgron: false,
	};
	if (options.width < 0 || options.height < 0)
		throw new Jbig2Error('A JBIG2 refined symbol has a negative size.');
	if (coding.kind === 'arith')
		return decodeRefinement(coding.decoder, coding.contexts.refinement, options, budget);
	const size = value(decodeHuffman(coding.reader, coding.tables.rsize));
	coding.reader.align();
	const start = coding.reader.byte;
	const end = Math.min(coding.reader.end, start + size);
	const decoder = new ArithmeticDecoder(coding.reader.data, start, end);
	const bitmap = decodeRefinement(decoder, coding.refinement, options, budget);
	coding.reader.at = (start + size) * 8;
	return bitmap;
}
