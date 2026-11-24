/**
 * JBIG2 symbol dictionary segments (ITU-T T.88 7.4.2 and 6.5): height classes of symbols coded
 * one by one with the generic region decoder, refined or aggregated from other symbols, or as
 * Huffman-coded collective bitmaps (uncompressed or MMR), and the export flags.
 */
import { ArithmeticDecoder, decodeId, decodeInteger, integerContexts } from './jbig2-arith';
import { Bitmap, crop, Jbig2Error, type Jbig2Budget } from './jbig2-bitmap';
import {
	decodeGeneric,
	decodeMmr,
	decodeRefinement,
	genericContexts,
	refinementContexts,
	type At,
} from './jbig2-generic';
import { BitReader, decodeHuffman, standardTable, type HuffmanTable } from './jbig2-huffman';
import { decodeText, TextContexts, type HuffmanTextTables } from './jbig2-text';

export interface SymbolDictionary {
	exported: Bitmap[];
	/** The arithmetic coding contexts, when the segment says they are retained. */
	retained?: { generic: Uint8Array; refinement: Uint8Array };
}

const MAX_SYMBOLS = 1 << 20;

/** Reads `count` AT pixel pairs (signed bytes) at `at`. */
export function readAt(data: Uint8Array, at: number, count: number): At[] {
	const out: At[] = [];
	for (let index = 0; index < count; index++)
		out.push([(data[at + 2 * index]! << 24) >> 24, (data[at + 2 * index + 1]! << 24) >> 24]);
	return out;
}

const symbolCodeLength = (count: number): number => {
	let length = 0;
	while (1 << length < count) length++;
	return length;
};

/** Decodes a symbol dictionary segment's data `data[start, end)`. */
export function decodeSymbolDictionary(
	data: Uint8Array,
	start: number,
	end: number,
	inputs: Bitmap[],
	userTables: HuffmanTable[],
	previous: SymbolDictionary | undefined,
	budget: Jbig2Budget,
): SymbolDictionary {
	if (end - start < 10) throw new Jbig2Error('A JBIG2 symbol dictionary is too short.');
	const flags = (data[start]! << 8) | data[start + 1]!;
	const huffman = (flags & 1) === 1;
	const refAgg = (flags & 2) === 2;
	const template = (flags >> 10) & 3;
	const refTemplate = (flags >> 12) & 1;
	let at = start + 2;
	const gbAt = huffman ? [] : readAt(data, at, template === 0 ? 4 : 1);
	if (!huffman) at += template === 0 ? 8 : 2;
	const grAt = refAgg && refTemplate === 0 ? readAt(data, at, 2) : [];
	if (refAgg && refTemplate === 0) at += 4;
	if (at + 8 > end) throw new Jbig2Error('A JBIG2 symbol dictionary is too short.');
	const view = new DataView(data.buffer, data.byteOffset + at, 8);
	const exportedCount = view.getUint32(0);
	const newCount = view.getUint32(4);
	at += 8;
	if (newCount + inputs.length > MAX_SYMBOLS || exportedCount > MAX_SYMBOLS)
		throw new Jbig2Error('A JBIG2 symbol dictionary has too many symbols.');

	let userIndex = 0;
	/** A standard table by selector value, or the next user table when the value is `user`. */
	const pick = (value: number, standard: number[], user: number): HuffmanTable => {
		if (value !== user) {
			const number = standard[value];
			if (number === undefined) throw new Jbig2Error('A JBIG2 Huffman table selection is invalid.');
			return standardTable(number);
		}
		const table = userTables[userIndex++];
		if (!table) throw new Jbig2Error('A JBIG2 symbol dictionary refers to a missing table.');
		return table;
	};
	const dhTable = huffman ? pick((flags >> 2) & 3, [4, 5], 3) : undefined;
	const dwTable = huffman ? pick((flags >> 4) & 3, [2, 3], 3) : undefined;
	const bmTable = huffman ? pick((flags >> 6) & 1, [1], 1) : undefined;
	const aggTable = huffman && refAgg ? pick((flags >> 7) & 1, [1], 1) : undefined;

	const reader = huffman ? new BitReader(data, at, end) : undefined;
	const decoder = huffman ? undefined : new ArithmeticDecoder(data, at, end);
	// The contexts may continue from the last dictionary referred to (7.4.2.1.1, bit 8).
	const kept = (flags & 0x100) !== 0 ? previous?.retained : undefined;
	const generic = genericContexts(template);
	const refinement = refinementContexts(refTemplate);
	if (kept?.generic.length === generic.length) generic.set(kept.generic);
	if (kept?.refinement.length === refinement.length) refinement.set(kept.refinement);
	const iadh = integerContexts();
	const iadw = integerContexts();
	const iaex = integerContexts();
	const iaai = integerContexts();
	const codeLength = symbolCodeLength(inputs.length + newCount);
	const text = refAgg && !huffman ? new TextContexts(codeLength, refTemplate) : undefined;
	text?.refinement.set(refinement);
	const integer = (contexts: Uint8Array, table: HuffmanTable | undefined) =>
		decoder ? decodeInteger(decoder, contexts) : decodeHuffman(reader!, table!);
	const need = (value: number | null): number => {
		if (value === null) throw new Jbig2Error('A JBIG2 symbol dictionary has an unexpected OOB.');
		return value;
	};

	const symbols: Bitmap[] = [];
	let height = 0;
	while (symbols.length < newCount) {
		height += need(integer(iadh, dhTable));
		if (height < 0) throw new Jbig2Error('A JBIG2 symbol height is negative.');
		budget.spend(0);
		let width = 0;
		let totalWidth = 0;
		const widths: number[] = [];
		for (;;) {
			const dw = integer(iadw, dwTable);
			if (dw === null) break;
			if (symbols.length + widths.length >= newCount)
				throw new Jbig2Error('A JBIG2 symbol dictionary has more symbols than it says.');
			width += dw;
			if (width < 0) throw new Jbig2Error('A JBIG2 symbol width is negative.');
			totalWidth += width;
			if (huffman && !refAgg) {
				widths.push(width);
				budget.spend(0);
				continue;
			}
			if (!refAgg) {
				symbols.push(
					decodeGeneric(
						decoder!,
						generic,
						{ width, height, template, at: gbAt, tpgdon: false },
						budget,
					),
				);
				continue;
			}
			const instances = need(integer(iaai, aggTable));
			const all = [...inputs, ...symbols];
			if (instances > 1) {
				const tables: HuffmanTextTables | undefined = reader
					? {
							fs: standardTable(6),
							ds: standardTable(8),
							dt: standardTable(11),
							rdw: standardTable(15),
							rdh: standardTable(15),
							rdx: standardTable(15),
							rdy: standardTable(15),
							rsize: standardTable(1),
							id: codeLength,
						}
					: undefined;
				const region = {
					width,
					height,
					instances,
					strips: 1,
					symbols: all,
					defaultPixel: 0,
					operator: 0,
					transposed: false,
					corner: 1,
					dsOffset: 0,
					refine: true,
					refinementTemplate: refTemplate,
					refinementAt: grAt,
				};
				symbols.push(
					decodeText(
						region,
						decoder
							? { kind: 'arith', decoder, contexts: text! }
							: { kind: 'huffman', reader: reader!, tables: tables!, refinement },
						budget,
					),
				);
				continue;
			}
			symbols.push(refineOne(width, height, all));
		}
		if (huffman && !refAgg && widths.length) collective(widths, totalWidth, height);
	}

	function refineOne(width: number, height: number, all: Bitmap[]): Bitmap {
		const id = decoder ? decodeId(decoder, text!.id, codeLength) : reader!.bits(codeLength);
		const rdx = need(
			decoder ? decodeInteger(decoder, text!.rdx) : decodeHuffman(reader!, standardTable(15)),
		);
		const rdy = need(
			decoder ? decodeInteger(decoder, text!.rdy) : decodeHuffman(reader!, standardTable(15)),
		);
		const reference = all[id];
		if (!reference) throw new Jbig2Error('A JBIG2 symbol refers to a missing symbol.');
		const options = {
			width,
			height,
			template: refTemplate,
			reference,
			dx: rdx,
			dy: rdy,
			at: grAt,
			tpgron: false,
		};
		if (decoder) return decodeRefinement(decoder, text!.refinement, options, budget);
		const size = need(decodeHuffman(reader!, standardTable(1)));
		reader!.align();
		const from = reader!.byte;
		const bitmap = decodeRefinement(
			new ArithmeticDecoder(data, from, Math.min(end, from + size)),
			refinement,
			options,
			budget,
		);
		reader!.at = (from + size) * 8;
		return bitmap;
	}

	function collective(widths: number[], totalWidth: number, height: number): void {
		const size = need(decodeHuffman(reader!, bmTable!));
		reader!.align();
		const from = reader!.byte;
		let bitmap: Bitmap;
		if (size === 0) {
			bitmap = new Bitmap(totalWidth, height, budget);
			const stride = (totalWidth + 7) >> 3;
			if (from + stride * height > end)
				throw new Jbig2Error('A JBIG2 collective bitmap ends early.');
			for (let y = 0; y < height; y++)
				for (let x = 0; x < totalWidth; x++)
					bitmap.data[y * totalWidth + x] =
						(data[from + y * stride + (x >> 3)]! >> (7 - (x & 7))) & 1;
			reader!.at = (from + stride * height) * 8;
		} else {
			bitmap = decodeMmr(data, from, Math.min(end, from + size), totalWidth, height, budget).bitmap;
			reader!.at = (from + size) * 8;
		}
		let x = 0;
		for (const width of widths) {
			symbols.push(crop(bitmap, x, 0, width, height, budget));
			x += width;
		}
	}

	// The export flags (6.5.10): alternating runs of symbols not exported and exported.
	const all = [...inputs, ...symbols];
	const exported: Bitmap[] = [];
	let index = 0;
	let flag = 0;
	for (let runs = 0; index < all.length; runs++) {
		if (runs > 2 * all.length + 2) throw new Jbig2Error('The JBIG2 export flags do not end.');
		const run = need(integer(iaex, huffman ? standardTable(1) : undefined));
		if (run < 0 || index + run > all.length)
			throw new Jbig2Error('The JBIG2 export flags are out of range.');
		if (flag) exported.push(...all.slice(index, index + run));
		index += run;
		flag ^= 1;
	}
	const retained =
		(flags & 0x200) !== 0 ? { generic, refinement: text?.refinement ?? refinement } : undefined;
	return retained ? { exported, retained } : { exported };
}
