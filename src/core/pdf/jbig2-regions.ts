/**
 * The region segments of a JBIG2 stream (ITU-T T.88 7.4.3, 7.4.6, 7.4.7): their headers read
 * and their region bitmaps decoded (text, generic and generic refinement regions).
 */
import { ArithmeticDecoder } from './jbig2-arith';
import { type Bitmap, Jbig2Error, type Jbig2Budget } from './jbig2-bitmap';
import {
	decodeGeneric,
	decodeMmr,
	decodeRefinement,
	genericContexts,
	refinementContexts,
} from './jbig2-generic';
import {
	BitReader,
	buildTable,
	decodeHuffman,
	standardTable,
	type HuffmanTable,
	type TableLine,
} from './jbig2-huffman';
import { readRegionInfo, u32, type RegionInfo, type Segment } from './jbig2-segments';
import { readAt } from './jbig2-symbols';
import { decodeText, TextContexts, type HuffmanTextTables, type TextRegion } from './jbig2-text';

/** Decodes a text region segment (7.4.3) with the referred symbols and user tables. */
export function textRegion(
	data: Uint8Array,
	segment: Segment,
	symbols: Bitmap[],
	userTables: HuffmanTable[],
	budget: Jbig2Budget,
): { info: RegionInfo; bitmap: Bitmap } {
	const { start, end } = segment;
	const info = readRegionInfo(data, start, end);
	let at = start + 17;
	if (at + 2 > end) throw new Jbig2Error('A JBIG2 text region is too short.');
	const flags = (data[at]! << 8) | data[at + 1]!;
	at += 2;
	const huffman = (flags & 1) === 1;
	const refine = (flags & 2) === 2;
	const refinementTemplate = (flags >> 15) & 1;
	let huffmanFlags = 0;
	if (huffman) {
		huffmanFlags = (data[at]! << 8) | data[at + 1]!;
		at += 2;
	}
	const refinementAt = refine && refinementTemplate === 0 ? readAt(data, at, 2) : [];
	if (refine && refinementTemplate === 0) at += 4;
	const instances = u32(data, at);
	at += 4;
	const dsBits = (flags >> 10) & 31;
	const region: TextRegion = {
		width: info.width,
		height: info.height,
		instances,
		strips: 1 << ((flags >> 2) & 3),
		symbols,
		defaultPixel: (flags >> 9) & 1,
		operator: (flags >> 7) & 3,
		transposed: (flags & 64) !== 0,
		corner: (flags >> 4) & 3,
		dsOffset: dsBits > 15 ? dsBits - 32 : dsBits,
		refine,
		refinementTemplate,
		refinementAt,
	};
	if (!huffman) {
		let codeLength = 0;
		while (1 << codeLength < symbols.length) codeLength++;
		const coding = {
			kind: 'arith' as const,
			decoder: new ArithmeticDecoder(data, at, end),
			contexts: new TextContexts(codeLength, refinementTemplate),
		};
		return { info, bitmap: decodeText(region, coding, budget) };
	}
	let userIndex = 0;
	const pick = (value: number, standard: number[], user: number): HuffmanTable => {
		if (value === user) {
			const table = userTables[userIndex++];
			if (!table) throw new Jbig2Error('A JBIG2 text region refers to a missing table.');
			return table;
		}
		const number = standard[value];
		if (number === undefined) throw new Jbig2Error('A JBIG2 Huffman table selection is invalid.');
		return standardTable(number);
	};
	const fs = pick(huffmanFlags & 3, [6, 7], 3);
	const ds = pick((huffmanFlags >> 2) & 3, [8, 9, 10], 3);
	const dt = pick((huffmanFlags >> 4) & 3, [11, 12, 13], 3);
	const rdw = pick((huffmanFlags >> 6) & 3, [14, 15], 3);
	const rdh = pick((huffmanFlags >> 8) & 3, [14, 15], 3);
	const rdx = pick((huffmanFlags >> 10) & 3, [14, 15], 3);
	const rdy = pick((huffmanFlags >> 12) & 3, [14, 15], 3);
	const rsize = pick((huffmanFlags >> 14) & 1, [1], 1);
	const reader = new BitReader(data, at, end);
	const id = symbolIdTable(reader, symbols.length);
	const tables: HuffmanTextTables = { fs, ds, dt, rdw, rdh, rdx, rdy, rsize, id };
	const coding = {
		kind: 'huffman' as const,
		reader,
		tables,
		refinement: refinementContexts(refinementTemplate),
	};
	return { info, bitmap: decodeText(region, coding, budget) };
}

/** Reads the symbol ID Huffman table of a text region (7.4.3.1.7). */
function symbolIdTable(reader: BitReader, count: number): HuffmanTable {
	const runLines: TableLine[] = [];
	for (let index = 0; index < 35; index++)
		runLines.push({ prefix: reader.bits(4), range: 0, low: index });
	const runCodes = buildTable(runLines);
	const lengths: number[] = [];
	while (lengths.length < count) {
		const code = decodeHuffman(reader, runCodes);
		let repeat = 1;
		let length = code ?? 0;
		if (code === 32) {
			if (!lengths.length) throw new Jbig2Error('A JBIG2 symbol ID table is invalid.');
			length = lengths[lengths.length - 1]!;
			repeat = 3 + reader.bits(2);
		} else if (code === 33) {
			length = 0;
			repeat = 3 + reader.bits(3);
		} else if (code === 34) {
			length = 0;
			repeat = 11 + reader.bits(7);
		} else if (code === null || code > 34)
			throw new Jbig2Error('A JBIG2 symbol ID table is invalid.');
		if (lengths.length + repeat > count)
			throw new Jbig2Error('A JBIG2 symbol ID table is invalid.');
		for (let index = 0; index < repeat; index++) lengths.push(length);
	}
	reader.align();
	return buildTable(lengths.map((prefix, index) => ({ prefix, range: 0, low: index })));
}

/** Decodes a generic region segment (7.4.6). */
export function genericRegion(
	data: Uint8Array,
	segment: Segment,
	budget: Jbig2Budget,
): { info: RegionInfo; bitmap: Bitmap } {
	const { start } = segment;
	let end = segment.end;
	const info = readRegionInfo(data, start, end);
	const flags = data[start + 17];
	if (flags === undefined) throw new Jbig2Error('A JBIG2 generic region is too short.');
	if (flags & 16) throw new Jbig2Error('JBIG2 extended generic templates are not supported.');
	const mmr = (flags & 1) === 1;
	const template = (flags >> 1) & 3;
	let at = start + 18;
	const pixels = mmr ? [] : readAt(data, at, template === 0 ? 4 : 1);
	if (!mmr) at += template === 0 ? 8 : 2;
	if (segment.unknownLength) {
		// The data ends with the end marker and the row count (7.2.7).
		info.height = Math.min(info.height, u32(data, end - 4));
		end -= 6;
	}
	if (at > end) throw new Jbig2Error('A JBIG2 generic region is too short.');
	const bitmap = mmr
		? decodeMmr(data, at, end, info.width, info.height, budget).bitmap
		: decodeGeneric(
				new ArithmeticDecoder(data, at, end),
				genericContexts(template),
				{ width: info.width, height: info.height, template, at: pixels, tpgdon: (flags & 8) !== 0 },
				budget,
			);
	return { info, bitmap };
}

/** Decodes a generic refinement region segment (7.4.7) against `reference`. */
export function refinementRegion(
	data: Uint8Array,
	segment: Segment,
	reference: (info: RegionInfo) => Bitmap,
	budget: Jbig2Budget,
): { info: RegionInfo; bitmap: Bitmap } {
	const { start, end } = segment;
	const info = readRegionInfo(data, start, end);
	const flags = data[start + 17];
	if (flags === undefined) throw new Jbig2Error('A JBIG2 refinement region is too short.');
	const template = flags & 1;
	let at = start + 18;
	const pixels = template === 0 ? readAt(data, at, 2) : [];
	if (template === 0) at += 4;
	const bitmap = decodeRefinement(
		new ArithmeticDecoder(data, at, end),
		refinementContexts(template),
		{
			width: info.width,
			height: info.height,
			template,
			reference: reference(info),
			dx: 0,
			dy: 0,
			at: pixels,
			tpgron: (flags & 2) !== 0,
		},
		budget,
	);
	return { info, bitmap };
}
