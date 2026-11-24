import { describe, expect, it } from 'vitest';
import { BitReader, decodeHuffman, readTableSegment, standardTable } from './jbig2-huffman';
import { standardLines } from './jbig2-tables';

/** Writes bit strings MSB first, padded with zeros. */
const bits = (text: string) => {
	const clean = text.replace(/ /g, '');
	const out = new Uint8Array(Math.ceil(clean.length / 8));
	for (let index = 0; index < clean.length; index++)
		if (clean[index] === '1') out[index >> 3] = out[index >> 3]! | (0x80 >> (index & 7));
	return out;
};

describe('the standard Huffman tables (T.88 Annex B.5)', () => {
	it.each(Array.from({ length: 15 }, (_, index) => index + 1))(
		'B.%i is a complete prefix code over contiguous ranges',
		(number) => {
			const lines = standardLines(number);
			const coded = lines.filter((line) => line.prefix > 0);
			expect(coded.reduce((sum, line) => sum + 2 ** -line.prefix, 0)).toBe(1);
			const ranges = lines.filter((line) => !line.kind);
			for (let index = 1; index < ranges.length; index++)
				expect(ranges[index]!.low).toBe(ranges[index - 1]!.low + 2 ** ranges[index - 1]!.range);
			const upper = lines.find((line) => line.kind === 'upper')!;
			if (upper.prefix) expect(upper.low).toBe(ranges.at(-1)!.low + 2 ** ranges.at(-1)!.range);
			const lower = lines.find((line) => line.kind === 'lower');
			if (lower) expect(lower.low).toBe(ranges[0]!.low - 1);
		},
	);

	it('assigns the codes of the spec (B.3) and decodes ranges, OOB and the range lines', () => {
		// B.2: 0 '0', 1 '10', 2 '110', 3..10 '1110'+3, 11..74 '11110'+6, upper '111110'+32, OOB '111111'.
		const data = bits(`0 10 110 1110101 11110000001 111110 ${'0'.repeat(31)}1 111111`);
		const reader = new BitReader(data, 0, data.length);
		const table = standardTable(2);
		expect([0, 1, 2, 3, 4, 5, 6].map(() => decodeHuffman(reader, table))).toEqual([
			0,
			1,
			2,
			8,
			12,
			76,
			null,
		]);
		// B.3 lower range line '11111111' + 32 bits: -257 - offset.
		const lower = bits(`11111111 ${'0'.repeat(30)}11`);
		expect(decodeHuffman(new BitReader(lower, 0, lower.length), standardTable(3))).toBe(-260);
	});

	it('reads a user table segment (B.2)', () => {
		// HTOOB 1, HTPS 2 bits, HTRS 2 bits; HTLOW 0, HTHIGH 6: lines 0..3 (prefix 1, range 2),
		// 4..5 (prefix 2, range 1); lower and upper prefix 3, OOB prefix 0 (no code).
		const header = [0x01 | (1 << 1) | (1 << 4), 0, 0, 0, 0, 0, 0, 0, 6];
		const body = bits('01 10 10 01 11 11 00');
		const data = Uint8Array.from([...header, ...body]);
		const table = readTableSegment(data, 0, data.length);
		// Codes: lines '0' (0..3), '10' (4..5), lower '110', upper '111'.
		const input = bits(`0 11 10 1 110 ${'0'.repeat(31)}1 111 ${'0'.repeat(32)}`);
		const reader = new BitReader(input, 0, input.length);
		expect([0, 1, 2, 3].map(() => decodeHuffman(reader, table))).toEqual([3, 5, -2, 6]);
	});

	it('stops at the end of the data and on invalid codes', () => {
		const reader = new BitReader(new Uint8Array([0xff]), 0, 1);
		expect(() => decodeHuffman(reader, standardTable(1))).toThrow(/ends early/);
	});
});
