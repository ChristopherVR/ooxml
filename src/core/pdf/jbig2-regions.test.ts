import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { decodeJbig2Pages } from './jbig2';
import {
	atBytes,
	bitmapRows as rows,
	joinSegments as stream,
	pageInfo,
	regionInfo,
	SAMPLE_ROWS as ROWS,
	segment,
} from './test-support/jbig2-encode';
import { encodeGeneric, encodeRefinement } from './test-support/jbig2-encode-regions';

const W = ROWS[0]!.length;
const H = ROWS.length;

/** An immediate lossless generic region segment (type 39) with arithmetic data. */
function generic(template: number, at: [number, number][], tpgdon: boolean, number = 1): number[] {
	const flags = (template << 1) | (tpgdon ? 8 : 0);
	return segment(number, 39, [
		...regionInfo(W, H),
		flags,
		...atBytes(at),
		...encodeGeneric(ROWS, template, at, tpgdon),
	]);
}

describe('generic regions', () => {
	const cases: [number, [number, number][]][] = [
		[
			0,
			[
				[3, -1],
				[-3, -1],
				[2, -2],
				[-2, -2],
			],
		],
		[
			0,
			[
				[-2, 0],
				[1, -2],
				[-4, -1],
				[4, -2],
			],
		],
		[1, [[3, -1]]],
		[1, [[-3, -2]]],
		[2, [[2, -1]]],
		[3, [[2, -1]]],
		[3, [[-1, -1]]],
	];
	it.each(cases.flatMap(([t, at]) => [[t, at, false] as const, [t, at, true] as const]))(
		'decode template %i with AT %j (TPGDON %s)',
		(template, at, tpgdon) => {
			const [page] = decodeJbig2Pages(
				stream(segment(0, 48, pageInfo(W, H)), generic(template, at, tpgdon)),
			);
			expect(rows(page!)).toEqual(ROWS);
		},
	);

	it('decode MMR regions with the CCITT Group 4 decoder', () => {
		const item = (
			JSON.parse(readFileSync(new URL('./__fixtures__/ccitt.json', import.meta.url), 'utf8')) as {
				name: string;
				width: number;
				height: number;
				data: string;
				expected: string;
			}[]
		).find((entry) => entry.name === 'sample group4')!;
		const data = Buffer.from(item.data, 'base64');
		const [page] = decodeJbig2Pages(
			stream(
				segment(0, 48, pageInfo(item.width, item.height)),
				segment(1, 38, [...regionInfo(item.width, item.height), 1, ...data]),
			),
		);
		// The CCITT fixture holds 1 for white; JBIG2 holds 1 for black.
		const expected = Buffer.from(item.expected, 'base64');
		const stride = (item.width + 7) >> 3;
		for (let y = 0; y < item.height; y++)
			for (let x = 0; x < item.width; x++)
				expect(page!.data[y * item.width + x]).toBe(
					((expected[y * stride + (x >> 3)]! >> (7 - (x & 7))) & 1) ^ 1,
				);
	});

	it('place regions with combination operators on the page default pixel', () => {
		const data = encodeGeneric(
			ROWS,
			0,
			[
				[3, -1],
				[-3, -1],
				[2, -2],
				[-2, -2],
			],
			false,
		);
		const region = (number: number, x: number, y: number, op: number) =>
			segment(number, 38, [
				...regionInfo(W, H, x, y, op),
				0,
				...atBytes([
					[3, -1],
					[-3, -1],
					[2, -2],
					[-2, -2],
				]),
				...data,
			]);
		// Default pixel 1 (flags bit 2): XOR at (2, 1) inverts, REPLACE at (30, 0) copies.
		const [page] = decodeJbig2Pages(
			stream(segment(0, 48, pageInfo(60, 20, 4)), region(1, 2, 1, 2), region(2, 30, 0, 4)),
		);
		const out = rows(page!);
		expect(out[0]!.slice(0, 30)).toBe('1'.repeat(30));
		expect(out[2]!.slice(2, 2 + W)).toBe(ROWS[1]!.replace(/./g, (c) => (c === '1' ? '0' : '1')));
		expect(out.slice(0, H).map((row) => row.slice(30, 30 + W))).toEqual(ROWS);
		expect(out[19]).toBe('1'.repeat(60));
	});

	it('grow a striped page of unknown height to its end of stripe, with an unknown-length region', () => {
		const at: [number, number][] = [
			[3, -1],
			[-3, -1],
			[2, -2],
			[-2, -2],
		];
		const coded = encodeGeneric(ROWS, 0, at, false);
		const data = [...regionInfo(W, 0xffffffff, 0, 0), 0, ...atBytes(at), ...coded, ...[0, 0, 0, H]];
		const [page] = decodeJbig2Pages(
			stream(
				segment(0, 48, pageInfo(W, 0xffffffff, 0, 0x8000 | 64)),
				segment(1, 38, data, { length: 0xffffffff }),
				segment(2, 50, [0, 0, 0, H + 4]),
			),
		);
		expect(page!.height).toBe(H + 5);
		expect(rows(page!).slice(0, H)).toEqual(ROWS);
	});
});

describe('refinement regions', () => {
	const target = ROWS.map((row, y) => (y === 4 ? row.replace('0000', '0110') : row));
	it.each([
		[0, false],
		[0, true],
		[1, false],
		[1, true],
	] as const)('refine an intermediate region (template %i, TPGRON %s)', (template, tpgron) => {
		const at: [number, number][] = [
			[-1, -1],
			[-1, -1],
		];
		const flags = template | (tpgron ? 2 : 0);
		const coded = encodeRefinement(target, ROWS, template, at, tpgron);
		const [page] = decodeJbig2Pages(
			stream(
				segment(0, 48, pageInfo(W + 4, H + 2)),
				segment(1, 36, [
					...regionInfo(W, H),
					0,
					...atBytes([
						[3, -1],
						[-3, -1],
						[2, -2],
						[-2, -2],
					]),
					...encodeGeneric(
						ROWS,
						0,
						[
							[3, -1],
							[-3, -1],
							[2, -2],
							[-2, -2],
						],
						false,
					),
				]),
				segment(
					2,
					42,
					[...regionInfo(W, H, 2, 1, 4), flags, ...(template ? [] : atBytes(at)), ...coded],
					{ referred: [1] },
				),
			),
		);
		expect(
			rows(page!)
				.slice(1, H + 1)
				.map((row) => row.slice(2, 2 + W)),
		).toEqual(target);
	});

	it('refine the page itself when no region is referred to', () => {
		const coded = encodeRefinement(target, ROWS, 1, [], true);
		const at: [number, number][] = [
			[3, -1],
			[-3, -1],
			[2, -2],
			[-2, -2],
		];
		const [page] = decodeJbig2Pages(
			stream(
				segment(0, 48, pageInfo(W, H)),
				generic(0, at, false),
				segment(2, 42, [...regionInfo(W, H, 0, 0, 4), 3, ...coded]),
			),
		);
		expect(rows(page!)).toEqual(target);
	});
});
