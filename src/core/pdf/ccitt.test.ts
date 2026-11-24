import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { decodeCcitt } from './ccitt';

interface Case {
	name: string;
	width: number;
	height: number;
	k: number;
	encodedByteAlign: boolean;
	data: string;
	expected: string;
}

const cases = JSON.parse(
	readFileSync(new URL('./__fixtures__/ccitt.json', import.meta.url), 'utf8'),
) as Case[];
const bytes = (base64: string) => new Uint8Array(Buffer.from(base64, 'base64'));

describe('decodeCcitt (data encoded by libtiff)', () => {
	it.each(cases.map((item) => [item.name, item] as const))('decodes %s', (_name, item) => {
		const out = decodeCcitt(
			bytes(item.data),
			{
				k: item.k,
				columns: item.width,
				rows: item.height,
				encodedByteAlign: item.encodedByteAlign,
			},
			1e7,
		);
		expect(Buffer.from(out).toString('base64')).toBe(item.expected);
	});

	it('inverts with BlackIs1 and stops at damage, the size limit and bad widths', () => {
		const group4 = cases.find((item) => item.name === 'sample group4')!;
		const options = { k: -1, columns: group4.width, rows: group4.height };
		const normal = decodeCcitt(bytes(group4.data), options, 1e7);
		const inverted = decodeCcitt(bytes(group4.data), { ...options, blackIs1: true }, 1e7);
		expect(inverted.every((byte, index) => byte === (~normal[index]! & 0xff))).toBe(true);
		const damaged = bytes(group4.data).slice(0, 20);
		expect(decodeCcitt(damaged, options, 1e7)).toHaveLength(normal.length);
		expect(decodeCcitt(new Uint8Array(64), options, 1e7).every((byte) => byte === 0xff)).toBe(true);
		expect(decodeCcitt(bytes(group4.data), options, 100).length).toBeLessThanOrEqual(100);
		expect(() => decodeCcitt(new Uint8Array(4), { ...options, columns: 0 }, 1e7)).toThrow();
	});
});
