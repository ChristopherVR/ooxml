import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { decodeJpx } from './index';

const fixtures = JSON.parse(
	readFileSync(new URL('./__fixtures__/jpx.json', import.meta.url), 'utf8'),
) as { cases: { name: string; data: string }[] };
const file = (name: string) =>
	new Uint8Array(Buffer.from(fixtures.cases.find((item) => item.name === name)!.data, 'base64'));

const u16 = (data: Uint8Array, at: number) => (data[at]! << 8) | data[at + 1]!;
const be16 = (value: number) => [value >> 8, value & 0xff];
const be32 = (value: number) => [
	value >>> 24,
	(value >> 16) & 0xff,
	(value >> 8) & 0xff,
	value & 0xff,
];

/**
 * Moves the packet headers of a codestream written with SOP and EPH markers into PPM (main
 * header) or PPT (tile-part header) marker segments, as an encoder that packs headers would.
 */
function pack(codestream: Uint8Array, where: 'ppm' | 'ppt'): Uint8Array {
	const main: number[] = [0xff, 0x4f];
	let at = 2;
	while (u16(codestream, at) !== 0xff90) {
		const end = at + 2 + u16(codestream, at + 2);
		// TLM lengths would be stale once the tile-parts change.
		if (u16(codestream, at) !== 0xff55) main.push(...codestream.subarray(at, end));
		at = end;
	}
	const parts: { sot: Uint8Array; markers: number[]; headers: number[]; data: number[] }[] = [];
	while (u16(codestream, at) === 0xff90) {
		const length = (u16(codestream, at + 6) << 16) + u16(codestream, at + 8);
		const end = at + length;
		const sot = codestream.slice(at, at + 12);
		const markers: number[] = [];
		let p = at + 12;
		while (u16(codestream, p) !== 0xff93) {
			const next = p + 2 + u16(codestream, p + 2);
			if (u16(codestream, p) !== 0xff58) markers.push(...codestream.subarray(p, next));
			p = next;
		}
		const body = codestream.subarray(p + 2, end);
		const headers: number[] = [];
		const data: number[] = [];
		for (let q = 0; q < body.length;) {
			expect(u16(body, q)).toBe(0xff91);
			let eph = q + 6;
			while (u16(body, eph) !== 0xff92) eph++;
			headers.push(...body.subarray(q + 6, eph + 2));
			let next = eph + 2;
			while (next < body.length && u16(body, next) !== 0xff91) next++;
			data.push(...body.subarray(q, q + 6), ...body.subarray(eph + 2, next));
			q = next;
		}
		parts.push({ sot, markers, headers, data });
		at = end;
	}
	if (where === 'ppm') {
		const records = parts.flatMap((part) => [...be32(part.headers.length), ...part.headers]);
		main.push(0xff, 0x60, ...be16(3 + records.length), 0, ...records);
	}
	const out = [...main];
	const zppt = new Map<number, number>();
	for (const part of parts) {
		const tile = u16(part.sot, 4);
		const ppt =
			where === 'ppt'
				? [0xff, 0x61, ...be16(3 + part.headers.length), zppt.get(tile) ?? 0, ...part.headers]
				: [];
		zppt.set(tile, (zppt.get(tile) ?? 0) + 1);
		const length = 12 + part.markers.length + ppt.length + 2 + part.data.length;
		out.push(...part.sot.subarray(0, 6), ...be32(length), ...part.sot.subarray(10, 12));
		out.push(...part.markers, ...ppt, 0xff, 0x93, ...part.data);
	}
	out.push(0xff, 0xd9);
	return Uint8Array.from(out);
}

describe('packed packet headers', () => {
	it('reads headers from PPM and PPT marker segments', () => {
		const original = decodeJpx(file('rgb-markers'));
		for (const where of ['ppm', 'ppt'] as const) {
			const data = pack(file('rgb-markers'), where);
			const code = where === 'ppm' ? 0x60 : 0x61;
			expect(data.some((byte, i) => byte === 0xff && data[i + 1] === code)).toBe(true);
			const packed = decodeJpx(data);
			packed.components.forEach((component, c) =>
				expect([...component.data]).toEqual([...original.components[c]!.data]),
			);
		}
	});
});
