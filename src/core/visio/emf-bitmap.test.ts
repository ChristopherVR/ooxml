import { describe, expect, it } from 'vitest';
import { EmfBitmapError, encodeBitmapEmf } from './emf-bitmap';

const solid = (width: number, height: number, rgba: [number, number, number, number]) =>
	Uint8Array.from({ length: width * height * 4 }, (_, i) => rgba[i % 4]!);

describe('encodeBitmapEmf', () => {
	it('writes a header, one stretched DIB and an end record in the sizes it declares', () => {
		const emf = encodeBitmapEmf(3, 2, solid(3, 2, [10, 20, 30, 255]));
		const view = new DataView(emf.buffer);
		expect(view.getUint32(0, true)).toBe(1); // EMR_HEADER
		expect(view.getUint32(40, true)).toBe(0x464d4520); // " EMF"
		expect(view.getUint32(48, true)).toBe(emf.byteLength);
		expect(view.getUint32(52, true)).toBe(3);
		// Bounds are inclusive device pixels.
		expect([view.getInt32(16, true), view.getInt32(20, true)]).toEqual([2, 1]);
		const bitmap = 88;
		expect(view.getUint32(bitmap, true)).toBe(81); // EMR_STRETCHDIBITS
		const bitmapSize = view.getUint32(bitmap + 4, true);
		// 3 pixels * 3 bytes = 9, padded to 12 per row, 2 rows.
		expect(bitmapSize).toBe(80 + 40 + 24);
		const end = bitmap + bitmapSize;
		expect(view.getUint32(end, true)).toBe(14); // EMR_EOF
		expect(end + 20).toBe(emf.byteLength);
	});

	it('stores pixels bottom-up as BGR, padded to four bytes', () => {
		// Top row red, bottom row blue (2 x 2).
		const rgba = Uint8Array.of(255, 0, 0, 255, 255, 0, 0, 255, 0, 0, 255, 255, 0, 0, 255, 255);
		const emf = encodeBitmapEmf(2, 2, rgba);
		const bits = 88 + 80 + 40;
		// First stored row is the bottom row: blue = B 255, G 0, R 0.
		expect([...emf.slice(bits, bits + 3)]).toEqual([255, 0, 0]);
		// Second stored row is the top row: red = B 0, G 0, R 255.
		expect([...emf.slice(bits + 8, bits + 11)]).toEqual([0, 0, 255]);
	});

	it('flattens transparency onto white', () => {
		const emf = encodeBitmapEmf(1, 1, Uint8Array.of(0, 0, 0, 0));
		const bits = 88 + 80 + 40;
		expect([...emf.slice(bits, bits + 3)]).toEqual([255, 255, 255]);
	});

	it('refuses sizes and data it cannot describe', () => {
		expect(() => encodeBitmapEmf(0, 1, new Uint8Array())).toThrow(EmfBitmapError);
		expect(() => encodeBitmapEmf(1.5, 1, new Uint8Array(6))).toThrow(EmfBitmapError);
		expect(() => encodeBitmapEmf(2, 2, new Uint8Array(4))).toThrow(/does not match/);
		expect(() => encodeBitmapEmf(9000, 1, new Uint8Array(36000))).toThrow(/larger than/);
	});
});
