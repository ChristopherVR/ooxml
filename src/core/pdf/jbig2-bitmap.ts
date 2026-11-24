/**
 * The bitmaps the JBIG2 decoder (ITU-T T.88) works on: one byte per pixel, 1 for black, and the
 * budget every allocation and composition is charged to, so a small hostile stream cannot make
 * the decoder allocate or loop without bound.
 */

export class Jbig2Error extends Error {}

/** Allocation and work limits for one JBIG2 stream (pixels). */
export class Jbig2Budget {
	allocated = 0;
	work = 0;
	constructor(
		readonly maxPixels: number,
		readonly maxAllocated: number,
		readonly maxWork: number,
	) {}
	allocate(width: number, height: number): void {
		const pixels = width * height;
		if (
			!Number.isInteger(width) ||
			!Number.isInteger(height) ||
			width < 0 ||
			height < 0 ||
			pixels > this.maxPixels
		)
			throw new Jbig2Error('A JBIG2 bitmap is larger than the size limit.');
		this.allocated += pixels;
		if (this.allocated > this.maxAllocated)
			throw new Jbig2Error('The JBIG2 bitmaps are larger than the size limit.');
	}
	spend(pixels: number): void {
		this.work += pixels + 1;
		if (this.work > this.maxWork) throw new Jbig2Error('The JBIG2 data needs too much work.');
	}
}

export class Bitmap {
	readonly data: Uint8Array;
	constructor(
		readonly width: number,
		readonly height: number,
		budget: Jbig2Budget,
		fill = 0,
	) {
		budget.allocate(width, height);
		this.data = new Uint8Array(width * height);
		if (fill) this.data.fill(1);
	}
	get(x: number, y: number): number {
		return x < 0 || y < 0 || x >= this.width || y >= this.height
			? 0
			: this.data[y * this.width + x]!;
	}
	set(x: number, y: number, value: number): void {
		this.data[y * this.width + x] = value;
	}
}

/** Combination operators (7.4.1.5, 7.4.3.1.1): OR, AND, XOR, XNOR, REPLACE. */
export const OR = 0;
export const REPLACE = 4;

/** Draws `source` into `target` with its top-left corner at (x, y), clipped to the target. */
export function compose(
	target: Bitmap,
	source: Bitmap,
	x: number,
	y: number,
	operator: number,
	budget: Jbig2Budget,
): void {
	const x0 = Math.max(0, x);
	const y0 = Math.max(0, y);
	const x1 = Math.min(target.width, x + source.width);
	const y1 = Math.min(target.height, y + source.height);
	if (x0 >= x1 || y0 >= y1) return budget.spend(0);
	budget.spend((x1 - x0) * (y1 - y0));
	const t = target.data;
	const s = source.data;
	for (let row = y0; row < y1; row++) {
		let at = row * target.width + x0;
		let from = (row - y) * source.width + (x0 - x);
		for (let col = x0; col < x1; col++, at++, from++) {
			const a = t[at]!;
			const b = s[from]!;
			t[at] =
				operator === 0
					? a | b
					: operator === 1
						? a & b
						: operator === 2
							? a ^ b
							: operator === 3
								? 1 ^ a ^ b
								: b;
		}
	}
}

/** A copy of the `width` x `height` area of `source` at (x, y); pixels outside are 0. */
export function crop(
	source: Bitmap,
	x: number,
	y: number,
	width: number,
	height: number,
	budget: Jbig2Budget,
): Bitmap {
	const out = new Bitmap(width, height, budget);
	compose(out, source, -x, -y, REPLACE, budget);
	return out;
}

/** Packs rows MSB first, 1 for white (as PDF's JBIG2Decode delivers them), at width x height. */
export function packInverted(bitmap: Bitmap, width: number, height: number): Uint8Array {
	const stride = (width + 7) >> 3;
	const out = new Uint8Array(stride * height).fill(0xff);
	const rows = Math.min(height, bitmap.height);
	const cols = Math.min(width, bitmap.width);
	for (let y = 0; y < rows; y++)
		for (let x = 0; x < cols; x++)
			if (bitmap.data[y * bitmap.width + x]) {
				const at = y * stride + (x >> 3);
				out[at] = out[at]! & ~(0x80 >> (x & 7));
			}
	return out;
}
