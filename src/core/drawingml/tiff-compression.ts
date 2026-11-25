import { RasterImageError } from './raster-image';

function invalidStrip(): never {
	throw new RasterImageError('invalid-image', 'Invalid or truncated TIFF strip.');
}

/** TIFF 6.0 LZW, MSB first with early code-width changes. A validation pass needs only lengths. */
export function tiffLzw(source: Uint8Array, size: number, output?: Uint8Array): void {
	const lengths = new Uint16Array(4096);
	lengths.fill(1, 0, 256);
	const starts = output && new Uint32Array(4096);
	let bit = 0;
	let written = 0;
	let width = 9;
	let next = 258;
	let previous = -1;
	if (source[0] !== 128 || (source[1] ?? 255) >= 128) invalidStrip();
	while (bit + width <= source.length * 8) {
		const at = bit >>> 3;
		const word = (source[at]! << 16) | ((source[at + 1] ?? 0) << 8) | (source[at + 2] ?? 0);
		const code = (word >>> (24 - (bit & 7) - width)) & ((1 << width) - 1);
		bit += width;
		if (code === 256) {
			width = 9;
			next = 258;
			previous = -1;
			continue;
		}
		if (code === 257) {
			if (written !== size) invalidStrip();
			return;
		}
		if (code > next || (code === next && (previous === -1 || next === 4096))) invalidStrip();
		const special = code === next;
		const length = special ? lengths[previous]! + 1 : lengths[code]!;
		if (written + length > size) invalidStrip();
		if (output && starts) {
			const from = special ? previous : code;
			if (from < 256) output[written] = from;
			else output.copyWithin(written, starts[from]!, starts[from]! + lengths[from]!);
			if (special) output[written + length - 1] = output[written]!;
			if (previous !== -1 && next < 4096) starts[next] = written - lengths[previous]!;
		}
		if (previous !== -1 && next < 4096) {
			lengths[next++] = lengths[previous]! + 1;
			if (width < 12 && next === (1 << width) - 1) width++;
		}
		written += length;
		previous = code;
	}
	invalidStrip();
}

/** Each PackBits run must fit inside the strip's encoded and decoded ranges. */
export function tiffPackBits(source: Uint8Array, size: number, output?: Uint8Array): void {
	let written = 0;
	for (let at = 0; at < source.length;) {
		const header = source[at++]!;
		if (header === 128) continue;
		const literal = header < 128;
		const count = literal ? header + 1 : 257 - header;
		const end = at + (literal ? count : 1);
		if (end > source.length || written + count > size) invalidStrip();
		if (output) {
			if (literal) output.set(source.subarray(at, end), written);
			else output.fill(source[at]!, written, written + count);
		}
		at = end;
		written += count;
	}
	if (written !== size) invalidStrip();
}
