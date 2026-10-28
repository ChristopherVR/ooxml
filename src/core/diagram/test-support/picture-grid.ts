import { deflateSync } from 'node:zlib';
/** Nonuniform RGBA grid for native/shared picture acceptance: RGB bands, half alpha below. */
export function rgbaPng(width: number, height: number, data: Uint8ClampedArray): Buffer {
	function crc(bytes: Uint8Array) {
		let n = -1;
		for (const byte of bytes) {
			n ^= byte;
			for (let bit = 0; bit < 8; bit++) n = (n >>> 1) ^ (n & 1 ? 0xedb88320 : 0);
		}
		return (n ^ -1) >>> 0;
	}
	function chunk(name: string, bytes: Uint8Array) {
		const content = Buffer.concat([Buffer.from(name), bytes]);
		const size = Buffer.alloc(4);
		size.writeUInt32BE(bytes.length);
		const hash = Buffer.alloc(4);
		hash.writeUInt32BE(crc(content));
		return Buffer.concat([size, content, hash]);
	}
	const header = Buffer.alloc(13);
	header.writeUInt32BE(width, 0);
	header.writeUInt32BE(height, 4);
	header[8] = 8;
	header[9] = 6;
	const raw = Buffer.alloc(height * (width * 4 + 1));
	for (let y = 0; y < height; y++)
		raw.set(data.subarray(y * width * 4, (y + 1) * width * 4), y * (width * 4 + 1) + 1);
	return Buffer.concat([
		Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]),
		chunk('IHDR', header),
		chunk('IDAT', deflateSync(raw)),
		chunk('IEND', Buffer.alloc(0)),
	]);
}
export function pictureGridPixels(): Uint8ClampedArray {
	const data = new Uint8ClampedArray(48 * 32 * 4);
	for (let y = 0; y < 32; y++)
		for (let x = 0; x < 48; x++) {
			const i = (y * 48 + x) * 4;
			const colour = x < 16 ? [255, 0, 0] : x < 32 ? [0, 255, 0] : [0, 0, 255];
			data.set([...colour, y < 16 ? 255 : 128], i);
		}
	return data;
}
export function pictureGridPng(): Buffer {
	return rgbaPng(48, 32, pictureGridPixels());
}
