import { deflateSync } from 'node:zlib';
/** Nonuniform RGBA grid for native/shared picture acceptance: RGB bands, half alpha below. */
export function pictureGridPng(): Buffer {
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
	header.writeUInt32BE(48, 0);
	header.writeUInt32BE(32, 4);
	header[8] = 8;
	header[9] = 6;
	const raw = Buffer.alloc(32 * (48 * 4 + 1));
	for (let y = 0; y < 32; y++)
		for (let x = 0; x < 48; x++) {
			const i = y * 193 + 1 + x * 4;
			const col = x < 16 ? [255, 0, 0] : x < 32 ? [0, 255, 0] : [0, 0, 255];
			raw[i] = col[0]!;
			raw[i + 1] = col[1]!;
			raw[i + 2] = col[2]!;
			raw[i + 3] = y < 16 ? 255 : 128;
		}
	const png = Buffer.concat([
		Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]),
		chunk('IHDR', header),
		chunk('IDAT', deflateSync(raw)),
		chunk('IEND', Buffer.alloc(0)),
	]);

	return png;
}
