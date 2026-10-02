// Entirely generated format cases. No corpus or converter source bytes are redistributed.
export function record(type: number, words: number[] = []): Uint8Array {
	const b = new Uint8Array(8 + words.length * 4),
		v = new DataView(b.buffer);
	v.setUint32(0, type, true);
	v.setUint32(4, b.length, true);
	words.forEach((n, i) => v.setUint32(8 + i * 4, n, true));
	return b;
}
export function emf(records: Uint8Array[], headerSize = 108): Uint8Array {
	const eof = record(14, [0, 16, 20]);
	const bytes = new Uint8Array(headerSize + records.reduce((n, r) => n + r.length, 0) + eof.length);
	const v = new DataView(bytes.buffer);
	for (const [at, n] of [
		[0, 1],
		[4, headerSize],
		[16, 99],
		[20, 99],
		[32, 1000],
		[36, 1000],
		[40, 0x464d4520],
		[44, 0x10000],
		[48, bytes.length],
		[52, records.length + 2],
		[56, 32],
		[72, 1000],
		[76, 1000],
		[80, 100],
		[84, 100],
	])
		v.setUint32(at!, n!, true);
	if (headerSize === 108) {
		v.setUint32(100, 100000, true);
		v.setUint32(104, 100000, true);
	}
	let offset = headerSize;
	for (const r of [...records, eof]) {
		bytes.set(r, offset);
		offset += r.length;
	}
	return bytes;
}
export function polygon16(points: [number, number][], type = 86): Uint8Array {
	const b = record(type, [
		0,
		0,
		99,
		99,
		points.length,
		...points.map(([x, y]) => (x & 65535) | ((y & 65535) << 16)),
	]);
	return b;
}
export function setWord(bytes: Uint8Array, at: number, word: number): Uint8Array {
	const result = bytes.slice();
	new DataView(result.buffer).setUint32(at, word, true);
	return result;
}
export function sourceComment(): Uint8Array {
	// The ignored payload is deliberately not a decoded WMF stream.
	return record(70, [44, 0x43494447, 0x80000001, 0x300, 0, 0, 20, 0x2b464d45, 0, 0, 0, 0]);
}
export function checksumComment(bytes: Uint8Array, commentOffset = 108): Uint8Array {
	const result = setWord(bytes, commentOffset + 24, 0),
		v = new DataView(result.buffer);
	let sum = 0;
	for (let i = 0; i < result.length; i += 4) sum = (sum + v.getUint32(i, true)) >>> 0;
	v.setUint32(commentOffset + 24, -sum >>> 0, true);
	return result;
}
export const triangle: [number, number][] = [
	[0, 0],
	[10, 0],
	[0, 10],
];
