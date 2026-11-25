export function bmpPicture(): Uint8Array<ArrayBuffer> {
	const bytes = new Uint8Array(62);
	const view = new DataView(bytes.buffer);
	view.setUint16(0, 0x4d42, true);
	view.setUint32(2, bytes.length, true);
	view.setUint32(10, 54, true);
	view.setUint32(14, 40, true);
	view.setInt32(18, 2, true);
	view.setInt32(22, 1, true);
	view.setUint16(26, 1, true);
	view.setUint16(28, 24, true);
	bytes.set([0, 0, 255, 0, 255, 0], 54);
	return bytes;
}

export function svgPicture(
	content = '<rect width="2" height="1" fill="red"/>',
): Uint8Array<ArrayBuffer> {
	return new TextEncoder().encode(
		`<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 2 1">${content}</svg>`,
	);
}

export function tiffPicture(
	compression: 1 | 5 | 32773 = 1,
	little = true,
): Uint8Array<ArrayBuffer> {
	const pixels = [255, 0, 0, 0, 255, 0];
	const source =
		compression === 5
			? packLzw([256, ...pixels, 257])
			: compression === 32773
				? [5, ...pixels]
				: pixels;
	const tags = [
		[256, 3, 1, 2],
		[257, 3, 1, 1],
		[258, 3, 3, 122],
		[259, 3, 1, compression],
		[262, 3, 1, 2],
		[273, 4, 1, 128],
		[277, 3, 1, 3],
		[278, 3, 1, 1],
		[279, 4, 1, source.length],
	];
	const bytes = new Uint8Array(128 + source.length);
	const view = new DataView(bytes.buffer);
	view.setUint16(0, little ? 0x4949 : 0x4d4d);
	view.setUint16(2, 42, little);
	view.setUint32(4, 8, little);
	view.setUint16(8, tags.length, little);
	for (const [index, [tag, type, count, value]] of tags.entries()) {
		const at = 10 + index * 12;
		view.setUint16(at, tag!, little);
		view.setUint16(at + 2, type!, little);
		view.setUint32(at + 4, count!, little);
		if (type === 3 && count === 1) view.setUint16(at + 8, value!, little);
		else view.setUint32(at + 8, value!, little);
	}
	for (const at of [122, 124, 126]) view.setUint16(at, 8, little);
	bytes.set(source, 128);
	return bytes;
}

function packLzw(codes: number[]): Uint8Array {
	const bytes = new Uint8Array(Math.ceil((codes.length * 9) / 8));
	let bit = 0;
	for (const code of codes) {
		for (let shift = 8; shift >= 0; shift--, bit++)
			bytes[bit >>> 3] = bytes[bit >>> 3]! | (((code >>> shift) & 1) << (7 - (bit & 7)));
	}
	return bytes;
}
