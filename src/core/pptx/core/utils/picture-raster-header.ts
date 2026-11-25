/** Limits for the raster authoring commands, before handing bytes to a browser decoder. */
export const MAX_PICTURE_EDIT_BYTES = 24 * 1024 * 1024;
export const MAX_PICTURE_EDIT_PIXELS = 16_000_000;
export const MAX_PICTURE_EDIT_EDGE = 8192;

export interface PictureRasterHeader {
	type: 'image/png' | 'image/jpeg';
	width: number;
	height: number;
}

export function canEditRasterDimensions(width: number, height: number): boolean {
	return (
		Number.isInteger(width) &&
		Number.isInteger(height) &&
		width > 0 &&
		height > 0 &&
		width <= MAX_PICTURE_EDIT_EDGE &&
		height <= MAX_PICTURE_EDIT_EDGE &&
		width * height <= MAX_PICTURE_EDIT_PIXELS
	);
}

function pngHeader(bytes: DataView): PictureRasterHeader | undefined {
	const width = bytes.getUint32(16);
	const height = bytes.getUint32(20);
	if (
		bytes.getUint32(8) !== 13 ||
		bytes.getUint32(12) !== 0x49484452 ||
		!canEditRasterDimensions(width, height)
	)
		return undefined;
	// Exclude animated PNGs and malformed chunk boundaries before native decoding.
	let cursor = 33;
	for (let count = 0; count < 50_000 && cursor + 12 <= bytes.byteLength; count++) {
		const length = bytes.getUint32(cursor);
		const type = bytes.getUint32(cursor + 4);
		cursor += length + 12;
		if (cursor > bytes.byteLength || type === 0x6163544c || type === 0x49484452) return undefined;
		if (type === 0x49454e44) {
			return length === 0 && cursor === bytes.byteLength
				? { type: 'image/png', width, height }
				: undefined;
		}
	}
	return undefined;
}

function jpegHeader(bytes: DataView): PictureRasterHeader | undefined {
	let cursor = 2;
	let header: PictureRasterHeader | undefined;
	for (let count = 0; count < 4096 && cursor + 4 <= bytes.byteLength; count++) {
		if (bytes.getUint8(cursor++) !== 0xff) return undefined;
		while (cursor < bytes.byteLength && bytes.getUint8(cursor) === 0xff) cursor++;
		if (cursor + 3 > bytes.byteLength) return undefined;
		const marker = bytes.getUint8(cursor++);
		const length = bytes.getUint16(cursor);
		if (length < 2 || cursor + length > bytes.byteLength) return undefined;
		if (marker === 0xda) return header;
		if (marker >= 0xc0 && marker <= 0xcf && ![0xc4, 0xcc].includes(marker)) {
			if (header || ![0xc0, 0xc1, 0xc2].includes(marker) || length < 8) return undefined;
			const height = bytes.getUint16(cursor + 3);
			const width = bytes.getUint16(cursor + 5);
			if (!canEditRasterDimensions(width, height)) return undefined;
			header = { type: 'image/jpeg', width, height };
		}
		// Standalone markers and hierarchical/multiple-frame JPEG are outside this path.
		if (
			marker <= 0x01 ||
			(marker >= 0xd0 && marker <= 0xd9) ||
			marker === 0xde ||
			marker === 0xdf
		) {
			return undefined;
		}
		cursor += length;
	}
	return undefined;
}

export function readPictureRasterHeader(buffer: ArrayBuffer): PictureRasterHeader | undefined {
	if (buffer.byteLength < 24 || buffer.byteLength > MAX_PICTURE_EDIT_BYTES) return undefined;
	const bytes = new DataView(buffer);
	if (bytes.getUint32(0) === 0x89504e47 && bytes.getUint32(4) === 0x0d0a1a0a)
		return pngHeader(bytes);
	return bytes.getUint16(0) === 0xffd8 ? jpegHeader(bytes) : undefined;
}
