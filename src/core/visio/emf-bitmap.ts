// A minimal enhanced metafile that draws one raster: header, one EMR_STRETCHDIBITS with a 24-bit
// bottom-up DIB, end of file. Visio keeps the preview of an embedded Office object as an EMF, and
// a browser can produce pixels but not a metafile, so this wraps them. No compression, no
// palette, no state records: any EMF reader plays it back as the bitmap at the frame size.

const HEADER_SIZE = 88;
const STRETCH_HEADER = 80;
const DIB_HEADER = 40;
const EOF_SIZE = 20;
const MAX_SIDE = 8192;
const PIXELS_PER_MM = 96 / 25.4;

const EMR_HEADER = 1;
const EMR_EOF = 14;
const EMR_STRETCHDIBITS = 81;
const SRCCOPY = 0x00cc0020;
const EMF_SIGNATURE = 0x464d4520;

export class EmfBitmapError extends Error {}

/**
 * The EMF for `width` x `height` RGBA pixels (row-major, top row first). Transparent pixels are
 * flattened onto white, as EMF bitmaps carry no alpha here. Frame and device sizes assume 96 dpi.
 */
export function encodeBitmapEmf(
	width: number,
	height: number,
	rgba: Uint8Array | Uint8ClampedArray,
): Uint8Array {
	if (!Number.isInteger(width) || !Number.isInteger(height) || width < 1 || height < 1)
		throw new EmfBitmapError('The bitmap needs a positive whole width and height.');
	if (width > MAX_SIDE || height > MAX_SIDE)
		throw new EmfBitmapError(`The bitmap is larger than ${MAX_SIDE} pixels on a side.`);
	if (rgba.length !== width * height * 4)
		throw new EmfBitmapError('The pixel data does not match the width and height.');
	const rowBytes = Math.ceil((width * 3) / 4) * 4;
	const bitsSize = rowBytes * height;
	const stretchSize = STRETCH_HEADER + DIB_HEADER + bitsSize;
	const total = HEADER_SIZE + stretchSize + EOF_SIZE;
	const out = new Uint8Array(total);
	const view = new DataView(out.buffer);
	const u32 = (at: number, value: number) => view.setUint32(at, value >>> 0, true);
	const i32 = (at: number, value: number) => view.setInt32(at, value, true);

	// EMR_HEADER (88 bytes): bounds in device pixels (inclusive), frame in 0.01 mm.
	const mmWidth = width / PIXELS_PER_MM;
	const mmHeight = height / PIXELS_PER_MM;
	u32(0, EMR_HEADER);
	u32(4, HEADER_SIZE);
	i32(8, 0);
	i32(12, 0);
	i32(16, width - 1);
	i32(20, height - 1);
	i32(24, 0);
	i32(28, 0);
	i32(32, Math.round(mmWidth * 100));
	i32(36, Math.round(mmHeight * 100));
	u32(40, EMF_SIGNATURE);
	u32(44, 0x00010000);
	u32(48, total);
	u32(52, 3); // records: header, bitmap, end
	view.setUint16(56, 1, true); // handles: the stock object slot
	u32(60, 0); // nDescription
	u32(64, 0); // offDescription
	u32(68, 0); // nPalEntries
	i32(72, width); // szlDevice
	i32(76, height);
	i32(80, Math.max(1, Math.round(mmWidth))); // szlMillimeters
	i32(84, Math.max(1, Math.round(mmHeight)));

	// EMR_STRETCHDIBITS
	const s = HEADER_SIZE;
	u32(s, EMR_STRETCHDIBITS);
	u32(s + 4, stretchSize);
	i32(s + 8, 0);
	i32(s + 12, 0);
	i32(s + 16, width - 1);
	i32(s + 20, height - 1);
	i32(s + 24, 0); // xDest
	i32(s + 28, 0); // yDest
	i32(s + 32, 0); // xSrc
	i32(s + 36, 0); // ySrc
	i32(s + 40, width); // cxSrc
	i32(s + 44, height); // cySrc
	u32(s + 48, STRETCH_HEADER); // offBmiSrc
	u32(s + 52, DIB_HEADER); // cbBmiSrc
	u32(s + 56, STRETCH_HEADER + DIB_HEADER); // offBitsSrc
	u32(s + 60, bitsSize); // cbBitsSrc
	u32(s + 64, 0); // DIB_RGB_COLORS
	u32(s + 68, SRCCOPY);
	i32(s + 72, width); // cxDest
	i32(s + 76, height); // cyDest
	// BITMAPINFOHEADER
	const b = s + STRETCH_HEADER;
	u32(b, DIB_HEADER);
	i32(b + 4, width);
	i32(b + 8, height); // positive: rows are stored bottom-up
	view.setUint16(b + 12, 1, true);
	view.setUint16(b + 14, 24, true);
	u32(b + 16, 0); // BI_RGB
	u32(b + 20, bitsSize);
	// Pixels, bottom row first, BGR, composited over white.
	const bits = b + DIB_HEADER;
	for (let y = 0; y < height; y++) {
		const source = (height - 1 - y) * width * 4;
		let target = bits + y * rowBytes;
		for (let x = 0; x < width; x++) {
			const alpha = rgba[source + x * 4 + 3]! / 255;
			for (const channel of [2, 1, 0])
				out[target++] = Math.round(rgba[source + x * 4 + channel]! * alpha + 255 * (1 - alpha));
		}
	}
	// EMR_EOF
	const e = HEADER_SIZE + stretchSize;
	u32(e, EMR_EOF);
	u32(e + 4, EOF_SIZE);
	u32(e + 8, 0);
	u32(e + 12, 16);
	u32(e + 16, EOF_SIZE);
	return out;
}
