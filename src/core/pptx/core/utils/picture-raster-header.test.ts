import { describe, expect, it } from 'vitest';
import { MAX_PICTURE_EDIT_BYTES, readPictureRasterHeader } from './picture-raster-header';

function png(width: number, height: number): ArrayBuffer {
	const buffer = new ArrayBuffer(45);
	const view = new DataView(buffer);
	view.setBigUint64(0, 0x89504e470d0a1a0an);
	view.setUint32(8, 13);
	view.setUint32(12, 0x49484452);
	view.setUint32(16, width);
	view.setUint32(20, height);
	view.setUint32(37, 0x49454e44);
	return buffer;
}

describe('picture decode guard', () => {
	it('reads dimensions without decoding the raster', () => {
		expect(readPictureRasterHeader(png(400, 300))).toEqual({
			type: 'image/png',
			width: 400,
			height: 300,
		});
	});

	it('rejects excessive pixel count, edge length and compressed byte size', () => {
		expect(readPictureRasterHeader(png(8000, 8000))).toBeUndefined();
		expect(readPictureRasterHeader(png(9000, 2))).toBeUndefined();
		expect(readPictureRasterHeader(new ArrayBuffer(MAX_PICTURE_EDIT_BYTES + 1))).toBeUndefined();
	});

	it('rejects animated and truncated PNG', () => {
		const animated = png(20, 20);
		new DataView(animated).setUint32(37, 0x6163544c);
		expect(readPictureRasterHeader(animated)).toBeUndefined();
		expect(readPictureRasterHeader(png(20, 20).slice(0, 27))).toBeUndefined();
	});

	it('bounds baseline and progressive JPEG dimensions before finding scan data', () => {
		const jpeg = Uint8Array.of(
			0xff,
			0xd8,
			0xff,
			0xc2,
			0,
			11,
			8,
			0,
			30,
			0,
			40,
			1,
			1,
			0x11,
			0,
			0xff,
			0xda,
			0,
			8,
			1,
			1,
			0,
			0,
			63,
			0,
			0,
			0xff,
			0xd9,
		);
		expect(readPictureRasterHeader(jpeg.buffer)).toEqual({
			type: 'image/jpeg',
			width: 40,
			height: 30,
		});
		jpeg[7] = 0xff;
		expect(readPictureRasterHeader(jpeg.buffer)).toBeUndefined();
	});
});
