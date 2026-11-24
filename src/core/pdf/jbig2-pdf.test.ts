import { readFileSync } from 'node:fs';
import { zlibSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { decodePng } from '../drawingml/raster-decode';
import { openPdf } from './index';
import { simplePdf, stream } from './test-support/build';

const cases = JSON.parse(
	readFileSync(new URL('./__fixtures__/jbig2/jbig2.json', import.meta.url), 'utf8'),
) as {
	name: string;
	globals?: string;
	data: string;
	pages: { width: number; height: number; bits?: string }[];
}[];
const symbol = cases.find((item) => item.name === 'symbol')!;
const { width, height } = symbol.pages[0]!;
const data = new Uint8Array(Buffer.from(symbol.data, 'base64'));
const globals = new Uint8Array(Buffer.from(symbol.globals!, 'base64'));
const black = Buffer.from(symbol.pages[0]!.bits!, 'base64');
const isBlack = (x: number, y: number) =>
	((black[y * ((width + 7) >> 3) + (x >> 3)]! >> (7 - (x & 7))) & 1) === 1;

async function image(dict: string, body: Uint8Array, fill = '') {
	const pdf = simplePdf(`${fill} q 240 0 0 90 72 500 cm /Im Do Q`, {
		resources: '/XObject << /Im 6 0 R >>',
		objects: [
			stream(
				`/Type /XObject /Subtype /Image /Width ${width} /Height ${height} ${dict}`,
				body,
				true,
			),
			// The globals stream is Flate compressed, as PDF writers usually store it.
			stream('', globals),
		],
	});
	return (await (await openPdf(pdf)).page(0)).images[0]!;
}

describe('JBIG2Decode images', () => {
	it('decode with their /JBIG2Globals into pictures (1 bits black)', async () => {
		const picture = await image(
			'/ColorSpace /DeviceGray /BitsPerComponent 1 /Filter /JBIG2Decode /DecodeParms << /JBIG2Globals 7 0 R >>',
			data,
		);
		expect(picture.format).toBe('png');
		const { rgb } = decodePng(picture.data!)!;
		for (let y = 0; y < height; y++)
			for (let x = 0; x < width; x++)
				expect(rgb[(y * width + x) * 3]).toBe(isBlack(x, y) ? 0 : 255);
	});

	it('decode after other filters, as image masks and with an inverting Decode array', async () => {
		const mask = await image(
			'/ImageMask true /Decode [1 0] /Filter [/FlateDecode /JBIG2Decode] /DecodeParms [null << /JBIG2Globals 7 0 R >>]',
			zlibSync(data),
			'0 0 1 rg',
		);
		const { rgb, alpha } = decodePng(mask.data!)!;
		expect([...rgb.subarray(0, 3)]).toEqual([0, 0, 255]);
		// With /Decode [1 0] the mask paints the white pixels.
		for (let y = 0; y < height; y++)
			for (let x = 0; x < width; x++) expect(alpha![y * width + x]).toBe(isBlack(x, y) ? 0 : 255);
	});

	it('report a missing globals stream through the decoding error', async () => {
		const picture = await image(
			'/BitsPerComponent 1 /ColorSpace /DeviceGray /Filter /JBIG2Decode',
			data,
		);
		expect(picture).toMatchObject({ format: 'unsupported' });
		expect(picture.reason).toMatch(/^its JBIG2 data could not be decoded \(/);
	});
});
