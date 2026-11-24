import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { decodePng } from '../drawingml/raster-decode';
import { openPdf, type PdfPageImage } from './index';
import { adobeTransform, cmykJpegForDocument } from './jpeg-cmyk';
import { simplePdf, stream } from './test-support/build';

const latin1 = (bytes: Uint8Array) => String.fromCharCode(...bytes);
const ccitt = (
	JSON.parse(readFileSync(new URL('./__fixtures__/ccitt.json', import.meta.url), 'utf8')) as {
		name: string;
		data: string;
		expected: string;
		width: number;
		height: number;
	}[]
).find((item) => item.name === 'sample group4')!;
const cmykJpeg = new Uint8Array(readFileSync(new URL('./__fixtures__/cmyk.jpg', import.meta.url)));

async function images(
	content: string,
	extra: { resources?: string; objects?: (string | Uint8Array)[] } = {},
): Promise<PdfPageImage[]> {
	return (await (await openPdf(simplePdf(content, extra))).page(0)).images;
}
const pixels = (image: PdfPageImage) => decodePng(image.data!)!;
const xobject = (dict: string, data: Uint8Array | string) =>
	simplePdf('q 50 0 0 50 100 100 cm /Im Do Q', {
		resources: '/XObject << /Im 6 0 R >>',
		objects: [stream(`/Type /XObject /Subtype /Image ${dict}`, data, true)],
	});
const one = async (bytes: Uint8Array) => (await (await openPdf(bytes)).page(0)).images[0]!;

describe('inline images', () => {
	it('decodes abbreviated keys and data that contains EI, and reads on after it', async () => {
		const reader = await openPdf(
			simplePdf(
				`q 20 0 0 20 100 100 cm BI /W 2 /H 2 /CS /RGB /BPC 8 ID ${latin1(Uint8Array.of(255, 0, 0, 32, 69, 73, 32, 0, 255, 0, 0, 255))}\nEI Q BT /F1 12 Tf 72 700 Td (After) Tj ET`,
			),
		);
		const page = await reader.page(0);
		expect(page.spans.map((span) => span.text)).toEqual(['After']);
		expect(page.images).toHaveLength(1);
		expect(page.images[0]).toMatchObject({ format: 'png', pixelWidth: 2, x: 100, width: 20 });
		const { rgb } = pixels(page.images[0]!);
		expect([...rgb.subarray(0, 6)]).toEqual([255, 0, 0, 32, 69, 73]);
		expect(page.notes).toEqual([]);
	});

	it('undoes filters and reads named and indexed colour spaces with a Decode array', async () => {
		const [hex, named, indexed] = await images(
			'q 10 0 0 10 0 0 cm BI /W 2 /H 1 /CS /G /BPC 8 /F /AHx ID 00FF> EI Q ' +
				'q 10 0 0 10 0 0 cm BI /W 1 /H 1 /CS /CS0 /BPC 8 ID \u0000\u0000\u00ff EI Q ' +
				'q 10 0 0 10 0 0 cm BI /W 2 /H 1 /CS [/I /RGB 1 <FF0000 00FF00>] /BPC 1 /D [1 0] ID \u0040 EI Q',
			{ resources: '/ColorSpace << /CS0 /DeviceRGB >>' },
		);
		expect([...pixels(hex!).rgb]).toEqual([0, 0, 0, 255, 255, 255]);
		expect([...pixels(named!).rgb]).toEqual([0, 0, 255]);
		// Bits 0 and 1, inverted by the Decode array: palette entries 1 then 0.
		expect([...pixels(indexed!).rgb]).toEqual([0, 255, 0, 255, 0, 0]);
	});
});

describe('image codecs and masks', () => {
	it('decodes CCITT Group 4 pictures and paints stencil masks in the fill colour', async () => {
		const data = new Uint8Array(Buffer.from(ccitt.data, 'base64'));
		const parms = `/Filter /CCITTFaxDecode /DecodeParms << /K -1 /Columns ${ccitt.width} /Rows ${ccitt.height} >>`;
		const image = await one(
			xobject(
				`/Width ${ccitt.width} /Height ${ccitt.height} /ColorSpace /DeviceGray /BitsPerComponent 1 ${parms}`,
				data,
			),
		);
		expect(image.format).toBe('png');
		const { rgb } = pixels(image);
		const expected = Buffer.from(ccitt.expected, 'base64');
		const stride = Math.ceil(ccitt.width / 8);
		for (let y = 0; y < ccitt.height; y++)
			for (let x = 0; x < ccitt.width; x++) {
				const white = (expected[y * stride + (x >> 3)]! >> (7 - (x & 7))) & 1;
				expect(rgb[(y * ccitt.width + x) * 3]).toBe(white ? 255 : 0);
			}
		const mask = await (
			await openPdf(
				simplePdf('1 0 0 rg q 50 0 0 50 100 100 cm /Im Do Q', {
					resources: '/XObject << /Im 6 0 R >>',
					objects: [
						stream(
							`/Type /XObject /Subtype /Image /Width ${ccitt.width} /Height ${ccitt.height} /ImageMask true ${parms}`,
							data,
							true,
						),
					],
				}),
			)
		).page(0);
		const decoded = pixels(mask.images[0]!);
		expect([...decoded.rgb.subarray(0, 3)]).toEqual([255, 0, 0]);
		expect(decoded.alpha).toBeDefined();
	});

	it('keeps CMYK JPEGs, with the Adobe marker matching the Decode array', async () => {
		expect(adobeTransform(cmykJpeg)).toBe(0);
		const kept = await one(
			xobject(
				'/Width 16 /Height 8 /ColorSpace /DeviceCMYK /BitsPerComponent 8 /Filter /DCTDecode',
				cmykJpeg,
			),
		);
		expect(kept).toMatchObject({ format: 'jpeg', data: cmykJpeg });
		const inverted = await one(
			xobject(
				'/Width 16 /Height 8 /ColorSpace /DeviceCMYK /BitsPerComponent 8 /Decode [1 0 1 0 1 0 1 0] /Filter /DCTDecode',
				cmykJpeg,
			),
		);
		expect(inverted.format).toBe('jpeg');
		expect(adobeTransform(inverted.data!)).toBeUndefined();
		expect(inverted.data!.length).toBe(cmykJpeg.length - 16);
		const restored = cmykJpegForDocument(inverted.data!, true)!;
		expect(adobeTransform(restored)).toBe(0);
	});

	it('applies Decode arrays, soft masks with a matte colour and colour key masks', async () => {
		const decoded = await one(
			xobject(
				'/Width 2 /Height 1 /ColorSpace /DeviceGray /BitsPerComponent 8 /Decode [1 0]',
				'\u0000\u00c0',
			),
		);
		expect([...pixels(decoded).rgb]).toEqual([255, 255, 255, 63, 63, 63]);
		const matte = await (
			await openPdf(
				simplePdf('q 50 0 0 50 100 100 cm /Im Do Q', {
					resources: '/XObject << /Im 6 0 R >>',
					objects: [
						stream(
							'/Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceGray /BitsPerComponent 8 /SMask 7 0 R',
							'\u0080',
							true,
						),
						stream(
							'/Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceGray /BitsPerComponent 8 /Matte [1]',
							'\u0080',
							true,
						),
					],
				}),
			)
		).page(0);
		const premultiplied = pixels(matte.images[0]!);
		expect(premultiplied.rgb[0]).toBeLessThan(6);
		expect(premultiplied.alpha![0]).toBe(128);
		const keyed = pixels(
			await one(
				xobject(
					'/Width 2 /Height 1 /ColorSpace /DeviceGray /BitsPerComponent 8 /Mask [0 10]',
					'\u0005\u00ff',
				),
			),
		);
		expect([...keyed.alpha!]).toEqual([0, 255]);
	});

	it('reports JBIG2 and JPEG 2000 pictures with the reason', async () => {
		expect(
			await one(xobject('/Width 4 /Height 4 /BitsPerComponent 1 /Filter /JBIG2Decode', 'xx')),
		).toMatchObject({
			format: 'unsupported',
			reason: 'its JBIG2 compression is not supported',
		});
		expect((await one(xobject('/Width 4 /Height 4 /Filter /JPXDecode', 'xx'))).reason).toBe(
			'its JPEG 2000 compression is not supported',
		);
	});
});
