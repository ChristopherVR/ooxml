import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { decodePng } from '../drawingml/raster-decode';
import { openPdf, type PdfPageImage } from './index';
import { simplePdf, stream } from './test-support/build';

const fixtures = JSON.parse(
	readFileSync(new URL('./jpx/__fixtures__/jpx.json', import.meta.url), 'utf8'),
) as {
	width: number;
	height: number;
	sources: Record<string, { data: string }>;
	cases: { name: string; data: string }[];
};
const W = fixtures.width;
const H = fixtures.height;
const bytes = (base64: string) => new Uint8Array(Buffer.from(base64, 'base64'));
const file = (name: string) => bytes(fixtures.cases.find((item) => item.name === name)!.data);
const source = (name: string) => bytes(fixtures.sources[name]!.data);

const xobject = (dict: string, data: Uint8Array, objects: Uint8Array[] = []) =>
	simplePdf('q 50 0 0 50 100 100 cm /Im Do Q', {
		resources: '/XObject << /Im 6 0 R >>',
		objects: [stream(`/Type /XObject /Subtype /Image ${dict}`, data, true), ...objects],
	});
const image = async (pdf: Uint8Array): Promise<PdfPageImage> =>
	(await (await openPdf(pdf)).page(0)).images[0]!;
const jpx = (name: string, dict = '') =>
	image(xobject(`/Width ${W} /Height ${H} /Filter /JPXDecode ${dict}`, file(name)));
const pixels = (picture: PdfPageImage) => {
	expect(picture.format).toBe('png');
	return decodePng(picture.data!)!;
};

const u32 = (value: number) => [
	value >>> 24,
	(value >> 16) & 0xff,
	(value >> 8) & 0xff,
	value & 0xff,
];
const box = (type: string, body: number[] | Uint8Array) => [
	...u32(8 + body.length),
	...[...type].map((char) => char.charCodeAt(0)),
	...body,
];

/** A JP2 file around a codestream, with the header boxes given. */
function jp2(codestream: Uint8Array, components: number, boxes: number[][]): Uint8Array {
	const ihdr = box('ihdr', [...u32(H), ...u32(W), 0, components, 7, 7, 0, 0]);
	return Uint8Array.from([
		...box('jP  ', [0x0d, 0x0a, 0x87, 0x0a]),
		...box(
			'ftyp',
			[...'jp2 '].map((c) => c.charCodeAt(0)).concat([0, 0, 0, 0, 0x6a, 0x70, 0x32, 0x20]),
		),
		...box('jp2h', [...ihdr, ...boxes.flat()]),
		...box('jp2c', codestream),
	]);
}
const colr = (space: number) => box('colr', [1, 0, 0, ...u32(space)]);

describe('JPXDecode images', () => {
	it('decodes RGB pictures in the colour space of the JP2 header', async () => {
		const { rgb, alpha } = pixels(await jpx('rgb-53-tiles'));
		const expected = source('rgb8');
		for (let i = 0; i < W * H; i++)
			for (let c = 0; c < 3; c++) expect(rgb[i * 3 + c]).toBe(expected[c * W * H + i]);
		expect(alpha).toBeUndefined();
	});

	it('ignores /Decode and /BitsPerComponent, and scales deep and signed samples', async () => {
		const gray = source('gray8');
		const kept = pixels(
			await jpx('gray-53', '/ColorSpace /DeviceGray /BitsPerComponent 1 /Decode [1 0]'),
		);
		expect([...kept.rgb.filter((_, i) => i % 3 === 0)]).toEqual([...gray]);
		const deep = pixels(await jpx('gray16'));
		const wide = source('gray16');
		for (let i = 0; i < W * H; i++)
			expect(deep.rgb[i * 3]).toBe(
				Math.round((((wide[i * 2]! << 8) | wide[i * 2 + 1]!) * 255) / 65535),
			);
		const signed = pixels(await jpx('signed8', '/ColorSpace /DeviceGray'));
		const values = source('signed8');
		for (let i = 0; i < W * H; i++) expect(signed.rgb[i * 3]).toBe((values[i]! + 128) & 0xff);
	});

	it('uses the embedded opacity only with /SMaskInData', async () => {
		const expected = source('rgba8');
		const without = pixels(await jpx('rgba'));
		expect(without.alpha).toBeUndefined();
		const masked = pixels(await jpx('rgba', '/SMaskInData 1'));
		expect([...masked.alpha!]).toEqual([...expected.subarray(3 * W * H)]);
		expect(masked.rgb[0]).toBe(expected[0]);
	});

	it('spreads sub-sampled components and converts sYCC', async () => {
		const planes = source('yuv420');
		const cw = Math.ceil(W / 2);
		const at = (x: number, y: number) => Math.floor(y / 2) * cw + Math.floor(x / 2);
		const raw = pixels(await jpx('yuv420-cprl'));
		for (const [x, y] of [
			[0, 0],
			[5, 3],
			[32, 16],
			[17, 9],
		] as const) {
			const i = y * W + x;
			expect(raw.rgb[i * 3]).toBe(planes[i]);
			expect(raw.rgb[i * 3 + 1]).toBe(planes[W * H + at(x, y)]);
			expect(raw.rgb[i * 3 + 2]).toBe(planes[W * H + cw * Math.ceil(H / 2) + at(x, y)]);
		}
		const codestream = file('yuv420-cprl');
		const ycc = pixels(
			await image(
				xobject(`/Width ${W} /Height ${H} /Filter /JPXDecode`, jp2(codestream, 3, [colr(18)])),
			),
		);
		const y = planes[0]!;
		const cr = planes[W * H + cw * Math.ceil(H / 2)]! - 128;
		expect(ycc.rgb[0]).toBe(Math.min(255, Math.max(0, Math.round(y + 1.402 * cr))));
	});

	it('applies the JP2 palette, unless the PDF colour space is Indexed', async () => {
		const codestream = file('gray-53');
		const entries = Array.from({ length: 256 }, (_, i) => [i, 255 - i, i >> 1]).flat();
		const palette = box('pclr', [1, 0, 3, 7, 7, 7, ...entries]);
		const cmap = box('cmap', [0, 0, 1, 0, 0, 0, 1, 1, 0, 0, 1, 2]);
		const data = jp2(codestream, 1, [colr(16), palette, cmap]);
		const gray = source('gray8');
		const mapped = pixels(
			await image(xobject(`/Width ${W} /Height ${H} /Filter /JPXDecode`, data)),
		);
		for (const i of [0, 40, 300, W * H - 1])
			expect([...mapped.rgb.subarray(i * 3, i * 3 + 3)]).toEqual([
				gray[i],
				255 - gray[i]!,
				gray[i]! >> 1,
			]);
		const lookup = Array.from({ length: 256 }, (_, i) =>
			(255 - i).toString(16).padStart(2, '0').repeat(3),
		).join('');
		const indexed = pixels(
			await image(
				xobject(
					`/Width ${W} /Height ${H} /ColorSpace [/Indexed /DeviceRGB 255 <${lookup}>] /Filter /JPXDecode`,
					data,
				),
			),
		);
		expect(indexed.rgb[300 * 3]).toBe(255 - gray[300]!);
	});

	it('reads JPEG 2000 soft masks and reports damaged data', async () => {
		const gray = source('gray8');
		const masked = pixels(
			await image(
				xobject(
					`/Width ${W} /Height ${H} /ColorSpace /DeviceGray /BitsPerComponent 8 /SMask 7 0 R`,
					new Uint8Array(W * H).fill(128),
					[
						stream(
							`/Type /XObject /Subtype /Image /Width ${W} /Height ${H} /ColorSpace /DeviceGray /Filter /JPXDecode`,
							file('gray-53'),
							true,
						),
					],
				),
			),
		);
		expect([...masked.alpha!]).toEqual([...gray]);
		const damaged = file('gray-53').slice(0, 30);
		const picture = await image(xobject(`/Width ${W} /Height ${H} /Filter /JPXDecode`, damaged));
		expect(picture).toMatchObject({ format: 'unsupported' });
		expect(picture.reason).toMatch(/JPEG 2000/);
	});
});
