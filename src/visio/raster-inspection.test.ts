import JSZip from 'jszip';
import { describe, expect, it, vi } from 'vitest';
import {
	inspectVisioRasterImage,
	VisioImageError,
	VISIO_RASTER_IMAGE_LIMITS,
	type VisioImageOptions,
} from './index.js';
import { readVisioImage } from './media.js';
import { crc32 } from './package-common.js';
import { VisioPackage } from './package.js';
import { child, type Report } from './sheet.js';
import { relations, xml } from './test-fixtures.js';

const png = Uint8Array.from(
	Buffer.from(
		'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAABHNCSVQICAgIfAhkiAAAAAFzUkdCAK7OHOkAAAALSURBVAiZY2AAAgAABQABYlUyiAAAAABJRU5ErkJggg==',
		'base64',
	),
);
const gif = Uint8Array.from(
	Buffer.from('R0lGODlhAQABAIAAAAAAAP///ywAAAAAAQABAAACAUwAOw==', 'base64'),
);
const jpeg = Uint8Array.from(
	Buffer.from(
		'/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAAMCAgICAgMCAgIDAwMDBAYEBAQEBAgGBgUGCQgKCgkICQkKDA8MCgsOCwkJDRENDg8QEBEQCgwSExIQEw8QEBD/2wBDAQMDAwQDBAgEBAgQCwkLEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBD/wAARCAADAAIDASIAAhEBAxEB/8QAFQABAQAAAAAAAAAAAAAAAAAAAAn/xAAUEAEAAAAAAAAAAAAAAAAAAAAA/8QAFAEBAAAAAAAAAAAAAAAAAAAAAP/EABQRAQAAAAAAAAAAAAAAAAAAAAD/2gAMAwEAAhEDEQA/AJVAA//Z',
		'base64',
	),
);
const fixtures = [
	['PNG', png, 'image/png', 1, 1],
	['GIF', gif, 'image/gif', 1, 1],
	['JPEG', jpeg, 'image/jpeg', 2, 3],
] as const;

function pngSize(width: number, height: number): Uint8Array {
	const bytes = png.slice(),
		view = new DataView(bytes.buffer);
	view.setUint32(16, width);
	view.setUint32(20, height);
	view.setUint32(29, (crc32(bytes.subarray(12, 29)) ^ 0xffffffff) >>> 0);
	return bytes;
}

describe('public Visio raster inspection', () => {
	it.each(fixtures)(
		'detects %s MIME and intrinsic dimensions without retaining bytes',
		(_, bytes, mimeType, pixelWidth, pixelHeight) => {
			const before = bytes.slice();
			expect(inspectVisioRasterImage(bytes)).toEqual({ mimeType, pixelWidth, pixelHeight });
			expect(bytes).toEqual(before);
		},
	);
	it.each(fixtures)(
		'reads only a nonzero-offset %s view',
		(_, bytes, mimeType, pixelWidth, pixelHeight) => {
			const allocation = new Uint8Array(bytes.length + 32).fill(0xff);
			allocation.set(bytes, 13);
			const view = allocation.subarray(13, 13 + bytes.length);
			expect(inspectVisioRasterImage(view, { maxImageBytes: bytes.length })).toEqual({
				mimeType,
				pixelWidth,
				pixelHeight,
			});
		},
	);
	it('returns intrinsic dimensions so a host can reject mismatched declared dimensions', () => {
		const declared = { mimeType: 'image/png', pixelWidth: 10, pixelHeight: 20 };
		const detected = inspectVisioRasterImage(png);
		expect(detected.mimeType).toBe(declared.mimeType);
		expect(detected.pixelWidth).not.toBe(declared.pixelWidth);
		expect(detected.pixelHeight).not.toBe(declared.pixelHeight);
		expect(inspectVisioRasterImage(jpeg).mimeType).not.toBe(declared.mimeType);
	});
	it('rechecks mutable bytes on every call instead of caching by identity', () => {
		const bytes = png.slice();
		expect(inspectVisioRasterImage(bytes).mimeType).toBe('image/png');
		bytes[29]! ^= 1;
		expect(() => inspectVisioRasterImage(bytes)).toThrowError(
			expect.objectContaining({ code: 'invalid-image', message: 'PNG chunk checksum mismatch.' }),
		);
		bytes.set(pngSize(2, 3));
		expect(inspectVisioRasterImage(bytes)).toMatchObject({ pixelWidth: 2, pixelHeight: 3 });
	});
	it.each([
		'<svg xmlns="http://www.w3.org/2000/svg" onload="alert(1)"/>',
		'<!DOCTYPE html><html><script>alert(1)</script></html>',
		'GIF89b',
		'',
	])('rejects unsupported or active content regardless of a host MIME claim: %s', (content) => {
		expect(() => inspectVisioRasterImage(new TextEncoder().encode(content))).toThrowError(
			expect.objectContaining({ name: 'VisioImageError', code: 'unsupported-image' }),
		);
	});
	it.each(fixtures)('rejects a %s signature followed by HTML', (_, bytes) => {
		const signature = bytes.subarray(0, bytes === png ? 8 : bytes === gif ? 6 : 2);
		const html = new TextEncoder().encode('<html><script>alert(1)</script></html>');
		expect(() => inspectVisioRasterImage(Uint8Array.from([...signature, ...html]))).toThrowError(
			VisioImageError,
		);
	});
	it.each(fixtures)('rejects truncated recognized %s headers and payloads', (_, bytes) => {
		const signatureLength = bytes === png ? 8 : bytes === gif ? 6 : 2;
		for (const end of [signatureLength, Math.floor(bytes.length / 2), bytes.length - 1])
			expect(() => inspectVisioRasterImage(bytes.subarray(0, end))).toThrowError(
				expect.objectContaining({ code: 'invalid-image' }),
			);
	});
	it.each(fixtures)('rejects data appended after the %s terminator', (_, bytes) => {
		const content = Uint8Array.from([...bytes, ...new TextEncoder().encode('<svg/>')]);
		expect(() => inspectVisioRasterImage(content)).toThrowError(
			expect.objectContaining({ code: 'invalid-image' }),
		);
	});
	it('checks PNG payload checksums as well as the header checksum', () => {
		const bytes = png.slice();
		const at = Buffer.from(bytes).indexOf('IDAT');
		expect(at).toBeGreaterThan(0);
		bytes[at + 4]! ^= 1;
		expect(() => inspectVisioRasterImage(bytes)).toThrowError(
			expect.objectContaining({ code: 'invalid-image', message: 'PNG chunk checksum mismatch.' }),
		);
	});
	it('rejects corrupt GIF descriptors and JPEG segment markers', () => {
		const badGif = gif.slice();
		badGif[badGif.indexOf(0x2c)] = 0;
		const badJpeg = jpeg.slice();
		badJpeg[2] = 0;
		for (const bytes of [badGif, badJpeg, pngSize(0, 1)])
			expect(() => inspectVisioRasterImage(bytes)).toThrowError(
				expect.objectContaining({ code: 'invalid-image' }),
			);
	});
	it.each([null, undefined, 'PNG', new ArrayBuffer(16), new Uint16Array(8)])(
		'rejects a non-Uint8Array runtime input with a structured error',
		(bytes) => {
			expect(() => inspectVisioRasterImage(bytes as unknown as Uint8Array)).toThrowError(
				expect.objectContaining({ code: 'invalid-image' }),
			);
		},
	);
});

describe('public Visio raster inspection limits', () => {
	it('exposes frozen default limits that are also hard inspection ceilings', () => {
		expect(VISIO_RASTER_IMAGE_LIMITS).toEqual({
			maxImageBytes: 8 * 1024 * 1024,
			maxPixels: 16_000_000,
			maxDimension: 16384,
		});
		expect(Object.isFrozen(VISIO_RASTER_IMAGE_LIMITS)).toBe(true);
		expect(inspectVisioRasterImage(png, { ...VISIO_RASTER_IMAGE_LIMITS })).toBeDefined();
	});
	it.each(['maxImageBytes', 'maxPixels', 'maxDimension'] as const)(
		'rejects invalid or raised %s options rather than weakening the hard ceiling',
		(key) => {
			for (const value of [
				0,
				-1,
				0.5,
				NaN,
				Infinity,
				Number.MAX_SAFE_INTEGER + 1,
				VISIO_RASTER_IMAGE_LIMITS[key] + 1,
				undefined,
				null,
				'1',
			])
				expect(() => inspectVisioRasterImage(png, { [key]: value })).toThrowError(
					expect.objectContaining({ code: 'invalid-limit' }),
				);
		},
	);
	it.each([null, 1, 'limits', []])('rejects an invalid limits object', (options) => {
		expect(() => inspectVisioRasterImage(png, options as VisioImageOptions)).toThrowError(
			expect.objectContaining({ code: 'invalid-limit' }),
		);
	});
	it.each([{ maxImageBytes: jpeg.length - 1 }, { maxPixels: 5 }, { maxDimension: 2 }])(
		'honors lower caller bounds %j',
		(options) => {
			expect(() => inspectVisioRasterImage(jpeg, options)).toThrowError(
				expect.objectContaining({ code: 'image-limit' }),
			);
		},
	);
	it('accepts exact lower bounds inclusively', () => {
		expect(
			inspectVisioRasterImage(jpeg, {
				maxImageBytes: jpeg.length,
				maxPixels: 6,
				maxDimension: 3,
			}),
		).toEqual({ mimeType: 'image/jpeg', pixelWidth: 2, pixelHeight: 3 });
	});
	it.each([pngSize(16385, 1), pngSize(4001, 4000), pngSize(0xffffffff, 0xffffffff)])(
		'enforces default dimension and pixel ceilings',
		(bytes) => {
			expect(() => inspectVisioRasterImage(bytes)).toThrowError(
				expect.objectContaining({ code: 'image-limit' }),
			);
		},
	);
	it('applies the byte ceiling before inspecting format structure', () => {
		const bytes = new Uint8Array(VISIO_RASTER_IMAGE_LIMITS.maxImageBytes + 1);
		expect(() => inspectVisioRasterImage(bytes)).toThrowError(
			expect.objectContaining({ code: 'image-limit' }),
		);
	});
	it('counts all GIF frames toward the pixel ceiling', () => {
		const start = gif.indexOf(0x2c);
		const bytes = Uint8Array.from([...gif.subarray(0, -1), ...gif.subarray(start)]);
		expect(() => inspectVisioRasterImage(bytes, { maxPixels: 1 })).toThrowError(
			expect.objectContaining({ code: 'image-limit' }),
		);
	});
	it('preserves package parser options that raise limits above inspection ceilings', async () => {
		const zip = new JSZip(),
			source = 'visio/pages/page1.xml';
		zip.file(source, xml('PageContents', '<ForeignData><Rel r:id="rId1"/></ForeignData>'));
		zip.file(
			'visio/pages/_rels/page1.xml.rels',
			relations(
				'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image.png"/>',
			),
		);
		zip.file('visio/media/image.png', pngSize(16385, 1000));
		const pkg = await VisioPackage.open(await zip.generateAsync({ type: 'uint8array' }));
		const report = vi.fn<Report>();
		const image = await readVisioImage(
			pkg,
			source,
			child(await pkg.readXml(source), 'ForeignData')!,
			report,
			{
				maxImageBytes: VISIO_RASTER_IMAGE_LIMITS.maxImageBytes + 1,
				maxDimension: 16385,
				maxPixels: 16_385_000,
			},
		);
		expect(image).toMatchObject({ mimeType: 'image/png', pixelWidth: 16385, pixelHeight: 1000 });
		expect(report).not.toHaveBeenCalled();
	});
});
