import JSZip from 'jszip';
import { describe, expect, it, vi } from 'vitest';
import { crc32 } from './package-common.js';
import { VisioPackage } from './package.js';
import { readVisioImage, type VisioImageOptions } from './media.js';
import { child, type Report } from './sheet.js';
import { xml, relations } from './test-fixtures.js';

const png = Uint8Array.from(
	Buffer.from(
		'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAABHNCSVQICAgIfAhkiAAAAAFzUkdCAK7OHOkAAAALSURBVAiZY2AAAgAABQABYlUyiAAAAABJRU5ErkJggg==',
		'base64',
	),
);
const gif = Uint8Array.from(
	Buffer.from('R0lGODlhAQABAIAAAAAAAP///ywAAAAAAQABAAACAUwAOw==', 'base64'),
);
// A generated 2 x 3 JPEG; no native decoder or external fixtures are needed at test time.
const jpeg = Uint8Array.from(
	Buffer.from(
		'/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAAMCAgICAgMCAgIDAwMDBAYEBAQEBAgGBgUGCQgKCgkICQkKDA8MCgsOCwkJDRENDg8QEBEQCgwSExIQEw8QEBD/2wBDAQMDAwQDBAgEBAgQCwkLEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBD/wAARCAADAAIDASIAAhEBAxEB/8QAFQABAQAAAAAAAAAAAAAAAAAAAAn/xAAUEAEAAAAAAAAAAAAAAAAAAAAA/8QAFAEBAAAAAAAAAAAAAAAAAAAAAP/EABQRAQAAAAAAAAAAAAAAAAAAAAD/2gAMAwEAAhEDEQA/AJVAA//Z',
		'base64',
	),
);
const source = 'visio/pages/page1.xml';
const media = 'visio/media/image.png';
const imageType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';
interface Setup {
	bytes?: Uint8Array;
	type?: string;
	target?: string;
	mode?: string;
	reference?: string;
	foreignType?: string;
	relationships?: string;
	edit?: (zip: JSZip) => void;
}
async function setup(options: Setup = {}) {
	const zip = new JSZip();
	zip.file(
		source,
		xml(
			'PageContents',
			`<ForeignData ForeignType="${options.foreignType ?? 'Bitmap'}">${options.reference ?? '<Rel r:id="rId1"/>'}</ForeignData>`,
		),
	);
	zip.file(
		'visio/pages/_rels/page1.xml.rels',
		relations(
			options.relationships ??
				`<Relationship Id="rId1" Type="${options.type ?? imageType}" Target="${options.target ?? '../media/image.png'}" TargetMode="${options.mode ?? 'Internal'}"/>`,
		),
	);
	zip.file(media, options.bytes ?? png);
	options.edit?.(zip);
	const archive = await zip.generateAsync({ type: 'uint8array', compression: 'STORE' });
	const pkg = await VisioPackage.open(archive);
	const foreignData = child(await pkg.readXml(source), 'ForeignData')!;
	const report = vi.fn<Report>();
	return { pkg, foreignData, report, archive };
}
async function read(options: Setup = {}, limits?: VisioImageOptions) {
	const fixture = await setup(options);
	const image = await readVisioImage(
		fixture.pkg,
		source,
		fixture.foreignData,
		fixture.report,
		limits,
	);
	return { ...fixture, image };
}
function pngSize(width: number, height: number): Uint8Array {
	const bytes = png.slice(),
		view = new DataView(bytes.buffer);
	view.setUint32(16, width);
	view.setUint32(20, height);
	view.setUint32(29, (crc32(bytes.subarray(12, 29)) ^ 0xffffffff) >>> 0);
	return bytes;
}
function jpegSize(width: number, height: number): Uint8Array {
	const bytes = jpeg.slice(),
		view = new DataView(bytes.buffer);
	const at = bytes.findIndex((byte, index) => byte === 0xff && bytes[index + 1] === 0xc0);
	view.setUint16(at + 5, height);
	view.setUint16(at + 7, width);
	return bytes;
}

describe('safe Visio raster images', () => {
	it('uses declared part size to reject lowered raster limits before materialization', async () => {
		const { pkg, foreignData, report } = await setup();
		const read = vi.spyOn(pkg, 'readBytes');
		expect(
			await readVisioImage(pkg, source, foreignData, report, { maxImageBytes: 8 }),
		).toBeUndefined();
		expect(read).not.toHaveBeenCalledWith(media);
		expect(report).toHaveBeenCalledWith(
			'image-limit',
			expect.stringContaining('before inflation'),
			expect.any(Object),
		);
	});

	it.each([
		['PNG', png, 'image/png', 1, 1],
		['GIF', gif, 'image/gif', 1, 1],
		['JPEG', jpeg, 'image/jpeg', 2, 3],
	] as const)(
		'loads valid %s by bytes, not the filename',
		async (_, bytes, mimeType, pixelWidth, pixelHeight) => {
			const { image, report } = await read({ bytes });
			expect(image).toMatchObject({ mimeType, pixelWidth, pixelHeight });
			expect(image?.bytes).toEqual(bytes);
			expect(report).not.toHaveBeenCalled();
		},
	);
	it('returns the same cached payload for concurrent repeated image instances', async () => {
		const { pkg, foreignData, report } = await setup();
		const [a, b] = await Promise.all([
			readVisioImage(pkg, source, foreignData, report),
			readVisioImage(pkg, source, foreignData, report),
		]);
		expect(a).toBe(b);
		expect(a?.bytes).toBe(b?.bytes);
		expect(a?.bytes).toBe(await pkg.readBytes(media));
	});
	it('keeps cached validation separate for different limits and packages', async () => {
		const { pkg, foreignData, report } = await setup({ bytes: jpeg });
		const image = await readVisioImage(pkg, source, foreignData, report);
		expect(image).toBeDefined();
		expect(
			await readVisioImage(pkg, source, foreignData, report, { maxPixels: 5 }),
		).toBeUndefined();
		expect(await readVisioImage(pkg, source, foreignData, report)).toBe(image);
		const other = await read();
		expect(other.image?.mimeType).toBe('image/png');
		expect(other.image?.bytes).not.toBe(image?.bytes);
	});
	it('shares the binary cache with XML without counting a part twice', async () => {
		const zip = new JSZip(),
			bytes = new TextEncoder().encode('<Root/>');
		zip.file('root.xml', bytes);
		const pkg = await VisioPackage.open(await zip.generateAsync({ type: 'uint8array' }), {
			maxTotalBytes: bytes.length,
		});
		expect((await pkg.readXml('root.xml')).localName).toBe('Root');
		expect(await pkg.readBytes('root.xml')).toEqual(bytes);
		expect(await pkg.readBytes('root.xml')).toBe(await pkg.readBytes('root.xml'));
	});
	it.each([
		pngSize(20000, 1),
		pngSize(4001, 4000),
		pngSize(0xffffffff, 0xffffffff),
		jpegSize(20000, 1),
	])('rejects dangerous dimensions', async (bytes) => {
		const { image, report } = await read({ bytes });
		expect(image).toBeUndefined();
		expect(report).toHaveBeenCalledWith('image-limit', expect.any(String), { part: source });
	});
	it('rejects excessive GIF screen dimensions', async () => {
		const bytes = gif.slice();
		new DataView(bytes.buffer).setUint16(6, 65535, true);
		const { image, report } = await read({ bytes });
		expect(image).toBeUndefined();
		expect(report.mock.calls[0]?.[0]).toBe('image-limit');
	});
	it('rejects a GIF frame outside its logical screen', async () => {
		const bytes = gif.slice();
		const at = bytes.indexOf(0x2c);
		new DataView(bytes.buffer).setUint16(at + 5, 2, true);
		const { image, report } = await read({ bytes });
		expect(image).toBeUndefined();
		expect(report.mock.calls[0]?.[0]).toBe('invalid-image');
	});
	it('bounds aggregate animation pixels', async () => {
		const start = gif.indexOf(0x2c),
			bytes = new Uint8Array(gif.length + gif.length - start - 1);
		bytes.set(gif.subarray(0, gif.length - 1));
		bytes.set(gif.subarray(start), gif.length - 1);
		const { image, report } = await read({ bytes }, { maxPixels: 1 });
		expect(image).toBeUndefined();
		expect(report.mock.calls[0]?.[0]).toBe('image-limit');
	});
	it('rejects images above the byte limit', async () => {
		const { image, report } = await read({}, { maxImageBytes: png.length - 1 });
		expect(image).toBeUndefined();
		expect(report.mock.calls[0]?.[0]).toBe('image-limit');
	});
	it.each([0, -1, NaN, Infinity, 0.5])(
		'rejects invalid configurable limits %s',
		async (maxPixels) => {
			await expect(read({}, { maxPixels })).rejects.toMatchObject({ code: 'INVALID_LIMIT' });
		},
	);
	it.each([
		'<svg xmlns="http://www.w3.org/2000/svg" onload="alert(1)"/>',
		'<!DOCTYPE html><html><script>alert(1)</script></html>',
		'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1',
	])('never accepts active/vector/binary content disguised as PNG', async (content) => {
		const { image, report } = await read({ bytes: new TextEncoder().encode(content) });
		expect(image).toBeUndefined();
		expect(report.mock.calls[0]?.[0]).toBe('unsupported-image');
	});
	it.each(['Object', 'MetaFile', 'EnhMetaFile'])(
		'does not interpret %s as bitmap data',
		async (foreignType) => {
			const { image, report } = await read({ foreignType });
			expect(image).toBeUndefined();
			expect(report.mock.calls[0]?.[0]).toBe('unsupported-image');
		},
	);
	it('does not load external image targets', async () => {
		const { pkg, foreignData, report } = await setup({
			target: 'https://example.invalid/image.png',
			mode: 'External',
		});
		const readBytes = vi.spyOn(pkg, 'readBytes');
		expect(await readVisioImage(pkg, source, foreignData, report)).toBeUndefined();
		expect(readBytes.mock.calls.every(([path]) => path.endsWith('.rels'))).toBe(true);
		expect(report.mock.calls[0]?.[0]).toBe('external-image');
	});
	it.each([
		{ type: 'http://attacker.invalid/image' },
		{ type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject' },
		{ reference: '<Rel r:id="missing"/>' },
		{ reference: '<Rel/>' },
		{ reference: '<Rel r:id="rId1"/><Rel r:id="rId1"/>' },
		{ reference: '<Rel xmlns:r="http://attacker.invalid/relationships" r:id="rId1"/>' },
	])('rejects unauthenticated image relationships %j', async (options) => {
		const { image, report } = await read(options);
		expect(image).toBeUndefined();
		expect(report.mock.calls[0]?.[0]).toBe('invalid-image-relationship');
	});
	it.each(['../../../outside.png', '../media/missing.png'])(
		'preserves package relationship safety rejection for %s',
		async (target) => {
			await expect(read({ target })).rejects.toMatchObject({ code: 'INVALID_RELATIONSHIP' });
		},
	);
	it('preserves binary CRC failure instead of disguising it as unsupported media', async () => {
		const { archive } = await setup();
		const at = Buffer.from(archive).indexOf(png);
		expect(at).toBeGreaterThan(0);
		archive[at + png.length - 1]! ^= 1;
		const pkg = await VisioPackage.open(archive),
			report = vi.fn<Report>();
		await expect(
			readVisioImage(pkg, source, child(await pkg.readXml(source), 'ForeignData')!, report),
		).rejects.toMatchObject({ code: 'ZIP_MISMATCH' });
		expect(report).not.toHaveBeenCalled();
	});
	it.each([
		png.subarray(0, 8),
		png.subarray(0, 24),
		png.subarray(0, 32),
		png.subarray(0, png.length - 1),
		gif.subarray(0, 6),
		gif.subarray(0, 12),
		gif.subarray(0, gif.length - 1),
		jpeg.subarray(0, 2),
		jpeg.subarray(0, 20),
		jpeg.subarray(0, jpeg.length - 1),
	])('reports malformed or truncated recognized image headers', async (bytes) => {
		const { image, report } = await read({ bytes });
		expect(image).toBeUndefined();
		expect(report.mock.calls[0]?.[0]).toBe('invalid-image');
	});
	it('rejects zero dimensions and PNG checksum mismatches', async () => {
		const corrupt = png.slice();
		corrupt[29]! ^= 1;
		for (const bytes of [pngSize(0, 1), corrupt]) {
			const { image, report } = await read({ bytes });
			expect(image).toBeUndefined();
			expect(report.mock.calls[0]?.[0]).toBe('invalid-image');
		}
	});
});
