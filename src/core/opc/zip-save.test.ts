import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { openZip } from './zip-limits';
import { keepCompressedMedia, writeZipBytes, writeZipText } from './zip-save';

/** The first entry's method and compressed payload, as advertised by its local header. */
function firstPayload(bytes: Uint8Array) {
	const header = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
	const start = 30 + header.getUint16(26, true) + header.getUint16(28, true);
	return {
		method: header.getUint16(8, true),
		bytes: bytes.subarray(start, start + header.getUint32(18, true)),
	};
}

const generate = (zip: JSZip, level = 6) =>
	zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', compressionOptions: { level } });

describe('save ZIP parts', () => {
	it.each(['STORE', 'DEFLATE'] as const)(
		'retains the source %s payload of media',
		async (compression) => {
			const path = 'sample.mp4';
			const media = new Uint8Array(1024).map((_, index) => index % 35);
			const source = await generate(new JSZip().file(path, media, { compression }), 1);
			const zip = await openZip(source);
			const entry = zip.file(path);
			await writeZipBytes(zip, path, media);
			expect(zip.file(path)).toBe(entry);
			keepCompressedMedia(zip);
			expect(firstPayload(await generate(zip, 9))).toEqual(firstPayload(source));
		},
	);

	it('stores new compressed media and deflates new XML', async () => {
		const zip = new JSZip().file('sample.woff2', 'font bytes');
		keepCompressedMedia(zip);
		expect(firstPayload(await generate(zip)).method).toBe(0);
		const xml = new JSZip().file('sample.xml', '<root/>');
		keepCompressedMedia(xml);
		expect(firstPayload(await generate(xml)).method).toBe(8);
	});

	it('updates the ZIP checksum when the caller edits a stored view in place', async () => {
		const path = 'thumbnail.png';
		const source = await generate(
			new JSZip().file(path, new Uint8Array([1, 2, 3]), { compression: 'STORE' }),
		);
		const zip = await openZip(source);
		const bytes = await zip.file(path)?.async('uint8array');
		if (!bytes) throw new Error('Missing thumbnail');
		bytes[0] = 4;
		await writeZipBytes(zip, path, bytes);
		keepCompressedMedia(zip);
		const saved = await JSZip.loadAsync(await generate(zip), { checkCRC32: true });
		expect(await saved.file(path)?.async('uint8array')).toEqual(new Uint8Array([4, 2, 3]));
	});

	it('keeps an equal XML entry and replaces changed or lossily decoded parts', async () => {
		const xml = '<root>same</root>';
		const source = await generate(new JSZip().file('sample.xml', xml), 1);
		const zip = await openZip(source);
		writeZipText(zip, 'sample.xml', xml, await zip.file('sample.xml')?.async('string'));
		expect(firstPayload(await generate(zip, 9))).toEqual(firstPayload(source));
		writeZipText(zip, 'sample.xml', '<root>updated</root>', xml);
		expect(await zip.file('sample.xml')?.async('string')).toBe('<root>updated</root>');
		zip.file('malformed.xml', new Uint8Array([0xff]));
		const text = await zip.file('malformed.xml')?.async('string');
		writeZipText(zip, 'malformed.xml', '\uFFFD', text);
		expect(await zip.file('malformed.xml')?.async('uint8array')).toEqual(
			new Uint8Array([0xef, 0xbf, 0xbd]),
		);
	});
});
