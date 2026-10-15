import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { readPackage } from '../docx/parse';
import { listEmbeddedPackages } from './embedded-packages';
import { deflateZeros, lyingZipOf } from './__fixtures__/lying-zip';
import { isSafePartName, openZip, ZipLimitError } from './zip-limits';

/** 64 MiB of zeros that the headers claim is 100 bytes. */
async function bombEntry(name: string) {
	return { name, data: await deflateZeros(64 << 20), declared: 100 };
}

async function zipOf(files: Record<string, string>): Promise<Uint8Array> {
	const zip = new JSZip();
	for (const [name, text] of Object.entries(files)) zip.file(name, text);
	return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

describe('isSafePartName', () => {
	it('accepts ordinary part names', () => {
		for (const name of ['word/document.xml', '/word/document.xml', '[Content_Types].xml'])
			expect(isSafePartName(name)).toBe(true);
	});
	it('rejects traversal, absolute, backslash, drive and NUL names', () => {
		for (const name of ['../x.xml', 'a/../../x', '//x', 'a\\b.xml', 'C:/x', 'c:x', 'a\0b'])
			expect(isSafePartName(name)).toBe(false);
	});
});

describe('openZip', () => {
	it('rejects traversal part names', async () => {
		const bytes = await zipOf({ '../evil.xml': '<a/>' });
		await expect(openZip(bytes)).rejects.toBeInstanceOf(ZipLimitError);
	});
	it('rejects forged-size entries while inflating, not after', async () => {
		const zip = await openZip(lyingZipOf([await bombEntry('a.xml')]));
		const started = Date.now();
		await expect(zip.file('a.xml')!.async('string')).rejects.toThrow(/inflates past/);
		expect(Date.now() - started).toBeLessThan(5000);
	});
	it('enforces the running total across entries', async () => {
		const bytes = await zipOf({ 'a.xml': 'x'.repeat(600), 'b.xml': 'y'.repeat(600) });
		await expect(openZip(bytes, { maxTotalBytes: 1000 })).rejects.toThrow(/content limit/);
	});
	it('rejects an oversized input', async () => {
		await expect(openZip(new Uint8Array(10), { maxInputBytes: 5 })).rejects.toThrow(/input limit/);
	});
});

describe('forged headers in each loader', () => {
	it('rejects a docx whose document part lies about its size', async () => {
		const bytes = lyingZipOf([await bombEntry('word/document.xml')]);
		await expect(readPackage(bytes)).rejects.toThrow(/inflates past/);
	});
	it('rejects a forged embedded package', async () => {
		const bytes = lyingZipOf([await bombEntry('word/embeddings/x.docx')]);
		await expect(listEmbeddedPackages(bytes)).rejects.toThrow(/inflates past/);
	});
});
