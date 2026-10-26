import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { readPackage } from '../docx/parse';
import { listEmbeddedPackages } from './embedded-packages';
import { deflateZeros, lyingZipOf } from './__fixtures__/lying-zip';
import { isSafePartName, openZip, streamEntryGuarded, ZipLimitError } from './zip-limits';

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

describe('limit kinds', () => {
	it('names the check each violation failed', async () => {
		const bytes = await zipOf({ 'a.xml': 'x'.repeat(600), 'b.xml': 'y'.repeat(600) });
		const fails = (input: Uint8Array, limits: Parameters<typeof openZip>[1], kind: string) =>
			expect(openZip(input, limits)).rejects.toMatchObject({ name: 'ZipLimitError', kind });
		await fails(new Uint8Array(10), { maxInputBytes: 5 }, 'maxInputBytes');
		await fails(bytes, { maxEntries: 1 }, 'maxEntries');
		await fails(bytes, { maxEntryBytes: 500 }, 'maxEntryBytes');
		await fails(bytes, { maxTotalBytes: 1000 }, 'maxTotalBytes');
		await fails(await zipOf({ '../evil.xml': '<a/>' }), {}, 'partName');
		const entry = { name: 'a.xml', data: await deflateZeros(1 << 16), declared: 100 };
		const forged = await openZip(lyingZipOf([entry]));
		await expect(forged.file('a.xml')!.async('string')).rejects.toMatchObject({
			kind: 'declaredSize',
		});
	});
	it('passes the kind to a custom error', async () => {
		const createError = (message: string, details: { limit: number; kind: string }) =>
			Object.assign(new Error(message), { details });
		const bytes = await zipOf({ 'a.xml': 'x'.repeat(600) });
		await expect(openZip(bytes, { maxEntryBytes: 500, createError })).rejects.toMatchObject({
			details: { limit: 500, kind: 'maxEntryBytes' },
		});
	});
	it('reports the observed value of each size check', async () => {
		const bytes = await zipOf({ 'a.xml': 'x'.repeat(600), 'b.xml': 'y'.repeat(600) });
		const fails = (input: Uint8Array, limits: Parameters<typeof openZip>[1], actual: number) =>
			expect(openZip(input, limits)).rejects.toMatchObject({ actual });
		await fails(new Uint8Array(10), { maxInputBytes: 5 }, 10);
		await fails(bytes, { maxEntries: 1 }, 2);
		await fails(bytes, { maxEntryBytes: 500 }, 600);
		await fails(bytes, { maxTotalBytes: 1000 }, 1200);
		await expect(openZip(await zipOf({ '../evil.xml': '<a/>' }))).rejects.toMatchObject({
			kind: 'partName',
			actual: undefined,
		});
		const entry = { name: 'a.xml', data: await deflateZeros(1 << 16), declared: 100 };
		const forged = await openZip(lyingZipOf([entry]));
		const error = await forged
			.file('a.xml')!
			.async('string')
			.catch((caught: unknown) => caught);
		// The size read so far when the entry passed the size its header declares.
		expect((error as ZipLimitError).actual).toBeGreaterThan(100);
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

/** Reads a stream to its end and returns how many bytes it gave. */
async function drain(chunks: AsyncIterable<Uint8Array>): Promise<number> {
	let size = 0;
	for await (const chunk of chunks) size += chunk.byteLength;
	return size;
}

describe('streamEntryGuarded', () => {
	/** An entry of about 5 MB of numbers, so it inflates in many chunks. */
	const text = Array.from({ length: 1 << 19 }, (_, i) => `${(i * 2654435761) % 1e9}\n`).join('');
	let opened: Promise<JSZip> | undefined;
	const bigZip = async () => ({
		text,
		zip: await (opened ??= zipOf({ 'a.xml': text }).then((bytes) => JSZip.loadAsync(bytes))),
	});

	it('hands out the inflated entry chunk by chunk', async () => {
		const { zip } = await bigZip();
		const chunks: Uint8Array[] = [];
		for await (const chunk of streamEntryGuarded(zip.file('a.xml')!, 0, {})) chunks.push(chunk);
		expect(chunks.length).toBeGreaterThan(1);
		expect(new TextDecoder().decode(Buffer.concat(chunks))).toBe(text);
	});

	it('yields the first chunk without inflating the whole entry', async () => {
		const { zip } = await bigZip();
		const entry = zip.file('a.xml')!;
		let read = 0;
		for await (const chunk of streamEntryGuarded(entry, 0, {})) {
			read += chunk.byteLength;
			break;
		}
		expect(read).toBeGreaterThan(0);
		expect(read).toBeLessThan(1 << 20);
	});

	it('keeps a rejected entry paused when a resume asked for earlier runs late', async () => {
		// An entry streamed the way JSZip streams: a pause is immediate, a resume runs on a later
		// turn, and a running stream schedules its next chunk itself.
		let paused = true;
		let scheduled = false;
		let sent = 0;
		const listeners: Record<string, ((chunk?: Uint8Array) => void)[]> = { data: [], end: [] };
		const tick = (): void => {
			scheduled = false;
			if (paused || sent === 1000) return;
			sent++;
			for (const listener of listeners.data ?? []) listener(new Uint8Array(1024));
			if (sent === 1000) for (const listener of listeners.end ?? []) listener();
			else {
				scheduled = true;
				setImmediate(tick);
			}
		};
		const stream = {
			on(event: string, listener: (chunk?: Uint8Array) => void) {
				(listeners[event] ??= []).push(listener);
				return stream;
			},
			pause() {
				paused = true;
				return stream;
			},
			resume() {
				setImmediate(() => {
					if (!paused) return;
					paused = false;
					if (!scheduled) setImmediate(tick);
					scheduled = true;
				});
				return stream;
			},
		};
		const entry = {
			name: 'x.xml',
			_data: { uncompressedSize: 5 * 1024 },
			internalStream: () => stream,
		} as unknown as JSZip.JSZipObject;
		await expect(drain(streamEntryGuarded(entry, 0, {}))).rejects.toMatchObject({
			kind: 'declaredSize',
		});
		await new Promise((resolve) => setTimeout(resolve, 100));
		expect(sent).toBeLessThan(10);
	});

	it('applies the limits while streaming, counting what was read before', async () => {
		const { zip } = await bigZip();
		const read = (entry: JSZip.JSZipObject, alreadyRead: number, limits = {}) =>
			drain(streamEntryGuarded(entry, alreadyRead, limits));
		const entry = zip.file('a.xml')!;
		await expect(read(entry, 0, { maxEntryBytes: 1 << 20 })).rejects.toMatchObject({
			kind: 'maxEntryBytes',
		});
		await expect(read(entry, 100, { maxTotalBytes: text.length + 99 })).rejects.toMatchObject({
			kind: 'maxTotalBytes',
		});
		await expect(read(entry, 100, { maxTotalBytes: text.length + 100 })).resolves.toBe(text.length);
		const forged = await openZip(lyingZipOf([await bombEntry('b.xml')]));
		await expect(read(forged.file('b.xml')!, 0)).rejects.toMatchObject({ kind: 'declaredSize' });
	});
});
