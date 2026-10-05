import { Readable } from 'node:stream';
import { createDeflateRaw } from 'node:zlib';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { readZipParts } from './package.js';

/** Raw-deflates `size` zero bytes without ever holding them in memory. */
async function deflateZeros(size: number): Promise<Buffer> {
	const chunk = Buffer.alloc(1 << 20);
	const source = Readable.from(
		(function* () {
			for (let done = 0; done < size; done += chunk.length) yield chunk;
		})(),
	);
	const out: Buffer[] = [];
	for await (const piece of source.pipe(createDeflateRaw({ level: 9 }))) out.push(piece as Buffer);
	return Buffer.concat(out);
}

/** A one-entry zip whose headers declare `declared` uncompressed bytes for `data`. */
function lyingZip(name: string, data: Buffer, declared: number): Uint8Array {
	const file = Buffer.from(name, 'utf8');
	const local = Buffer.alloc(30);
	local.writeUInt32LE(0x04034b50, 0);
	local.writeUInt16LE(20, 4);
	local.writeUInt16LE(8, 8);
	local.writeUInt32LE(data.length, 18);
	local.writeUInt32LE(declared, 22);
	local.writeUInt16LE(file.length, 26);
	const central = Buffer.alloc(46);
	central.writeUInt32LE(0x02014b50, 0);
	central.writeUInt16LE(20, 4);
	central.writeUInt16LE(20, 6);
	central.writeUInt16LE(8, 10);
	central.writeUInt32LE(data.length, 20);
	central.writeUInt32LE(declared, 24);
	central.writeUInt16LE(file.length, 28);
	const body = Buffer.concat([local, file, data]);
	const directory = Buffer.concat([central, file]);
	const end = Buffer.alloc(22);
	end.writeUInt32LE(0x06054b50, 0);
	end.writeUInt16LE(1, 8);
	end.writeUInt16LE(1, 10);
	end.writeUInt32LE(directory.length, 12);
	end.writeUInt32LE(body.length, 16);
	return new Uint8Array(Buffer.concat([body, directory, end]));
}

describe('readZipParts limits', () => {
	it('reads an ordinary package', async () => {
		const zip = new JSZip();
		zip.file('a.xml', '<a/>');
		zip.file('dir/b.xml', 'x'.repeat(100_000));
		const parts = await readZipParts(
			await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }),
		);
		expect(parts.get('a.xml')?.byteLength).toBe(4);
		expect(parts.get('dir/b.xml')?.byteLength).toBe(100_000);
	});

	it('rejects a declared-size precheck failure without inflating', async () => {
		const zip = new JSZip();
		zip.file('big.xml', 'x'.repeat(10_000));
		const bytes = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
		await expect(readZipParts(bytes, { maxPartBytes: 1000 })).rejects.toThrow(/part size limit/);
		await expect(readZipParts(bytes, { maxTotalBytes: 1000 })).rejects.toThrow(
			/uncompressed content limit/,
		);
	});

	it('aborts a 1 GiB bomb that declares 100 bytes before inflating it', async () => {
		const bomb = lyingZip('xl/worksheets/sheet1.xml', await deflateZeros(1 << 30), 100);
		const before = process.memoryUsage();
		let peak = before.rss;
		const timer = setInterval(() => {
			peak = Math.max(peak, process.memoryUsage().rss);
		}, 2);
		const started = Date.now();
		try {
			await expect(readZipParts(bomb)).rejects.toThrow(/inflates past its declared size/);
		} finally {
			clearInterval(timer);
		}
		peak = Math.max(peak, process.memoryUsage().rss);
		expect(peak - before.rss).toBeLessThan(64 * 1024 * 1024);
		expect(Date.now() - started).toBeLessThan(5000);
	}, 120_000);
});
