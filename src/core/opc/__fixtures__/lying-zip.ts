import { Readable } from 'node:stream';
import { createDeflateRaw } from 'node:zlib';

/** Raw-deflates `size` zero bytes without ever holding them in memory. */
export async function deflateZeros(size: number): Promise<Buffer> {
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

/** One entry of a hand-built zip: raw deflate `data` whose headers declare `declared` bytes. */
export interface LyingEntry {
	name: string;
	data: Buffer;
	declared: number;
}

/** A zip whose local and central headers declare `declared` uncompressed bytes per entry. */
export function lyingZipOf(entries: readonly LyingEntry[]): Uint8Array {
	const bodies: Buffer[] = [];
	const directory: Buffer[] = [];
	let offset = 0;
	for (const { name, data, declared } of entries) {
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
		central.writeUInt32LE(offset, 42);
		const body = Buffer.concat([local, file, data]);
		bodies.push(body);
		directory.push(central, file);
		offset += body.length;
	}
	const dir = Buffer.concat(directory);
	const end = Buffer.alloc(22);
	end.writeUInt32LE(0x06054b50, 0);
	end.writeUInt16LE(entries.length, 8);
	end.writeUInt16LE(entries.length, 10);
	end.writeUInt32LE(dir.length, 12);
	end.writeUInt32LE(offset, 16);
	return new Uint8Array(Buffer.concat([...bodies, dir, end]));
}
