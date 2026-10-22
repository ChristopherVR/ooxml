// A minimal ZIP writer for the packages a peer rebuilds from the room. Every entry is stored
// (not compressed), so building the package is copying bytes; the CRC-32 of a part is cached by
// the identity of its bytes, which the compose step keeps stable for parts that did not change
// (decoded shared parts, serialised page parts), so an unchanged part is neither re-checksummed
// nor re-deflated on the next rebuild.
import { crc32 } from '../package-common';

const encoder = new TextEncoder();
/** 1980-01-01 00:00:00: the same timestamp for every entry, so equal packages are equal bytes. */
const DOS_TIME = 0;
const DOS_DATE = 0x21;

export type CrcCache = WeakMap<Uint8Array, number>;

/** A ZIP of `files` in the given order. `crcs` remembers the checksum of each byte array. */
export function writeStoredZip(
	files: readonly (readonly [string, Uint8Array])[],
	crcs: CrcCache,
): Uint8Array {
	const prepared = files.map(([name, data]) => {
		let crc = crcs.get(data);
		if (crc === undefined) {
			crc = (crc32(data) ^ 0xffffffff) >>> 0;
			crcs.set(data, crc);
		}
		return { name: encoder.encode(name), data, crc };
	});
	let local = 0;
	let central = 0;
	for (const f of prepared) {
		local += 30 + f.name.length + f.data.length;
		central += 46 + f.name.length;
	}
	const out = new Uint8Array(local + central + 22);
	const view = new DataView(out.buffer);
	let at = 0;
	const offsets: number[] = [];
	for (const f of prepared) {
		offsets.push(at);
		view.setUint32(at, 0x04034b50, true);
		view.setUint16(at + 4, 20, true);
		view.setUint16(at + 6, 0x0800, true); // UTF-8 names
		view.setUint16(at + 8, 0, true); // stored
		view.setUint16(at + 10, DOS_TIME, true);
		view.setUint16(at + 12, DOS_DATE, true);
		view.setUint32(at + 14, f.crc, true);
		view.setUint32(at + 18, f.data.length, true);
		view.setUint32(at + 22, f.data.length, true);
		view.setUint16(at + 26, f.name.length, true);
		view.setUint16(at + 28, 0, true);
		out.set(f.name, at + 30);
		out.set(f.data, at + 30 + f.name.length);
		at += 30 + f.name.length + f.data.length;
	}
	const centralStart = at;
	prepared.forEach((f, i) => {
		view.setUint32(at, 0x02014b50, true);
		view.setUint16(at + 4, 20, true);
		view.setUint16(at + 6, 20, true);
		view.setUint16(at + 8, 0x0800, true);
		view.setUint16(at + 10, 0, true);
		view.setUint16(at + 12, DOS_TIME, true);
		view.setUint16(at + 14, DOS_DATE, true);
		view.setUint32(at + 16, f.crc, true);
		view.setUint32(at + 20, f.data.length, true);
		view.setUint32(at + 24, f.data.length, true);
		view.setUint16(at + 28, f.name.length, true);
		view.setUint32(at + 42, offsets[i]!, true);
		out.set(f.name, at + 46);
		at += 46 + f.name.length;
	});
	view.setUint32(at, 0x06054b50, true);
	view.setUint16(at + 8, prepared.length, true);
	view.setUint16(at + 10, prepared.length, true);
	view.setUint32(at + 12, at - centralStart, true);
	view.setUint32(at + 16, centralStart, true);
	return out;
}
