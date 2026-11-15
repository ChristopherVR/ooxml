/**
 * Write an sfnt (TrueType) font file from its tables: the table directory with checksums, each
 * table padded to four bytes, and the `head` table's checksum adjustment. The tables are written
 * in the order given (callers pass them sorted by tag). DOM-free.
 */

/** `n` rounded up to a multiple of four. */
export const pad4 = (n: number): number => (n + 3) & ~3;

function checksum(bytes: Uint8Array): number {
	let sum = 0;
	for (let i = 0; i < bytes.length; i += 4)
		sum =
			(sum +
				(((bytes[i] ?? 0) << 24) |
					((bytes[i + 1] ?? 0) << 16) |
					((bytes[i + 2] ?? 0) << 8) |
					(bytes[i + 3] ?? 0))) >>>
			0;
	return sum;
}

/** The font file holding `tables` (`[tag, bytes]`, sorted by tag). */
export function writeSfnt(tables: readonly (readonly [string, Uint8Array])[]): Uint8Array {
	const directory = 12 + tables.length * 16;
	let total = directory;
	for (const [, bytes] of tables) total += pad4(bytes.length);
	const out = new Uint8Array(total);
	const view = new DataView(out.buffer);
	view.setUint32(0, 0x00010000);
	view.setUint16(4, tables.length);
	const power = 2 ** Math.floor(Math.log2(tables.length));
	view.setUint16(6, power * 16);
	view.setUint16(8, Math.log2(power));
	view.setUint16(10, tables.length * 16 - power * 16);
	let position = directory;
	let headAt = -1;
	tables.forEach(([name, bytes], index) => {
		const entry = 12 + index * 16;
		for (let i = 0; i < 4; i++) out[entry + i] = name.charCodeAt(i);
		view.setUint32(entry + 4, checksum(bytes));
		view.setUint32(entry + 8, position);
		view.setUint32(entry + 12, bytes.length);
		out.set(bytes, position);
		if (name === 'head') headAt = position;
		position += pad4(bytes.length);
	});
	if (headAt >= 0) view.setUint32(headAt + 8, (0xb1b0afba - checksum(out)) >>> 0);
	return out;
}
