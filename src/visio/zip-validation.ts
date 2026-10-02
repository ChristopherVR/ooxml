import {
	fail,
	utf8,
	safePath,
	decodePath,
	type Entry,
	type VisioPackageLimits,
} from './package-common.js';

/** Inspect central AND local headers before JSZip can sanitize names or inflate anything. */
export function inspectZip(
	bytes: Uint8Array,
	limits: VisioPackageLimits,
	check: () => void,
): Map<string, Entry> {
	const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
	const bounds = (start: number, length: number) => {
		if (start < 0 || length < 0 || start + length > bytes.length)
			fail('INVALID_ZIP', 'Truncated ZIP record');
	};
	const u16 = (offset: number) => {
		bounds(offset, 2);
		return view.getUint16(offset, true);
	};
	const u32 = (offset: number) => {
		bounds(offset, 4);
		return view.getUint32(offset, true);
	};
	let end = -1;
	for (let i = bytes.length - 22; i >= Math.max(0, bytes.length - 65557); i--) {
		if (u32(i) === 0x06054b50 && i + 22 + u16(i + 20) === bytes.length) {
			end = i;
			break;
		}
	}
	if (end < 0) fail('INVALID_ZIP', 'Missing ZIP end record');
	const count = u16(end + 10),
		centralSize = u32(end + 12),
		centralStart = u32(end + 16);
	if (
		u16(end + 4) ||
		u16(end + 6) ||
		u16(end + 8) !== count ||
		count === 0xffff ||
		centralSize === 0xffffffff ||
		centralStart === 0xffffffff
	)
		fail('UNSUPPORTED_ZIP', 'ZIP64 and multidisk packages are unsupported');
	if (count > limits.maxEntries) fail('LIMIT_ENTRIES', 'Too many ZIP entries');
	if (centralStart + centralSize !== end) fail('INVALID_ZIP', 'Invalid central directory bounds');
	const entries = new Map<string, Entry>(),
		canonical = new Set<string>();
	const spans: { start: number; end: number }[] = [];
	let offset = centralStart,
		total = 0;
	const extras = (start: number, length: number, name: string) => {
		const stop = start + length;
		while (start < stop) {
			if (start + 4 > stop) fail('INVALID_ZIP', 'Malformed ZIP extra field');
			const id = u16(start),
				size = u16(start + 2);
			start += 4;
			if (start + size > stop) fail('INVALID_ZIP', 'Truncated ZIP extra field');
			if (id === 1 || id === 0x9901)
				fail('UNSUPPORTED_ZIP', 'ZIP64 and encrypted entries are unsupported');
			if (
				id === 0x7075 &&
				(size < 5 ||
					bytes[start] !== 1 ||
					utf8.decode(bytes.subarray(start + 5, start + size)) !== name)
			)
				fail('ZIP_MISMATCH', 'Unicode path override differs from ZIP name');
			start += size;
		}
	};
	for (let i = 0; i < count; i++) {
		check();
		bounds(offset, 46);
		if (u32(offset) !== 0x02014b50) fail('INVALID_ZIP', 'Invalid central directory entry');
		const flags = u16(offset + 8),
			method = u16(offset + 10),
			crc = u32(offset + 16);
		const compressed = u32(offset + 20),
			size = u32(offset + 24);
		const nameLength = u16(offset + 28),
			extraLength = u16(offset + 30),
			commentLength = u16(offset + 32);
		const local = u32(offset + 42);
		if (flags & ~0x080e || (method !== 0 && method !== 8) || u16(offset + 34))
			fail('UNSUPPORTED_ZIP', 'Encrypted, multidisk, or unsupported ZIP entry');
		if (compressed === 0xffffffff || size === 0xffffffff || local === 0xffffffff)
			fail('UNSUPPORTED_ZIP', 'ZIP64 entries are unsupported');
		bounds(offset + 46, nameLength + extraLength + commentLength);
		const name = utf8.decode(bytes.subarray(offset + 46, offset + 46 + nameLength));
		const directory = name.endsWith('/');
		safePath(name, directory);
		const key = decodePath(name).replace(/\/$/, '');
		if (entries.has(name) || canonical.has(key))
			fail('DUPLICATE_ENTRY', `Duplicate ZIP path: ${name}`);
		canonical.add(key);
		extras(offset + 46 + nameLength, extraLength, name);
		if (size > limits.maxEntryBytes) fail('LIMIT_ENTRY', `ZIP entry is too large: ${name}`);
		total += size;
		if (total > limits.maxTotalBytes) fail('LIMIT_TOTAL', 'Inflated ZIP total exceeds limit');
		if (size > Math.max(1, compressed) * limits.maxCompressionRatio)
			fail('LIMIT_RATIO', `ZIP compression ratio exceeds limit: ${name}`);
		if ((method === 0 && compressed !== size) || (directory && (compressed || size)))
			fail('ZIP_MISMATCH', 'Invalid stored entry size');
		bounds(local, 30);
		if (
			u32(local) !== 0x04034b50 ||
			u16(local + 6) !== flags ||
			u16(local + 8) !== method ||
			u16(local + 26) !== nameLength
		)
			fail('ZIP_MISMATCH', 'Local and central ZIP headers differ');
		const localExtra = u16(local + 28),
			dataStart = local + 30 + nameLength + localExtra;
		bounds(local + 30, nameLength + localExtra);
		if (
			!bytes
				.subarray(local + 30, local + 30 + nameLength)
				.every((byte, index) => byte === bytes[offset + 46 + index])
		)
			fail('ZIP_MISMATCH', 'Local and central ZIP filenames differ');
		extras(local + 30 + nameLength, localExtra, name);
		for (const [relative, expected] of [
			[14, crc],
			[18, compressed],
			[22, size],
		] as const) {
			const actual = u32(local + relative);
			if (actual !== expected && (!(flags & 8) || actual !== 0))
				fail('ZIP_MISMATCH', 'Local and central ZIP sizes or CRC differ');
		}
		let dataEnd = dataStart + compressed;
		bounds(dataStart, compressed);
		if (flags & 8) {
			if (u32(dataEnd) === 0x08074b50) dataEnd += 4;
			if (u32(dataEnd) !== crc || u32(dataEnd + 4) !== compressed || u32(dataEnd + 8) !== size)
				fail('ZIP_MISMATCH', 'Invalid ZIP data descriptor');
			dataEnd += 12;
		}
		if (dataEnd > centralStart) fail('ZIP_MISMATCH', 'ZIP entry overlaps the central directory');
		spans.push({ start: local, end: dataEnd });
		entries.set(name, { name, compressed, size, crc, directory });
		offset += 46 + nameLength + extraLength + commentLength;
	}
	if (offset !== end) fail('INVALID_ZIP', 'Central directory length or count mismatch');
	spans.sort((a, b) => a.start - b.start);
	let previousEnd = 0;
	for (const span of spans) {
		if (span.start !== previousEnd)
			fail('ZIP_MISMATCH', 'Overlapping, prefixed, or unindexed ZIP records');
		previousEnd = span.end;
	}
	if (previousEnd !== centralStart) fail('ZIP_MISMATCH', 'Unindexed ZIP data');
	return entries;
}
