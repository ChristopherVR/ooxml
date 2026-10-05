// Audit prototype only. Production metafile parsing belongs in the core project.
export const LIMITS = Object.freeze({
	inputBytes: 4 * 1024 * 1024,
	records: 20_000,
	objects: 4096,
	depth: 64,
	coordinate: 1_000_000,
	outputDimension: 2048,
	outputNodes: 25_000,
	outputChars: 2 * 1024 * 1024,
	wallMs: 3000,
});

// Narrow initial contract: no fonts, comments, images, paths, transforms or mapping.
const FIXED_SIZE = new Map([
	[14, 20],
	[17, 12],
	[19, 12],
	[27, 16],
	[33, 8],
	[34, 12],
	[37, 12],
	[38, 28],
	[39, 24],
	[40, 12],
	[42, 24],
	[43, 24],
	[54, 16],
]);
const NAMES = new Map([
	[1, 'HEADER'],
	[9, 'SETWINDOWEXTEX'],
	[11, 'SETVIEWPORTEXTEX'],
	[14, 'EOF'],
	[17, 'SETMAPMODE'],
	[70, 'COMMENT'],
	[76, 'BITBLT'],
	[82, 'EXTCREATEFONTINDIRECTW'],
	[84, 'EXTTEXTOUTW'],
	[0x4034, 'SetClipRegion'],
	[0x4036, 'DrawDriverString'],
]);

export function preflight(input, limits = LIMITS) {
	const diagnostics = [];
	const recordTypes = {};
	const plusTypes = {};
	let omittedDiagnostics = 0;
	let records = 0;
	let plusRecords = 0;
	let structurallyValid = true;
	let eligible = true;
	let dimensions = null;
	const emit = (code, offset, type, detail, structural = false) => {
		eligible = false;
		if (structural) structurallyValid = false;
		if (diagnostics.length < 128) diagnostics.push({ code, offset, type, detail });
		else omittedDiagnostics++;
	};
	const result = () => ({
		eligible,
		structurallyValid,
		records,
		plusRecords,
		dimensions,
		recordTypes,
		plusTypes,
		diagnostics,
		omittedDiagnostics,
	});
	if (!(input instanceof Uint8Array) || input.byteLength > limits.inputBytes) {
		emit('emf.input-limit', 0, null, 'Expected bounded Uint8Array input.', true);
		return result();
	}
	if (input.byteLength < 88) {
		emit('emf.header', 0, null, 'EMF header is truncated; WMF is not in this subset.', true);
		return result();
	}
	const view = new DataView(input.buffer, input.byteOffset, input.byteLength);
	const u32 = (o) => view.getUint32(o, true);
	const i32 = (o) => view.getInt32(o, true);
	const headerSize = u32(4);
	if (
		u32(0) !== 1 ||
		u32(40) !== 0x464d4520 ||
		u32(44) !== 0x10000 ||
		headerSize < 88 ||
		headerSize % 4 ||
		headerSize > input.byteLength
	) {
		emit('emf.header', 0, 1, 'Invalid type, signature, version or header size.', true);
		return result();
	}
	if (u32(48) !== input.byteLength || u32(52) < 2 || u32(52) > limits.records) {
		emit('emf.declared-size', 48, 1, 'Declared byte/record count is invalid or over budget.', true);
		return result();
	}
	const descriptionLength = u32(60) * 2;
	if (descriptionLength && (u32(64) < 88 || u32(64) + descriptionLength > headerSize)) {
		emit('emf.description-range', 60, 1, 'Description escapes its header record.', true);
	}
	if (view.getUint16(56, true) > limits.objects)
		emit('emf.object-limit', 56, 1, 'Too many declared handles.');
	const width = i32(16) - i32(8);
	const height = i32(20) - i32(12);
	dimensions = { width, height };
	if (
		width <= 0 ||
		height <= 0 ||
		[8, 12, 16, 20].some((o) => Math.abs(i32(o)) > limits.coordinate)
	) {
		emit('emf.bounds', 8, 1, 'Bounds must be positive and bounded.', true);
	}
	const objects = new Set();
	let depth = 0;
	let offset = 0;
	let sawEof = false;
	while (offset < input.byteLength) {
		if (++records + plusRecords > limits.records) {
			emit('emf.record-limit', offset, null, 'Aggregate EMF + EMF+ record budget exceeded.', true);
			break;
		}
		if (offset + 8 > input.byteLength) {
			emit('emf.record-header', offset, null, 'Truncated record header.', true);
			break;
		}
		const type = u32(offset);
		const size = u32(offset + 4);
		if (size < 8 || size % 4 || size > input.byteLength - offset) {
			emit('emf.record-size', offset, type, 'Record must progress, align and fit input.', true);
			break;
		}
		recordTypes[type] = (recordTypes[type] ?? 0) + 1;
		if (type === 1 && offset !== 0)
			emit('emf.repeated-header', offset, type, 'Header may occur only at offset zero.', true);
		if (type !== 1 && !FIXED_SIZE.has(type))
			emit('emf.unsupported-record', offset, type, NAMES.get(type) ?? `EMR_${type}`);
		if (FIXED_SIZE.has(type) && size !== FIXED_SIZE.get(type)) {
			emit(
				'emf.record-payload',
				offset,
				type,
				'Record does not match the subset payload size.',
				true,
			);
			offset += size;
			continue;
		}
		if (type === 70) scanComment(offset, size);
		if (type === 17 && u32(offset + 8) !== 1)
			emit('emf.map-mode', offset, type, 'Only MM_TEXT is in the initial subset.');
		if (type === 19 && ![1, 2].includes(u32(offset + 8)))
			emit('emf.fill-mode', offset, type, 'Invalid fill rule.');
		if ([27, 42, 43, 54].includes(type)) {
			for (let n = offset + 8; n < offset + size; n += 4) {
				if (Math.abs(i32(n)) > limits.coordinate)
					emit('emf.coordinate-limit', n, type, 'Geometry exceeds coordinate budget.');
			}
		}
		if (type === 33 && ++depth > limits.depth)
			emit('emf.state-depth', offset, type, 'Too many nested SaveDC records.');
		if (type === 34) {
			if (i32(offset + 8) !== -1 || depth === 0)
				emit(
					'emf.restore-state',
					offset,
					type,
					'Only balanced relative RestoreDC(-1) is admitted.',
				);
			else depth--;
		}
		if ([38, 39].includes(type)) {
			const handle = u32(offset + 8);
			if (!handle || handle >= limits.objects || objects.has(handle))
				emit('emf.object-handle', offset, type, 'Invalid or duplicate handle.');
			objects.add(handle);
			if (
				type === 38 &&
				(![0, 5].includes(u32(offset + 12)) ||
					i32(offset + 16) < 0 ||
					i32(offset + 16) > 1024 ||
					i32(offset + 20) !== 0)
			)
				emit('emf.pen-style', offset, type, 'Only bounded solid/null pens are admitted.');
			if (type === 39 && (![0, 1].includes(u32(offset + 12)) || u32(offset + 20) !== 0))
				emit('emf.brush-style', offset, type, 'Only solid/null brushes are admitted.');
			const color = u32(offset + (type === 38 ? 24 : 16));
			if (color >>> 24)
				emit('emf.palette-color', offset, type, 'Palette-relative colors are not admitted.');
		}
		if (type === 37) {
			const handle = u32(offset + 8);
			if (!objects.has(handle) && !(handle >= 0x80000000 && handle <= 0x80000008))
				emit('emf.object-reference', offset, type, 'Unknown or unsupported stock handle.');
		}
		if (type === 40 && !objects.delete(u32(offset + 8)))
			emit('emf.object-reference', offset, type, 'Delete references an unknown handle.');
		if (type === 14) {
			sawEof = true;
			if (u32(offset + 8) !== 0 || u32(offset + 16) !== 20)
				emit('emf.eof-payload', offset, type, 'Only a palette-free EOF is admitted.');
			if (offset + size !== input.byteLength)
				emit('emf.trailing-data', offset + size, type, 'Data follows EOF.', true);
			break;
		}
		offset += size;
	}
	if (!sawEof) emit('emf.missing-eof', offset, null, 'No valid EOF record.', true);
	if (records !== u32(52))
		emit('emf.record-count', 52, 1, 'Declared record count differs from framed records.', true);
	if (depth) emit('emf.unbalanced-state', offset, null, 'SaveDC stack is not balanced.');
	return result();

	function scanComment(start, size) {
		if (size < 12 || u32(start + 8) > size - 12) {
			emit('emf.comment-range', start, 70, 'Comment payload escapes its record.', true);
			return;
		}
		const end = start + 12 + u32(start + 8);
		if (end - start < 16 || u32(start + 12) !== 0x2b464d45) return;
		let p = start + 16;
		while (p < end) {
			if (++plusRecords + records > limits.records || end - p < 12) {
				emit('emf.plus-record-limit', p, 70, 'Truncated or over-budget EMF+ stream.', true);
				return;
			}
			const kind = view.getUint16(p, true);
			const n = u32(p + 4);
			const dataSize = u32(p + 8);
			if (n < 12 || n % 4 || n > end - p || dataSize > n - 12 || n - 12 - dataSize > 3) {
				emit(
					'emf.plus-record-size',
					p,
					kind,
					'EMF+ data/record size exceeds enclosing boundaries.',
					true,
				);
				return;
			}
			plusTypes[`0x${kind.toString(16)}`] = (plusTypes[`0x${kind.toString(16)}`] ?? 0) + 1;
			if ([0x4034, 0x4036].includes(kind))
				emit('emf.known-fidelity-risk', p, kind, NAMES.get(kind));
			p += n;
		}
	}
}
