import { emfInputOptions, emfInputView } from './emf-admission-input.js';
import { EmfAdmissionContext } from './emf-admission-context.js';
import { inspectEmfComment } from './emf-admission-comment.js';
import { inspectEmfGeometry } from './emf-admission-geometry.js';
import {
	EMF_FIXED_SIZES,
	inspectEmfObjectRecord,
	inspectEmfStateRecord,
} from './emf-admission-records.js';
import {
	type VisioEmfAdmissionOptions,
	type VisioEmfAdmissionResult,
} from './emf-admission-types.js';
export * from './emf-admission-types.js';

/** Bounded, DOM-free classic EMF inspection. Never invokes a converter or enables rendering.
 * Unknown or unresolved semantics reject the entire asset; no partial drawing is returned.
 * Caller owns package/document-wide limits and must recheck bytes at the conversion boundary.
 * MS-EMF records: https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-emf/67fd513f-e633-419c-93c2-8368779ac15a
 */
export function inspectVisioEmfAdmission(
	input: Uint8Array,
	options: VisioEmfAdmissionOptions = {},
): VisioEmfAdmissionResult {
	const snapshot = emfInputOptions(options);
	const view = snapshot === null ? null : emfInputView(input);
	const c = new EmfAdmissionContext(view ?? new DataView(new ArrayBuffer(0)), snapshot ?? {});
	if (snapshot === null) {
		c.emit('invalid', 'limits', 'Options must contain only plain own numeric data properties.');
		return c.result();
	}
	if (c.invalid) return c.result();
	if (view === null) {
		c.emit(
			'invalid',
			'input',
			'Expected a Uint8Array backed by an unshared, fixed, attached ArrayBuffer.',
		);
		return c.result();
	}
	if (
		!c.bound(
			view.byteLength,
			c.limits.maxInputBytes,
			'allocation-limit',
			'Encoded input exceeds byte budget.',
		)
	)
		return c.result();
	if (view.byteLength < 88 || view.byteLength % 4 !== 0) {
		c.emit('invalid', 'header', 'Input must contain an aligned complete EMF header.');
		return c.result();
	}
	const headerSize = c.u32(4);
	if (!inspectHeader(c, headerSize)) return c.result();
	let eof = false;
	while (c.offset < view.byteLength && !c.budgetExceeded) {
		if (++c.metrics.records > c.limits.maxRecords) {
			c.emit(
				'budget-exceeded',
				'record-limit',
				'Aggregate record budget exceeded.',
				c.limits.maxRecords,
			);
			break;
		}
		if (view.byteLength - c.offset < 8) {
			c.emit('invalid', 'record-size', 'Truncated record header.');
			break;
		}
		c.type = c.u32(0);
		const size = c.u32(4);
		if (size < 8 || size % 4 || size > view.byteLength - c.offset) {
			c.emit(
				'invalid',
				'record-size',
				'Record must progress, align, and fit entirely within input.',
			);
			break;
		}
		if (!c.inventory.has(c.type) && c.inventory.size >= c.limits.maxRecordTypes) {
			c.emit(
				'budget-exceeded',
				'record-limit',
				'Distinct record type budget exceeded.',
				c.limits.maxRecordTypes,
			);
			break;
		}
		c.inventory.set(c.type, (c.inventory.get(c.type) ?? 0) + 1);
		if (!c.charge(1)) break;
		const fixedSize = EMF_FIXED_SIZES.get(c.type);
		if (fixedSize !== undefined && size !== fixedSize) {
			c.emit('invalid', 'payload-size', 'Fixed record size does not match its defined payload.');
			c.semanticStateKnown = false;
		} else if (c.type === 1) {
			if (c.offset !== 0) c.emit('invalid', 'header', 'EMF header must occur exactly once.');
		} else if (c.type === 14) {
			inspectEof(c, size);
			eof = true;
			if (c.offset + size !== view.byteLength)
				c.emit('invalid', 'trailing-data', 'Data follows EOF.');
			else c.scanComplete = true;
			break;
		} else if (c.type === 70) inspectEmfComment(c, size, headerSize);
		else if (
			!inspectEmfStateRecord(c) &&
			!inspectEmfObjectRecord(c) &&
			!inspectEmfGeometry(c, size)
		) {
			c.emit(
				'unsupported',
				'unsupported-record',
				'Record has no admitted semantics; images, fonts, text and nested formats are excluded.',
			);
			// Unknown records may alter the object table or DC; do not fabricate downstream corruption.
			c.semanticStateKnown = false;
		}
		c.offset += size;
	}
	if (!c.budgetExceeded) {
		if (!eof) c.emit('invalid', 'missing-eof', 'A terminal EOF record is required.');
		if (c.metrics.records !== c.header?.declaredRecords)
			c.emit('invalid', 'record-count', 'Declared and framed record counts differ.');
		if (c.semanticStateKnown && c.stack.length)
			c.emit('unsupported', 'unbalanced-state', 'This subset requires a balanced SaveDC stack.');
	}
	return c.result();
}

function inspectHeader(c: EmfAdmissionContext, size: number): boolean {
	c.type = 1;
	if (
		c.u32(0) !== 1 ||
		c.u32(40) !== 0x464d4520 ||
		c.u32(44) !== 0x10000 ||
		size < 88 ||
		size % 4 ||
		size > c.view.byteLength ||
		c.u32(48) !== c.view.byteLength ||
		c.u32(52) < 2
	) {
		c.emit('invalid', 'header', 'Invalid EMF header type, signature, version, size, or counts.');
		return false;
	}
	const handles = c.view.getUint16(56, true);
	if (handles === 0 || c.view.getUint16(58, true) !== 0)
		c.emit('invalid', 'header', 'Reserved header fields or handle count are invalid.');
	c.bound(handles, c.limits.maxHandles, 'object-limit', 'Declared handle table exceeds budget.');
	c.bound(c.u32(52), c.limits.maxRecords, 'record-limit', 'Declared record count exceeds budget.');
	const count = c.u32(60),
		at = c.u32(64);
	if (count && at && (at < 88 || at % 2 || count * 2 > size - at)) {
		c.emit('invalid', 'header', 'Description exceeds its header record.');
		return false;
	}
	if (count || at || ![88, 100, 108].includes(size))
		c.emit('unsupported', 'header', 'Only fixed headers without descriptions are characterized.');
	if ((size === 100 || size === 108) && (c.u32(88) || c.u32(92) || c.u32(96)))
		c.emit('unsupported', 'header', 'Pixel format and OpenGL header extensions are excluded.');
	c.rect(8, false);
	c.rect(24, false);
	const bounds = [c.i32(8), c.i32(12), c.i32(16), c.i32(20)] as const;
	const width = bounds[2] - bounds[0] + 1,
		height = bounds[3] - bounds[1] + 1;
	if (width <= 0 || height <= 0)
		c.emit('invalid', 'header', 'Inclusive header bounds must be nonempty.');
	for (const offset of [72, 76, 80, 84, ...(size === 108 ? [100, 104] : [])]) {
		const value = c.i32(offset);
		if (value <= 0) c.emit('unsupported', 'header', 'Device dimensions must be positive.');
		c.coordinate(value);
	}
	c.headerPaletteEntries = c.u32(68);
	c.bound(
		c.headerPaletteEntries,
		c.limits.maxPaletteEntries,
		'palette-limit',
		'Declared palette exceeds budget.',
	);
	if (c.headerPaletteEntries)
		c.emit('unsupported', 'palette', 'Header/EOF palettes are outside this subset.');
	c.header = {
		bounds,
		pixelWidth: width,
		pixelHeight: height,
		declaredRecords: c.u32(52),
		declaredHandles: handles,
	};
	return !c.invalid && !c.budgetExceeded;
}

function inspectEof(c: EmfAdmissionContext, size: number): void {
	if (size < 20) {
		c.emit('invalid', 'eof', 'Truncated EOF record.');
		return;
	}
	const count = c.u32(8),
		offset = c.u32(12);
	c.metrics.paletteEntries += count;
	c.bound(
		c.metrics.paletteEntries,
		c.limits.maxPaletteEntries,
		'palette-limit',
		'Aggregate EOF palette exceeds budget.',
	);
	if (
		size !== 20 + count * 4 ||
		c.u32(size - 4) !== size ||
		(count && offset !== 16) ||
		count !== c.headerPaletteEntries
	)
		c.emit('invalid', 'eof', 'EOF size, palette count/offset, or terminal size is inconsistent.');
	if (count) c.emit('unsupported', 'palette', 'EOF palettes are outside this subset.');
}
