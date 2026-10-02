import { type EmfAdmissionContext as Context } from './emf-admission-context.js';

/** Original-equivalent WMF comments only: GdiComment documents checksum and placement.
 * https://learn.microsoft.com/en-us/windows/win32/api/wingdi/nf-wingdi-gdicomment
 * No WMF bytes are copied, decoded, replayed, or traversed as nested records.
 */
export function inspectEmfComment(c: Context, size: number, headerSize: number): void {
	if (size < 12) {
		c.emit('invalid', 'comment-range', 'Truncated comment length.');
		return;
	}
	const bytes = c.u32(8);
	if (size !== 12 + Math.ceil(bytes / 4) * 4) {
		c.emit('invalid', 'comment-range', 'Comment length must exactly fit its aligned record.');
		return;
	}
	c.metrics.commentBytes += bytes;
	if (
		!c.bound(
			c.metrics.commentBytes,
			c.limits.maxCommentBytes,
			'comment-limit',
			'Aggregate comment payload exceeds budget.',
		)
	)
		return;
	if (bytes >= 4 && c.u32(12) === 0x2b464d45) {
		c.emit('unsupported', 'emf-plus', 'All EMF+ payloads are rejected without decoding.');
		return;
	}
	if (bytes < 8 || c.u32(12) !== 0x43494447 || c.u32(16) !== 0x80000001) {
		c.emit(
			'unsupported',
			'comment-subtype',
			'Only a validated original-equivalent WMF comment is characterized.',
		);
		return;
	}
	if (bytes < 24 || c.u32(32) !== bytes - 24) {
		c.emit(
			'invalid',
			'comment-range',
			'Equivalent-source WMF length escapes or underfills its comment.',
		);
		return;
	}
	if (c.offset !== headerSize || c.u32(20) !== 0x00000300 || c.u32(28) !== 0 || c.u32(32) < 18) {
		c.emit(
			'unsupported',
			'embedded-wmf',
			'WMF metadata must be immediately after the header with version 3, reserved/flags zero.',
		);
		return;
	}
	// The containing input is already byte-bounded and aligned. At most one comment can qualify.
	if (!c.charge(c.view.byteLength / 4)) return;
	let checksum = 0;
	for (let p = 0; p < c.view.byteLength; p += 4)
		checksum = (checksum + c.view.getUint32(p, true)) >>> 0;
	if (checksum !== 0)
		c.emit('unsupported', 'embedded-wmf', 'Original-equivalent WMF checksum is stale or invalid.');
}
