// @vitest-environment jsdom
import { expect, it, vi } from 'vitest';
import { defaultDrawingWrap, saveDrawingWrap } from './drawing-layout-default';

it('uses the session preference after a storage write fails and clears it after recovery', () => {
	const frame = document.createElement('iframe');
	document.body.append(frame);
	const doc = frame.contentDocument!;
	try {
		const storage = doc.defaultView!.localStorage;
		const writes = vi
			.spyOn(Object.getPrototypeOf(storage), 'setItem')
			.mockImplementationOnce(() => {
				throw new DOMException('Storage limit reached', 'QuotaExceededError');
			});
		try {
			saveDrawingWrap(doc, 'shape', 'front');
			expect(defaultDrawingWrap(doc, 'shape')).toBe('front');
			expect(defaultDrawingWrap(doc, 'picture')).toBeUndefined();
			saveDrawingWrap(doc, 'shape', 'square');
			expect(defaultDrawingWrap(doc, 'shape')).toBe('square');
		} finally {
			writes.mockRestore();
			storage.removeItem('ooxml-word-drawing-wrap.shape');
		}
	} finally {
		frame.remove();
	}
});
