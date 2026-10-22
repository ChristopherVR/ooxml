// @vitest-environment jsdom
import type { DocumentStorage } from 'ooxml-core/embedding';
import { afterEach, describe, expect, it, vi } from 'vitest';

import { deliverSavedPresentation, presentationSaveMimeType } from './presentation-save';

afterEach(() => vi.restoreAllMocks());

describe('deliverSavedPresentation', () => {
	it('hands the bytes to storage.save instead of downloading', async () => {
		const click = vi.spyOn(HTMLAnchorElement.prototype, 'click');
		const save = vi.fn(async () => undefined);
		const storage: DocumentStorage = { save };
		const bytes = new Uint8Array([1, 2]);
		expect(
			await deliverSavedPresentation({ storage, bytes, format: 'pptx', fileName: 'deck.pptx' }),
		).toBe('deck.pptx');
		expect(save).toHaveBeenCalledWith(bytes, {
			fileName: 'deck.pptx',
			mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
			reason: 'save',
		});
		expect(click).not.toHaveBeenCalled();
	});

	it('reports other formats as Save As, with their own MIME type, and keeps a host rename', async () => {
		const save = vi.fn(async () => ({ fileName: 'final.ppt' }));
		const name = await deliverSavedPresentation({
			storage: { save },
			bytes: new Uint8Array(),
			format: 'ppt',
			fileName: 'deck.ppt',
		});
		expect(name).toBe('final.ppt');
		expect(save.mock.calls[0]![1]).toMatchObject({
			mimeType: 'application/vnd.ms-powerpoint',
			reason: 'saveAs',
		});
	});

	it('rejects when the storage does', async () => {
		await expect(
			deliverSavedPresentation({
				storage: { save: () => Promise.reject(new Error('offline')) },
				bytes: new Uint8Array(),
				format: 'pptx',
				fileName: 'deck.pptx',
			}),
		).rejects.toThrow('offline');
	});

	it('downloads when there is no storage', async () => {
		const createObjectURL = vi.fn(() => 'blob:deck');
		const revoke = vi.fn();
		vi.stubGlobal('URL', { createObjectURL, revokeObjectURL: revoke });
		const click = vi.spyOn(HTMLAnchorElement.prototype, 'click').mockImplementation(() => {});
		try {
			await deliverSavedPresentation({
				bytes: new Uint8Array([1]),
				format: 'ppsx',
				fileName: 'deck.ppsx',
			});
			expect(click).toHaveBeenCalledTimes(1);
			const blob = (createObjectURL.mock.calls[0] as unknown as [Blob])[0];
			expect(blob.type).toBe(presentationSaveMimeType('ppsx'));
		} finally {
			vi.unstubAllGlobals();
		}
	});
});
