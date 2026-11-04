import type { PptxSaveFormat } from 'pptx-viewer-core';
import { afterEach, describe, expect, it, vi } from 'vitest';

import { createEditingApi } from './editing-api';
import type { EditorState } from './editor-state.svelte';

/** A stub `EditorState` whose `save` and `serialize` return fixed bytes. */
function stubEditor(): EditorState {
	return {
		dirty: true,
		slides: [],
		templateElementsBySlideId: {},
		sections: [],
		save: (_format?: PptxSaveFormat) => Promise.resolve(new Uint8Array([1, 2, 3])),
		serialize: (_format?: PptxSaveFormat) => Promise.resolve(new Uint8Array([1, 2, 3])),
	} as unknown as EditorState;
}

afterEach(() => {
	vi.unstubAllGlobals();
});

describe('createEditingApi.downloadAs', () => {
	it('downloads a .ppt save as application/vnd.ms-powerpoint, not the OOXML MIME type', async () => {
		const createObjectURL = vi.fn(() => 'blob:x');
		vi.stubGlobal('URL', { createObjectURL, revokeObjectURL: vi.fn() });
		const clickSpy = vi.spyOn(HTMLAnchorElement.prototype, 'click').mockReturnValue(undefined);
		try {
			await createEditingApi(stubEditor()).downloadAs('ppt', 'deck.pptx');
			expect(createObjectURL).toHaveBeenCalledOnce();
			const blob = createObjectURL.mock.calls[0][0] as Blob;
			expect(blob.type).toBe('application/vnd.ms-powerpoint');
		} finally {
			clickSpy.mockRestore();
		}
	});

	it('still downloads .pptx/.pptm/.ppsx with their own OOXML-family MIME types', async () => {
		const createObjectURL = vi.fn(() => 'blob:x');
		vi.stubGlobal('URL', { createObjectURL, revokeObjectURL: vi.fn() });
		const clickSpy = vi.spyOn(HTMLAnchorElement.prototype, 'click').mockReturnValue(undefined);
		const expected: Record<Exclude<PptxSaveFormat, 'ppt'>, string> = {
			pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
			ppsx: 'application/vnd.openxmlformats-officedocument.presentationml.slideshow',
			pptm: 'application/vnd.ms-powerpoint.presentation.macroenabled.12',
		};
		try {
			for (const [format, mime] of Object.entries(expected) as [
				Exclude<PptxSaveFormat, 'ppt'>,
				string,
			][]) {
				createObjectURL.mockClear();
				await createEditingApi(stubEditor()).downloadAs(format, 'deck.pptx');
				const blob = createObjectURL.mock.calls[0][0] as Blob;
				expect(blob.type).toBe(mime);
			}
		} finally {
			clickSpy.mockRestore();
		}
	});
});

describe('createEditingApi with host-owned storage', () => {
	it('hands downloadPptx and downloadAs to storage.save instead of downloading', async () => {
		const clickSpy = vi.spyOn(HTMLAnchorElement.prototype, 'click');
		const save = vi.fn(async () => undefined);
		try {
			const api = createEditingApi(stubEditor(), () => ({ save }));
			await api.downloadPptx('report.ppt');
			await api.downloadAs('ppsx', 'report.ppt');
			expect(save.mock.calls.map(([, meta]) => meta)).toEqual([
				{
					fileName: 'report.pptx',
					mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
					reason: 'save',
				},
				{
					fileName: 'report.ppsx',
					mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.slideshow',
					reason: 'saveAs',
				},
			]);
			expect(clickSpy).not.toHaveBeenCalled();
		} finally {
			clickSpy.mockRestore();
		}
	});
});

describe('createEditingApi dirty state after a save', () => {
	it('marks the document clean once the host storage has stored it', async () => {
		const editor = stubEditor();
		const onSavedClean = vi.fn();
		await createEditingApi(
			editor,
			() => ({ save: async () => undefined }),
			onSavedClean,
		).downloadPptx('deck.pptx');
		expect(editor.dirty).toBe(false);
		expect(onSavedClean).toHaveBeenCalledOnce();
	});

	it('stays unsaved when the host storage rejects', async () => {
		const editor = stubEditor();
		const onSavedClean = vi.fn();
		const api = createEditingApi(
			editor,
			() => ({ save: () => Promise.reject(new Error('offline')) }),
			onSavedClean,
		);
		await expect(api.downloadPptx('deck.pptx')).rejects.toThrow('offline');
		expect(editor.dirty).toBe(true);
		expect(onSavedClean).not.toHaveBeenCalled();
	});

	it('stays unsaved when an edit lands while the upload is in flight', async () => {
		const editor = stubEditor();
		const onSavedClean = vi.fn();
		const api = createEditingApi(
			editor,
			() => ({
				save: async () => {
					(editor as { slides: unknown[] }).slides = [{ id: 'edited during upload' }];
				},
			}),
			onSavedClean,
		);
		await api.downloadPptx('deck.pptx');
		expect(editor.dirty).toBe(true);
		expect(onSavedClean).not.toHaveBeenCalled();
	});
});
