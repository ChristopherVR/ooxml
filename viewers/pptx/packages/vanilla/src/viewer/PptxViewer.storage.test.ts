import type { DocumentSaveMeta, DocumentStorage } from 'ooxml-ui/pptx';
import { afterEach, describe, expect, it, vi } from 'vitest';

import { createPptxViewer, PptxViewer } from './PptxViewer';
import type { PptxViewerInstance, PptxViewerOptions } from './types';

/**
 * Host-owned storage through the vanilla viewer: File > Save is `save`, Save As
 * is `saveAs`, and a failed save reaches the host's `onError` and the
 * save-failed notice instead of rejecting.
 */
let active: PptxViewerInstance[] = [];

async function deck(): Promise<Uint8Array> {
	const { PptxHandler } = await import('pptx-viewer-core');
	const { handler, data } = await PptxHandler.create({ initialSlideCount: 1 });
	const bytes = await handler.save(data.slides);
	handler.dispose();
	return bytes;
}

async function mount(options: PptxViewerOptions): Promise<PptxViewer> {
	const container = document.createElement('div');
	document.body.appendChild(container);
	const viewer = createPptxViewer(container, options);
	active.push(viewer);
	await viewer.loadFile(await deck());
	return viewer as PptxViewer;
}

afterEach(async () => {
	for (const viewer of active) {
		viewer.destroy();
	}
	active = [];
	document.body.replaceChildren();
	vi.restoreAllMocks();
	await new Promise<void>((resolve) => {
		setTimeout(resolve, 0);
	});
});

describe('PptxViewer host-owned storage', () => {
	it('reports Save as save and Save As as saveAs', async () => {
		const metas: DocumentSaveMeta[] = [];
		const storage: DocumentStorage = { save: async (_bytes, meta) => void metas.push(meta) };
		const viewer = await mount({ storage, fileName: 'Quarterly.pptx' });
		await viewer.downloadPptx('Quarterly.pptx');
		await viewer.downloadAs('pptx', 'Quarterly.pptx');
		expect(metas.map((meta) => [meta.fileName, meta.reason])).toStrictEqual([
			['Quarterly.pptx', 'save'],
			['Quarterly.pptx', 'saveAs'],
		]);
	});

	it('reports a failed save to onError and the save-failed notice without rejecting', async () => {
		vi.spyOn(console, 'error').mockImplementation(() => {});
		const onError = vi.fn();
		const viewer = await mount({
			storage: { save: () => Promise.reject(new Error('offline')) },
			onError,
		});
		await expect(viewer.downloadPptx()).resolves.toBeUndefined();
		expect(onError).toHaveBeenCalledWith('Save failed: offline', expect.any(Error));
		expect(viewer.store.get().compatToasts.map((toast) => toast.id)).toContain('save-failed');
	});
});
