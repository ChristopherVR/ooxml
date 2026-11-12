/**
 * Revision-aware acknowledgement of a Save the host persists itself, through the real vanilla
 * viewer: the snapshot leaves the deck dirty, only the current revision cleans it, and an edit
 * made while the host was saving stays unsaved without touching the undo history.
 */
import { afterEach, describe, expect, it, vi } from 'vitest';

import { createPptxViewer } from './PptxViewer';
import type { PptxViewerInstance } from './types';

let active: PptxViewerInstance[] = [];

async function deck(): Promise<Uint8Array> {
	const { PptxHandler } = await import('pptx-viewer-core');
	const { handler, data } = await PptxHandler.create({ initialSlideCount: 2 });
	data.slides[0]!.isDirty = true;
	const bytes = await handler.save(data.slides);
	handler.dispose();
	return bytes;
}

async function mount(): Promise<PptxViewerInstance> {
	const root = document.createElement('div');
	document.body.appendChild(root);
	const viewer = createPptxViewer(root, { editable: true, fileName: 'Host.pptx' });
	active.push(viewer);
	await viewer.loadFile(await deck());
	return viewer;
}

afterEach(() => {
	for (const viewer of active) {
		viewer.destroy();
	}
	active = [];
	document.body.replaceChildren();
	vi.restoreAllMocks();
});

describe('PptxViewer host-persisted save acknowledgement', () => {
	it('marks the deck clean for the current revision and keeps undo history', async () => {
		const viewer = await mount();
		viewer.addSlide(0);
		expect(viewer.isDirty()).toBe(true);
		const snapshot = await viewer.getSaveSnapshot();
		expect(snapshot.content.byteLength).toBeGreaterThan(0);
		// Taking the snapshot is not a save.
		expect(viewer.isDirty()).toBe(true);
		expect(viewer.acknowledgeSave(snapshot.revision).status).toBe('clean');
		expect(viewer.isDirty()).toBe(false);
		expect(viewer.canUndo()).toBe(true);
	});

	it('keeps an edit made while the host was persisting dirty, and ignores stale revisions', async () => {
		const viewer = await mount();
		viewer.addSlide(0);
		const first = await viewer.getSaveSnapshot();
		viewer.addSlide(0);
		expect(viewer.getSaveRevision()).toBeGreaterThan(first.revision);
		expect(viewer.acknowledgeSave(first.revision).status).toBe('pending');
		expect(viewer.isDirty()).toBe(true);
		const second = await viewer.getSaveSnapshot();
		expect(viewer.acknowledgeSave(second.revision).status).toBe('clean');
		viewer.addSlide(0);
		expect(viewer.acknowledgeSave(first.revision).status).toBe('stale');
		expect(viewer.isDirty()).toBe(true);
		expect(viewer.acknowledgeSave(9999).status).toBe('unknown');
	});
});
