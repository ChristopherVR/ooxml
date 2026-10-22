/**
 * The status bar shows the autosave state, and keeps showing it after a locale switch rebuilds
 * the chrome: a recovery snapshot leaves the deck dirty, so a rebuilt bar that forgot the save
 * would say "Unsaved changes" until the next edit is saved.
 */
import { PptxHandler } from 'pptx-viewer-core';
import { afterEach, describe, expect, it, vi } from 'vitest';

// No IndexedDB in this environment: accept the snapshot write. Everything else stays real.
vi.mock(import('ooxml-ui/pptx'), async (importOriginal) => ({
	...(await importOriginal()),
	saveAutosaveSnapshot: async () => true,
	probeAutosaveRecovery: async () => null,
}));

const { createPptxViewer } = await import('./PptxViewer');

let cleanup: (() => void) | undefined;
afterEach(() => {
	cleanup?.();
	cleanup = undefined;
	document.body.replaceChildren();
});

const saveText = (container: HTMLElement) =>
	container.querySelector('pptx-ui-status-bar')?.shadowRoot?.querySelector('[data-item="save"]')
		?.textContent;

describe('vanilla status bar and autosave', () => {
	it('shows the last save, also after the chrome is rebuilt', async () => {
		const { handler, data } = await PptxHandler.create({ initialSlideCount: 2 });
		const bytes = await handler.save(data.slides);
		handler.dispose();
		const container = document.createElement('div');
		document.body.append(container);
		const viewer = createPptxViewer(container, {
			editable: true,
			fileName: 'deck.pptx',
			messages: {
				de: {
					'pptx.autosave.saved': 'Gespeichert {{time}}',
					'pptx.autosave.justNow': 'gerade eben',
					'pptx.statusBar.unsavedChanges': 'Ungespeicherte Änderungen',
				},
			},
		});
		cleanup = () => viewer.destroy();
		await viewer.loadFile(bytes);
		(viewer as unknown as { store: { set(patch: object): void } }).store.set({ dirty: true });
		expect(saveText(container)).toBe('Unsaved changes');

		await viewer.autosaveNow();
		expect(saveText(container)).toBe('Saved just now');

		viewer.setLocale('de');
		expect(saveText(container)).toBe('Gespeichert gerade eben');
	});
});
