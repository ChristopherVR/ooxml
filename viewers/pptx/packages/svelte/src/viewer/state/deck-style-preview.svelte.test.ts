import type { PptxHandler, PptxSlide } from 'pptx-viewer-core';
import { flushSync } from 'svelte';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import { registerCollaborationEffects } from '../collab/collaboration-effects.svelte';
import { createEditorDeckStyleHost } from '../editor/editor-deck-style';
import { EditorState } from '../editor/editor-state.svelte';
import { AutosaveController } from './autosave.svelte';

/**
 * A Design gallery hover preview must not reach collaborators, mark the document dirty, or end up
 * in a recovery snapshot. Real editor, real host, real autosave and collaboration effects, wired
 * the way `create-viewer-state-collab` wires them.
 */
const { saveSnapshot } = vi.hoisted(() => ({ saveSnapshot: vi.fn().mockResolvedValue(true) }));
vi.mock(import('ooxml-ui/pptx'), async (importOriginal) => ({
	...(await importOriginal()),
	saveAutosaveSnapshot: saveSnapshot,
}));

const slide = (color: string): PptxSlide =>
	({ id: 's1', rId: 'r1', slideNumber: 1, elements: [], backgroundColor: color }) as PptxSlide;

const handler = {
	resolveStyleMatrixReferences: () => ({ shapeStyle: {} }),
	applyBackgroundStyleToSlides: (slides: PptxSlide[]) => ({
		slides: slides.map((s) => ({ ...s, backgroundColor: '#000000' })),
		templateElementsBySlideId: {},
	}),
	save: vi.fn(async (slides: PptxSlide[]) => {
		void slides;
		return new Uint8Array([1]);
	}),
} as unknown as PptxHandler;

beforeEach(() => {
	vi.useFakeTimers();
	saveSnapshot.mockClear();
	(handler.save as ReturnType<typeof vi.fn>).mockClear();
});
afterEach(() => vi.useRealTimers());

describe('Design gallery preview in the Svelte editor', () => {
	it('is not published, not dirty and not in a recovery snapshot', async () => {
		const editor = new EditorState({ getCurrent: () => 0, getHandler: () => handler });
		editor.editable = true;
		editor.setSlides([slide('#FFFFFF')]);
		const published: PptxSlide[][] = [];
		let ctl!: AutosaveController;
		const dispose = $effect.root(() => {
			ctl = new AutosaveController({
				getEnabled: () => true,
				getIntervalMs: () => 1000,
				getFilePath: () => 'deck.pptx',
				getSlides: () => editor.committedRenderedSlides,
				getPreviewHeld: () => editor.deckStyleHeld,
				getDocumentKey: () => [editor.slides, editor.templateElementsBySlideId],
				getHandler: () => handler,
				getLoadCount: () => 1,
				getSeedNonce: () => editor.seedNonce,
			});
			registerCollaborationEffects({
				getConfig: () => undefined,
				getSlides: () => editor.committedRenderedSlides,
				syncConfig: () => {},
				isPublishable: () => true,
				flushLocalSlides: (slides) => published.push(slides),
				stop: () => {},
				rejoin: () => {},
				leaveOnBeforeUnload: () => false,
			});
		});
		flushSync();
		const baseline = published.length;

		const host = createEditorDeckStyleHost(editor);
		host.preview({ kind: 'backgroundStyle', styleIndex: 4, scope: 'all' });
		flushSync();
		expect(editor.slides[0]!.backgroundColor).toBe('#000000');
		// Even a snapshot forced now holds the real slides.
		await ctl.save();
		expect(
			(handler.save as ReturnType<typeof vi.fn>).mock.calls.at(-1)![0][0].backgroundColor,
		).toBe('#FFFFFF');
		host.endPreview();
		flushSync();
		await vi.advanceTimersByTimeAsync(5000);

		expect(ctl.isDirty).toBe(false);
		for (const slides of published.slice(baseline)) {
			expect(slides[0]!.backgroundColor).toBe('#FFFFFF');
		}
		expect(editor.dirty).toBe(false);
		dispose();
	});
});
