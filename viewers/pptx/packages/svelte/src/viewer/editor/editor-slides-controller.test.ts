import {
	captureSlideClipboard,
	createSlidePaste,
	captureSlideClipboardTemplates,
} from 'ooxml-ui/pptx';
import type { PptxElement, PptxHandler, PptxSlide } from 'pptx-viewer-core';
import { describe, expect, it, vi } from 'vitest';

import { EditorState } from './editor-state.svelte';

/**
 * The sorter's context menu can target an arbitrary right-clicked slide, not
 * only the active one. `deleteCurrentSlide` / `duplicateCurrentSlide` only
 * ever operated on `currentSlideIndex`, so a Svelte host had no way to act on
 * a non-active slide from the sorter without first selecting it. These cover
 * the new arbitrary-index variants, including the "does the active slide stay
 * the same logical slide" bookkeeping when the mutated index sits before it.
 */

function slide(id: string, hidden?: boolean): PptxSlide {
	return { id, rId: `rId-${id}`, slideNumber: 1, elements: [], hidden };
}

function setup(current: number): { editor: EditorState } {
	const editor = new EditorState({ getCurrent: () => current, getHandler: () => null });
	editor.editable = true;
	editor.setSlides([slide('a'), slide('b'), slide('c')]);
	return { editor };
}

describe('svelte slides controller: arbitrary-index operations', () => {
	it('pastes the copied template artwork after deleting its source and template cache', () => {
		const { editor } = setup(0);
		const artwork = {
			id: 'template',
			type: 'text',
			x: 0,
			y: 0,
			width: 100,
			height: 30,
			text: 'Captured artwork',
		} as PptxElement;
		editor.templateElementsBySlideId = { a: [artwork] };
		const clipboard = captureSlideClipboard(editor.slides, ['a']);
		const templates = captureSlideClipboardTemplates(editor.templateElementsBySlideId, ['a']);
		(artwork as PptxElement & { text: string }).text = 'Changed artwork';
		editor.slidesOps.deleteSlideAtIndex(0);
		expect(editor.templateElementsBySlideId.a).toBeUndefined();
		const request = createSlidePaste(clipboard, 2, templates);
		editor.slidesOps.pasteSlides(request);
		expect(
			(
				editor.templateElementsBySlideId[request.copies[0]!.id]![0] as PptxElement & {
					text: string;
				}
			).text,
		).toBe('Captured artwork');
		editor.undo();
		expect(editor.slides.map((item) => item.id)).toStrictEqual(['b', 'c']);
	});

	it('pastes a range after the destination and undoes the whole paste in one step', () => {
		const { editor } = setup(0);
		const clipboard = captureSlideClipboard(editor.slides, ['a', 'b']);
		expect(editor.slidesOps.pasteSlides(createSlidePaste(clipboard, 3))).toBe(3);
		expect(editor.slides).toHaveLength(5);
		expect(editor.slides.slice(3).map((item) => item.sourceSlideId)).toStrictEqual(['a', 'b']);
		editor.undo();
		expect(editor.slides.map((item) => item.id)).toStrictEqual(['a', 'b', 'c']);
	});

	it('deleteSlideAtIndex on a non-active slide keeps the active slide the same logical slide', () => {
		const { editor } = setup(1); // active slide is 'b'
		const nextIndex = editor.slidesOps.deleteSlideAtIndex(0); // delete 'a'
		expect(editor.slides.map((s) => s.id)).toStrictEqual(['b', 'c']);
		// 'b' is now at index 0; the caller should navigate there.
		expect(nextIndex).toBe(0);
	});

	it('deleteSlideAtIndex on a slide AFTER the active one leaves the active index untouched', () => {
		const { editor } = setup(0); // active slide is 'a'
		const nextIndex = editor.slidesOps.deleteSlideAtIndex(2); // delete 'c'
		expect(editor.slides.map((s) => s.id)).toStrictEqual(['a', 'b']);
		expect(nextIndex).toBe(0);
	});

	it('deleteSlideAtIndex on the active slide behaves like deleteCurrentSlide', () => {
		const { editor } = setup(1);
		const nextIndex = editor.slidesOps.deleteSlideAtIndex(1); // delete 'b' (active)
		expect(editor.slides.map((s) => s.id)).toStrictEqual(['a', 'c']);
		expect(nextIndex).toBe(1);
	});

	it('duplicateSlideAtIndex on a non-active slide before the active one shifts it forward', () => {
		const { editor } = setup(1); // active slide is 'b'
		const nextIndex = editor.slidesOps.duplicateSlideAtIndex(0); // duplicate 'a'
		expect(editor.slides.map((s) => s.id)).toHaveLength(4);
		expect(editor.slides[0]?.id).toBe('a');
		expect(editor.slides[1]?.id).not.toBe('b'); // the duplicate lands here
		expect(editor.slides[2]?.id).toBe('b');
		// 'b' moved from index 1 to index 2.
		expect(nextIndex).toBe(2);
	});

	it('duplicateSlideAtIndex on a slide after the active one leaves the active index untouched', () => {
		const { editor } = setup(0); // active slide is 'a'
		const nextIndex = editor.slidesOps.duplicateSlideAtIndex(2); // duplicate 'c'
		expect(editor.slides).toHaveLength(4);
		expect(nextIndex).toBe(0);
	});

	it('toggleSlideHidden flips only the target slide, regardless of the active index', () => {
		const { editor } = setup(1);
		editor.slidesOps.toggleSlideHidden(0);
		expect(editor.slides[0]?.hidden).toBeTruthy();
		expect(editor.slides[1]?.hidden).toBeFalsy();
		editor.slidesOps.toggleSlideHidden(0);
		expect(editor.slides[0]?.hidden).toBeFalsy();
	});

	it('every operation no-ops when not editable', () => {
		const { editor } = setup(0);
		editor.editable = false;
		expect(editor.slidesOps.deleteSlideAtIndex(0)).toBeNull();
		expect(editor.slidesOps.duplicateSlideAtIndex(0)).toBeNull();
		editor.slidesOps.toggleSlideHidden(0);
		expect(editor.slides).toHaveLength(3);
		expect(editor.slides[0]?.hidden).toBeFalsy();
	});
});

describe('svelte slides controller: layout changes', () => {
	it('fetches and replaces inherited artwork after applying a layout', async () => {
		const nextLayoutPath = 'ppt/slideLayouts/slideLayout2.xml';
		const oldArtwork = [{ id: 'old-layout-art', type: 'shape' } as PptxElement];
		const nextArtwork = [{ id: 'new-layout-art', type: 'shape' } as PptxElement];
		const handler = {
			applyLayoutToSlide: vi.fn(
				async (index: number, layoutPath: string, slides: PptxSlide[]) =>
					({
						...slides[index],
						layoutPath,
						elements: [{ id: 'placeholder', type: 'shape' } as PptxElement],
					}) as PptxSlide,
			),
			getTemplateElementsForSlide: vi.fn().mockResolvedValue(nextArtwork),
		} as unknown as PptxHandler;
		const editor = new EditorState({ getCurrent: () => 0, getHandler: () => handler });
		editor.editable = true;
		editor.setSlides([slide('a'), slide('b')]);
		editor.templateElementsBySlideId = { a: oldArtwork };

		const appliedIndex = await editor.slidesOps.applyLayout(nextLayoutPath);

		expect(appliedIndex).toBe(0);
		expect(handler.applyLayoutToSlide).toHaveBeenCalledOnce();
		expect(handler.applyLayoutToSlide.mock.calls[0]?.slice(0, 2)).toStrictEqual([
			0,
			nextLayoutPath,
		]);
		expect(handler.getTemplateElementsForSlide).toHaveBeenCalledWith('a');
		expect(editor.slides[0]?.layoutPath).toBe(nextLayoutPath);
		expect(editor.templateElementsBySlideId.a).toStrictEqual(nextArtwork);
	});

	describe('and Undo', () => {
		const L1 = 'ppt/slideLayouts/slideLayout1.xml';
		const L2 = 'ppt/slideLayouts/slideLayout2.xml';
		const L3 = 'ppt/slideLayouts/slideLayout3.xml';

		/** The core contract: the slide at `index` is replaced in the array it is given, and returned. */
		function inPlaceEditor() {
			const handler = {
				applyLayoutToSlide: vi.fn(
					async (index: number, layoutPath: string, slides: PptxSlide[]) => {
						const updated = { ...slides[index]!, layoutPath, layoutName: 'Remapped' };
						slides[index] = updated;
						return updated;
					},
				),
				getTemplateElementsForSlide: vi.fn().mockResolvedValue([]),
			};
			const editor = new EditorState({
				getCurrent: () => 0,
				getHandler: () => handler as unknown as PptxHandler,
			});
			editor.editable = true;
			editor.setSlides([{ ...slide('a'), layoutPath: L1 }, slide('b')]);
			return { editor, handler };
		}

		it('one Undo takes back one layout pick', async () => {
			const { editor } = inPlaceEditor();
			await editor.slidesOps.applyLayout(L2);
			await editor.slidesOps.applyLayout(L3);
			expect(editor.slides[0]?.layoutPath).toBe(L3);
			editor.undo();
			expect(editor.slides[0]?.layoutPath).toBe(L2);
			editor.undo();
			expect(editor.slides[0]?.layoutPath).toBe(L1);
			expect(editor.canUndo).toBeFalsy();
			editor.redo();
			expect(editor.slides[0]?.layoutPath).toBe(L2);
		});

		it('one Undo takes back a new slide from a layout', async () => {
			const { editor, handler } = inPlaceEditor();
			expect(editor.slidesOps.insertSlideFromLayout(L2)).toBe(1);
			await vi.waitFor(() => expect(handler.getTemplateElementsForSlide).toHaveBeenCalledOnce());
			await vi.waitFor(() => expect(editor.slides[1]?.layoutName).toBe('Remapped'));
			expect(editor.slides).toHaveLength(3);
			editor.undo();
			expect(editor.slides.map((each) => each.id)).toStrictEqual(['a', 'b']);
			expect(editor.canUndo).toBeFalsy();
		});
	});
});
