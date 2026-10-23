// @vitest-environment jsdom
import { createDocument } from 'ooxml-core/docx';
import { EditorState, TextSelection } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { modelToDoc } from './model-adapter';
import type { FloatBox } from './print-float-geometry';
import { place } from './print-float-locate';
import { attachPrintFloatSelection } from './print-float-selection';
import { insertDrawnShape, selectedShape } from './shape-commands';

const cleanup: Array<() => void> = [];
afterEach(() => {
	while (cleanup.length) cleanup.pop()?.();
	vi.unstubAllGlobals();
});

/** The rendered float of paragraph `p` at `box`. */
function floatElement(box: FloatBox): HTMLElement {
	const float = document.createElement('div');
	float.className = 'dve-print-float';
	Object.assign(float.dataset, { blockId: 'p', floatIndex: '0', originX: '0', originY: '0' });
	place(float, box);
	return float;
}

/** What a layout puts in the Print Layout pages: one sheet holding the float. */
function laidOut(box: FloatBox): HTMLElement {
	const pages = document.createElement('div');
	const sheet = document.createElement('div');
	sheet.className = 'dve-print-page';
	sheet.append(floatElement(box));
	pages.append(sheet);
	return pages;
}

/** The next task, where a press that follows a relayout is handled (stubbed frames never run). */
const nextTask = () => new Promise((resolve) => setTimeout(resolve, 0));

/** A drawn shape, selected and shown with its frame on a Print Layout page. */
function setup() {
	const frames = new Map<number, FrameRequestCallback>();
	let lastFrame = 0;
	vi.stubGlobal('requestAnimationFrame', (callback: FrameRequestCallback) => {
		frames.set(++lastFrame, callback);
		return lastFrame;
	});
	vi.stubGlobal('cancelAnimationFrame', (id: number) => frames.delete(id));
	const model = createDocument();
	model.blocks = [{ type: 'paragraph', id: 'p', runs: [{ text: 'Hello' }] }];
	const doc = modelToDoc(model);
	const paper = document.createElement('div');
	const pages = document.createElement('div');
	document.body.append(paper, pages);
	const view = new EditorView(paper, {
		state: EditorState.create({ doc, selection: TextSelection.create(doc, 2) }),
	});
	insertDrawnShape(view, 'rect', 2, { x: 10, y: 5, widthPx: 100, heightPx: 50 });
	pages.append(laidOut({ x: 10, y: 5, width: 100, height: 50 }));
	const floats = attachPrintFloatSelection({
		pages,
		paper,
		view: () => view,
		printActive: () => true,
	});
	cleanup.push(() => {
		floats.destroy();
		view.destroy();
		paper.remove();
		pages.remove();
	});
	floats.refresh();
	for (const callback of frames.values()) callback(0);
	frames.clear();
	expect(pages.querySelectorAll('[data-float-handle]')).toHaveLength(8);
	return { view, pages };
}

describe('a floating shape selected on a Print Layout page', () => {
	it('keeps its handles through a relayout, so a press right after it resizes', async () => {
		const { view, pages } = setup();
		// The relayout after a move replaces the pages, the moved float with them.
		pages.replaceChildren(laidOut({ x: 40, y: 5, width: 100, height: 50 }));
		await nextTask();
		const frame = pages.querySelector<HTMLElement>('.dve-print-float-frame');
		expect(frame?.style.left).toBe('40px');
		frame!
			.querySelector('[data-float-handle="e"]')!
			.dispatchEvent(
				new PointerEvent('pointerdown', { bubbles: true, button: 0, clientX: 140, clientY: 30 }),
			);
		window.dispatchEvent(new PointerEvent('pointermove', { clientX: 170, clientY: 30 }));
		window.dispatchEvent(new PointerEvent('pointermove', { clientX: 200, clientY: 30 }));
		window.dispatchEvent(new PointerEvent('pointerup', { clientX: 200, clientY: 30 }));
		expect(selectedShape(view.state)).toMatchObject({ widthPx: 160, heightPx: 50 });
		expect(JSON.parse(String(view.state.doc.nodeAt(1)?.attrs.placement))).toMatchObject({
			offsetXPx: 40,
			offsetYPx: 5,
		});
	});

	it('draws its handles again when the page window mounts its page again', async () => {
		const { pages } = setup();
		const root = pages.firstElementChild as HTMLElement;
		const sheet = root.querySelector<HTMLElement>('.dve-print-page')!;
		// Scrolled away and back: the sheet is emptied, drawn again and the window's version bumped.
		sheet.replaceChildren(floatElement({ x: 10, y: 5, width: 100, height: 50 }));
		root.dataset.renderVersion = '2';
		await nextTask();
		expect(sheet.querySelectorAll('.dve-print-float-frame [data-float-handle]')).toHaveLength(8);
	});
});
