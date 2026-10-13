import { registerPptxWebControls } from 'ooxml-ui/pptx';
import type { ToolbarActionId } from 'ooxml-ui/pptx';
import type { PptxElement } from 'pptx-viewer-core';
import { flushSync, mount, unmount } from 'svelte';
import { afterEach, describe, expect, it } from 'vitest';

import { EditorState } from '../../../editor/editor-state.svelte';
import ContextualTab from './ContextualTab.svelte';

/**
 * Picture Format > Size > Crop: the contextual tab draws the shared Size
 * strip, and its Crop split button drives the editor's real crop controller
 * (the same one Home > Arrange uses).
 */
registerPptxWebControls();

let cleanup: (() => void) | undefined;

afterEach(() => {
	cleanup?.();
	cleanup = undefined;
});

const PICTURE = {
	type: 'picture',
	id: 'pic',
	x: 200,
	y: 100,
	width: 200,
	height: 100,
	imageData: 'data:image/png;base64,AAAA',
} as PptxElement;
const RECT: PptxElement = { type: 'shape', id: 'a', x: 0, y: 0, width: 10, height: 10 };

function makeEditor(selected: string): EditorState {
	const editor = new EditorState({ getCurrent: () => 0, getHandler: () => null });
	editor.editable = true;
	editor.setSlides([{ id: 's1', rId: 'rId1', slideNumber: 1, elements: [PICTURE, RECT] }]);
	editor.selection.set(selected);
	return editor;
}

function render(editor: EditorState, hiddenActions?: ToolbarActionId[]): HTMLElement {
	cleanup?.();
	const target = document.createElement('div');
	document.body.appendChild(target);
	const instance = mount(ContextualTab, {
		target,
		props: { tab: 'pictureFormat', editor, hiddenActions },
	});
	flushSync();
	cleanup = () => {
		unmount(instance);
		target.remove();
	};
	return target;
}

function sizeControl(target: HTMLElement, hook: string): HTMLButtonElement {
	const node = target.querySelector<HTMLButtonElement>(
		`[data-ribbon-group="pictureFormat.size"] [data-pptx-ribbon-control="${hook}"]`,
	);
	if (!node) {
		throw new Error(`missing Size group control ${hook}`);
	}
	return node;
}

describe('picture Format > Size', () => {
	it('toggles crop mode and shows it pressed', () => {
		const editor = makeEditor('pic');
		const target = render(editor);
		expect(sizeControl(target, 'crop').getAttribute('aria-pressed')).toBe('false');
		sizeControl(target, 'crop').click();
		flushSync();
		expect(editor.cropOps.active).toBeTruthy();
		expect(sizeControl(target, 'crop').getAttribute('aria-pressed')).toBe('true');
	});

	it('applies a 1:1 aspect preset as one undoable update', () => {
		const editor = makeEditor('pic');
		const target = render(editor);
		sizeControl(target, 'crop-menu').click();
		flushSync();
		target.querySelector<HTMLButtonElement>('[data-pptx-crop-aspect="1:1"]')?.click();
		flushSync();
		expect(editor.elementById('pic')?.width).toBeCloseTo(100);
		editor.undo();
		expect(editor.elementById('pic')?.width).toBe(200);
	});

	it('is disabled for a shape and hidden when the host hides crop', () => {
		expect(sizeControl(render(makeEditor('a')), 'crop').disabled).toBeTruthy();
		const hidden = render(makeEditor('pic'), ['crop']);
		expect(sizeControl(hidden, 'crop').hidden).toBeTruthy();
		expect(sizeControl(hidden, 'crop-menu').hidden).toBeTruthy();
	});
});
