import {
	attachInlineListController,
	buildParagraphs,
	createInlineListSeed,
	initializeInlineListDom,
	registerPptxWebControls,
} from 'ooxml-ui/pptx';
import type { ListConversionScope } from 'ooxml-ui/pptx';
import type { PptxElement } from 'pptx-viewer-core';
import { flushSync, mount, unmount } from 'svelte';
import { afterEach, describe, expect, it } from 'vitest';

import { EditorState } from '../../../editor/editor-state.svelte';
import ParagraphGroup from './ParagraphGroup.svelte';

registerPptxWebControls();

let cleanup: (() => void) | undefined;
afterEach(() => {
	cleanup?.();
	cleanup = undefined;
});

function mountGroup(editable = true) {
	const editor = new EditorState({ getCurrent: () => 0, getHandler: () => null });
	editor.editable = editable;
	editor.setSlides([
		{
			id: 's1',
			rId: 'rId1',
			slideNumber: 1,
			elements: [
				{
					id: 't1',
					type: 'text',
					x: 0,
					y: 0,
					width: 100,
					height: 20,
					text: 'Hello',
					textStyle: { align: 'center' },
				},
			] as never,
		},
	]);
	editor.selection.setAll(['t1']);
	const target = document.createElement('div');
	document.body.appendChild(target);
	const instance = mount(ParagraphGroup, { target, props: { editor } });
	flushSync();
	cleanup = () => {
		unmount(instance);
		target.remove();
	};
	const button = (id: string) =>
		target.querySelector<HTMLButtonElement>(`[data-ribbon-control="home.paragraph.${id}"]`)!;
	const style = () => {
		const element = editor.slides[0]?.elements[0];
		return element?.type === 'text' ? element.textStyle : undefined;
	};
	return { button, style };
}

describe('paragraphGroup', () => {
	it('reflects the alignment and applies a new one through patchSelected', () => {
		const { button, style } = mountGroup();
		expect(button('alignCenter').getAttribute('aria-pressed')).toBe('true');
		button('alignRight').click();
		flushSync();
		expect(style()?.align).toBe('right');
		expect(button('alignRight').getAttribute('aria-pressed')).toBe('true');
		expect(button('alignCenter').getAttribute('aria-pressed')).toBe('false');
	});

	it('steps the indent and gates the strip when read-only', () => {
		const { button, style } = mountGroup();
		button('increaseIndent').click();
		flushSync();
		expect(style()?.paragraphMarginLeft).toBeGreaterThan(0);
		cleanup?.();
		const readOnly = mountGroup(false);
		expect(readOnly.button('alignLeft').disabled).toBeTruthy();
		expect(readOnly.button('increaseIndent').disabled).toBeTruthy();
	});
});

describe('paragraphGroup listConversionScope', () => {
	/**
	 * Mount the group over a three-item bulleted list that is being edited
	 * inline, with a bare caret inside the middle item, then press Numbering.
	 */
	function numberWithCaret(scope?: ListConversionScope): (string | undefined)[] {
		const source = {
			id: 't1',
			type: 'text',
			x: 0,
			y: 0,
			width: 100,
			height: 60,
			text: 'One\nTwo\nThree',
			textStyle: {},
			textSegments: [
				{ text: 'One', style: {}, bulletInfo: { char: '•' } },
				{ text: '\n', style: {}, isParagraphBreak: true },
				{ text: 'Two', style: {}, bulletInfo: { char: '•' } },
				{ text: '\n', style: {}, isParagraphBreak: true },
				{ text: 'Three', style: {}, bulletInfo: { char: '•' } },
			],
		} as PptxElement;
		const editor = new EditorState({
			getCurrent: () => 0,
			getHandler: () => null,
			getListConversionScope: () => scope,
		});
		editor.editable = true;
		editor.setSlides([{ id: 's1', rId: 'rId1', slideNumber: 1, elements: [source] }]);
		editor.selection.setAll(['t1']);

		const seed = createInlineListSeed(source)!;
		const root = document.createElement('div');
		document.body.append(root);
		initializeInlineListDom(root, seed);
		const controller = attachInlineListController(root, seed);
		editor.inlineListController = controller;
		const node = root.children[1]!.firstElementChild!.firstChild!;
		window.getSelection()!.setBaseAndExtent(node, 1, node, 1);

		const target = document.createElement('div');
		document.body.appendChild(target);
		const instance = mount(ParagraphGroup, { target, props: { editor } });
		flushSync();
		cleanup = () => {
			unmount(instance);
			controller.dispose();
			target.remove();
			root.remove();
		};
		// Numbering is a split control: its slot holds the toggle button and the gallery.
		target
			.querySelector<HTMLButtonElement>('[data-ribbon-control="home.paragraph.numbering"] button')!
			.click();
		flushSync();
		return buildParagraphs(editor.slides[0]!.elements[0]!).map((p) => p.bulletMarker);
	}

	it('numbers only the caret paragraph when the prop is unset (PowerPoint)', () => {
		expect(numberWithCaret()).toStrictEqual(['•', '1.', '•']);
	});

	it("numbers the whole list around the caret when the prop is 'list'", () => {
		expect(numberWithCaret('list')).toStrictEqual(['1.', '2.', '3.']);
	});
});
