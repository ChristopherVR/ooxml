import {
	attachInlineListController,
	buildParagraphs,
	createInlineListSeed,
	initializeInlineListDom,
} from 'ooxml-ui/pptx';
import type { InlineListController, ListConversionScope } from 'ooxml-ui/pptx';
import type { PptxElement } from 'pptx-viewer-core';
import { afterEach, describe, expect, it, vi } from 'vitest';

import { createTranslator } from '../i18n';
import { createInitialViewerState, createStore } from '../state';
import { createEditActions } from './editor-edit-ops';
import { createEditorOps } from './editor-operations';
import { toggleListType } from './editor-paragraph-mutations';

/** A three-item bulleted list: "One", "Two", "Three". */
function bulletedList(): PptxElement {
	return {
		type: 'text',
		id: 't1',
		x: 0,
		y: 0,
		width: 100,
		height: 50,
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
}

/** Mount the live inline list editor for `source` with a bare caret inside "Two". */
function caretInMiddleParagraph(source: PptxElement): InlineListController {
	const seed = createInlineListSeed(source)!;
	const root = document.createElement('div');
	document.body.append(root);
	initializeInlineListDom(root, seed);
	const controller = attachInlineListController(root, seed);
	const node = root.children[1]!.firstElementChild!.firstChild!;
	window.getSelection()!.setBaseAndExtent(node, 1, node, 1);
	return controller;
}

function markers(element: PptxElement): (string | undefined)[] {
	return buildParagraphs(element).map((paragraph) => paragraph.bulletMarker);
}

describe('listConversionScope', () => {
	let controller: InlineListController | undefined;
	afterEach(() => {
		controller?.dispose();
		controller = undefined;
		document.body.innerHTML = '';
	});

	function numberWithCaret(options?: { listConversionScope?: ListConversionScope }) {
		const source = bulletedList();
		controller = caretInMiddleParagraph(source);
		const read = controller.read();
		if (read.kind !== 'supported') {
			throw new Error(read.reason);
		}
		return markers({
			...source,
			...toggleListType(source, 'numbered', read.snapshot, options),
		} as PptxElement);
	}

	it('numbers only the caret paragraph by default (PowerPoint)', () => {
		expect(numberWithCaret()).toStrictEqual(['•', '1.', '•']);
		expect(numberWithCaret({ listConversionScope: 'paragraph' })).toStrictEqual(['•', '1.', '•']);
	});

	it("numbers the whole list around the caret with 'list'", () => {
		expect(numberWithCaret({ listConversionScope: 'list' })).toStrictEqual(['1.', '2.', '3.']);
	});

	function setupActions(getListConversionScope?: () => ListConversionScope | undefined) {
		const source = bulletedList();
		const store = createStore({
			...createInitialViewerState(),
			slides: [{ id: 's', rId: 'rId1', slideNumber: 1, elements: [source], notes: '' }],
			currentSlide: 0,
			editable: true,
			selectedElementId: source.id,
		});
		controller = caretInMiddleParagraph(source);
		const live = controller;
		const ops = createEditorOps({
			store,
			getHandler: () => null,
			onHistoryChange: vi.fn(),
			readInlineList: () => live.read(),
			formatInlineList: () => true,
		});
		const actions = createEditActions({
			doc: document,
			getTranslator: () => createTranslator(),
			store,
			ops,
			getHandler: () => null,
			setHandler: () => {},
			getListConversionScope,
		});
		const element = () => store.get().slides[0]!.elements[0]!;
		return { actions, element };
	}

	it('the Numbering action reaches one paragraph when the host leaves the option unset', () => {
		const { actions, element } = setupActions();
		actions.toggleNumberedList();
		expect(markers(element())).toStrictEqual(['•', '1.', '•']);
	});

	it("the Numbering action converts the whole list when the host passes 'list'", () => {
		const { actions, element } = setupActions(() => 'list');
		actions.toggleNumberedList();
		expect(markers(element())).toStrictEqual(['1.', '2.', '3.']);
	});
});
