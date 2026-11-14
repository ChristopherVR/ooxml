import { takePendingCaretOffset } from 'ooxml-ui/pptx';
import type { PptxElement } from 'pptx-viewer-core';
import { afterEach, describe, expect, it, vi } from 'vitest';

import type { Store, ViewerState } from '../state';
import { attachCanvasInlineRichPaste } from './editor-inline-rich-paste';

const element = {
	id: 'tb1',
	type: 'text',
	x: 0,
	y: 0,
	width: 100,
	height: 40,
	text: 'Hi',
	textSegments: [{ text: 'Hi', style: { fontSize: 20 } }],
} as unknown as PptxElement;

const state = {
	slides: [{ id: 's1', elements: [element] }],
	currentSlide: 0,
	editTemplateMode: false,
	templateElementsBySlideId: {},
} as unknown as ViewerState;
const store = { get: () => state } as unknown as Store<ViewerState>;

afterEach(() => {
	document.body.innerHTML = '';
	takePendingCaretOffset();
	vi.useRealTimers();
});

describe('Vanilla inline rich text paste wiring', () => {
	it('closes the editor, writes the body as one history step and reopens editing', () => {
		vi.useFakeTimers();
		const root = document.createElement('div');
		const editor = document.createElement('div');
		editor.setAttribute('data-inline-editor', '');
		editor.setAttribute('contenteditable', 'true');
		editor.textContent = 'Hi';
		root.append(editor);
		document.body.append(root);
		const range = document.createRange();
		range.setStart(editor.firstChild!, 2);
		range.collapse(true);
		window.getSelection()!.removeAllRanges();
		window.getSelection()!.addRange(range);

		const calls: string[] = [];
		let inlineId: string | null = 'tb1';
		const previewElement = vi.fn(() => calls.push('preview'));
		const detach = attachCanvasInlineRichPaste(root, {
			store,
			inlineElementId: () => inlineId,
			closeInline: (commit) => {
				calls.push(`close:${commit}`);
				inlineId = null;
			},
			openInlineEdit: (id) => calls.push(`open:${id}`),
			pushHistory: () => calls.push('history'),
			previewElement,
			commitChange: () => calls.push('commit'),
		});
		const event = new Event('paste', { bubbles: true, cancelable: true });
		Object.defineProperty(event, 'clipboardData', {
			value: {
				getData: (type: string) => (type === 'text/html' ? '<u>there</u>' : 'there'),
			},
		});
		editor.dispatchEvent(event);

		expect(event.defaultPrevented).toBe(true);
		expect(calls).toEqual(['close:false', 'history', 'preview', 'commit']);
		const patch = (
			previewElement.mock.calls[0] as unknown as [string, { textSegments: unknown[] }]
		)[1];
		expect(patch.textSegments[1]).toEqual({
			text: 'there',
			style: { fontSize: 20, underline: true },
		});
		vi.runAllTimers();
		expect(calls.at(-1)).toBe('open:tb1');
		expect(takePendingCaretOffset('tb1')).toBe(7);
		detach();
	});
});
