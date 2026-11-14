import { attachInlineRichPaste, takePendingCaretOffset } from 'ooxml-ui/pptx';
import type { PptxElement } from 'pptx-viewer-core';
import { afterEach, describe, expect, it, vi } from 'vitest';

import { inlineRichTextPasteOptions } from './useInlineRichTextPaste';

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

function mountEditor(): { root: HTMLElement; editor: HTMLElement } {
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
	return { root, editor };
}

function paste(target: HTMLElement, html: string): Event {
	const event = new Event('paste', { bubbles: true, cancelable: true });
	Object.defineProperty(event, 'clipboardData', {
		value: { getData: (type: string) => (type === 'text/html' ? html : 'there') },
	});
	target.dispatchEvent(event);
	return event;
}

afterEach(() => {
	document.body.innerHTML = '';
	takePendingCaretOffset();
	vi.useRealTimers();
});

describe('Vue inline rich text paste wiring', () => {
	it('closes the editor, writes the formatted body through ops and reopens editing', () => {
		vi.useFakeTimers();
		const { root, editor } = mountEditor();
		const cancelInlineEdit = vi.fn();
		const enterInlineEdit = vi.fn();
		const updateElement = vi.fn();
		const detach = attachInlineRichPaste(
			root,
			inlineRichTextPasteOptions({
				editingElement: () => element,
				cancelInlineEdit,
				enterInlineEdit,
				updateElement,
			}),
		);
		expect(paste(editor, '<i>there</i>').defaultPrevented).toBe(true);
		expect(cancelInlineEdit).toHaveBeenCalledOnce();
		expect(updateElement.mock.calls[0]?.[0]).toBe('tb1');
		expect(updateElement.mock.calls[0]?.[1].textSegments[1]).toEqual({
			text: 'there',
			style: { fontSize: 20, italic: true },
		});
		vi.runAllTimers();
		expect(enterInlineEdit).toHaveBeenCalledWith('tb1');
		expect(takePendingCaretOffset('tb1')).toBe(7);
		detach();
	});
});
