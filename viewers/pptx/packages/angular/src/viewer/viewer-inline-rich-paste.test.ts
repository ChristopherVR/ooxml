import { attachInlineRichPaste, takePendingCaretOffset } from 'ooxml-ui/pptx';
import type { PptxElement, PptxSlide } from 'pptx-viewer-core';
import { afterEach, describe, expect, it, vi } from 'vitest';

import { inlineRichPasteOptions } from './viewer-inline-rich-paste';

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

afterEach(() => {
	document.body.innerHTML = '';
	takePendingCaretOffset();
	vi.useRealTimers();
});

describe('Angular inline rich text paste wiring', () => {
	it('closes the editor, writes the formatted body through the editor state and reopens', () => {
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

		let editingId: string | null = 'tb1';
		const closeEdit = vi.fn(() => {
			editingId = null;
		});
		const enterEdit = vi.fn();
		const updateElement = vi.fn();
		const detach = attachInlineRichPaste(
			root,
			inlineRichPasteOptions({
				rootElement: () => root,
				activeSlide: () => ({ id: 's1', elements: [element] }) as unknown as PptxSlide,
				activeSlideIndex: () => 3,
				editingId: () => editingId,
				closeEdit,
				enterEdit,
				updateElement,
			}),
		);
		const event = new Event('paste', { bubbles: true, cancelable: true });
		Object.defineProperty(event, 'clipboardData', {
			value: {
				getData: (type: string) =>
					type === 'text/html' ? '<span style="color:#00ff00">there</span>' : 'there',
			},
		});
		editor.dispatchEvent(event);

		expect(event.defaultPrevented).toBe(true);
		expect(closeEdit).toHaveBeenCalledOnce();
		expect(updateElement.mock.calls[0]?.slice(0, 2)).toEqual([3, 'tb1']);
		expect(updateElement.mock.calls[0]?.[2].textSegments[1]).toEqual({
			text: 'there',
			style: { fontSize: 20, color: '#00FF00' },
		});
		vi.runAllTimers();
		expect(enterEdit).toHaveBeenCalledWith('tb1');
		expect(takePendingCaretOffset('tb1')).toBe(7);
		detach();
	});
});
