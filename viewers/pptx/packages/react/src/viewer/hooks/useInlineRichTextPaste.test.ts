// @vitest-environment happy-dom
import { attachInlineRichPaste, takePendingCaretOffset } from 'ooxml-ui/pptx';
import type { PptxElement } from 'pptx-viewer-core';
import { afterEach, describe, expect, it, vi } from 'vitest';

import { createInlineRichTextPasteOptions } from './useInlineRichTextPaste';

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

describe('React inline rich text paste wiring', () => {
	it('closes the editor, writes the formatted body and reopens editing', () => {
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

		const editingId = { current: 'tb1' as string | null };
		const setId = vi.fn((id: string | null) => {
			editingId.current = id;
		});
		const updateElementById = vi.fn();
		const detach = attachInlineRichPaste(
			root,
			createInlineRichTextPasteOptions(() => ({
				elementLookup: new Map([['tb1', element]]),
				inlineEditingElementIdRef: editingId,
				setInlineEditingElementId: setId,
				setInlineEditingText: vi.fn(),
				updateElementById,
			})),
		);
		const event = new Event('paste', { bubbles: true, cancelable: true });
		Object.defineProperty(event, 'clipboardData', {
			value: { getData: (type: string) => (type === 'text/html' ? ' <b>there</b>' : 'there') },
		});
		editor.dispatchEvent(event);

		expect(event.defaultPrevented).toBe(true);
		expect(setId).toHaveBeenCalledWith(null);
		const patch = updateElementById.mock.calls[0]?.[1];
		expect(patch.text).toBe('Hithere');
		expect(patch.textSegments[1]).toEqual({ text: 'there', style: { fontSize: 20, bold: true } });
		vi.runAllTimers();
		expect(setId).toHaveBeenLastCalledWith('tb1');
		expect(takePendingCaretOffset('tb1')).toBe(7);
		detach();
	});
});
