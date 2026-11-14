// @vitest-environment jsdom
import type { PptxElement } from 'ooxml-core/pptx';
import { afterEach, describe, expect, it, vi } from 'vitest';

import { placeCaretAt } from './inline-caret';
import { setPendingCaretOffset, takePendingCaretOffset } from './inline-caret-offset';
import { attachInlineRichPaste, resolveInlineRichPaste } from './inline-rich-paste';

const element = {
	id: 'tb1',
	type: 'text',
	x: 0,
	y: 0,
	width: 100,
	height: 40,
	text: 'Hello world',
	textStyle: { fontSize: 20 },
	textSegments: [{ text: 'Hello world', style: { fontSize: 20, fontFamily: 'Calibri' } }],
} as unknown as PptxElement;

function mountEditor(text = 'Hello world'): HTMLDivElement {
	const root = document.createElement('div');
	const editor = document.createElement('div');
	editor.setAttribute('data-inline-editor', '');
	editor.setAttribute('contenteditable', 'true');
	editor.textContent = text;
	root.appendChild(editor);
	document.body.appendChild(root);
	return root;
}

function caretAt(node: Node, offset: number, end = offset): void {
	const range = document.createRange();
	range.setStart(node, offset);
	range.setEnd(node, end);
	const selection = window.getSelection()!;
	selection.removeAllRanges();
	selection.addRange(range);
}

function pasteEvent(data: Record<string, string>): ClipboardEvent {
	const event = new Event('paste', { bubbles: true, cancelable: true }) as ClipboardEvent;
	Object.defineProperty(event, 'clipboardData', {
		value: { getData: (type: string) => data[type] ?? '' },
	});
	return event;
}

afterEach(() => {
	document.body.innerHTML = '';
	takePendingCaretOffset();
	vi.useRealTimers();
});

describe('resolveInlineRichPaste', () => {
	it('splices formatted runs at the caret over the destination run style', () => {
		const root = mountEditor();
		const editor = root.firstChild as HTMLElement;
		caretAt(editor.firstChild!, 6);
		const result = resolveInlineRichPaste(editor, element, '<b>big</b>');
		expect(result?.text).toBe('Hello bigworld');
		expect(result?.caret).toBe(9);
		expect(result?.textSegments.find((s) => s.text === 'big')?.style).toEqual({
			fontSize: 20,
			fontFamily: 'Calibri',
			bold: true,
		});
	});

	it('replaces the selected text and includes the uncommitted draft', () => {
		const root = mountEditor('Hello there world');
		const editor = root.firstChild as HTMLElement;
		caretAt(editor.firstChild!, 6, 11);
		const result = resolveInlineRichPaste(editor, element, '<i>new</i>');
		expect(result?.text).toBe('Hello new world');
	});

	it('leaves plain markup and Keep Text Only to the native text paste', () => {
		const root = mountEditor();
		const editor = root.firstChild as HTMLElement;
		caretAt(editor.firstChild!, 0);
		expect(resolveInlineRichPaste(editor, element, '<span>plain</span>')).toBeUndefined();
		expect(resolveInlineRichPaste(editor, element, '<b>x</b>', 'keep-text-only')).toBeUndefined();
	});
});

describe('attachInlineRichPaste', () => {
	it('writes a rich paste through the binding and reopens the editor after it', () => {
		vi.useFakeTimers();
		const root = mountEditor();
		const editor = root.firstChild as HTMLElement;
		caretAt(editor.firstChild!, 11);
		const calls: string[] = [];
		const updateElement = vi.fn();
		const detach = attachInlineRichPaste(root, {
			getEditingElement: () => element,
			cancelEdit: () => calls.push('cancel'),
			updateElement: (...args) => {
				calls.push('update');
				updateElement(...args);
			},
			enterEdit: (id) => calls.push(`enter:${id}`),
		});
		const event = pasteEvent({ 'text/html': '<ul><li>one</li></ul>', 'text/plain': 'one' });
		editor.dispatchEvent(event);
		expect(event.defaultPrevented).toBe(true);
		expect(calls).toEqual(['cancel', 'update']);
		expect(updateElement.mock.calls[0]?.[1].text).toBe('Hello worldone');
		vi.runAllTimers();
		expect(calls).toEqual(['cancel', 'update', 'enter:tb1']);
		expect(takePendingCaretOffset('tb1')).toBe(14);
		detach();
	});

	it('pastes text only after Ctrl+Shift+V and never inserts raw markup', () => {
		const root = mountEditor();
		const editor = root.firstChild as HTMLElement;
		caretAt(editor.firstChild!, 11);
		const exec = vi.fn(() => true);
		Object.defineProperty(document, 'execCommand', { value: exec, configurable: true });
		const updateElement = vi.fn();
		const detach = attachInlineRichPaste(root, {
			getEditingElement: () => element,
			cancelEdit: vi.fn(),
			updateElement,
			enterEdit: vi.fn(),
		});
		editor.dispatchEvent(
			new KeyboardEvent('keydown', { key: 'V', ctrlKey: true, shiftKey: true, bubbles: true }),
		);
		const event = pasteEvent({ 'text/html': '<b>bold</b>', 'text/plain': 'bold' });
		editor.dispatchEvent(event);
		expect(event.defaultPrevented).toBe(true);
		expect(updateElement).not.toHaveBeenCalled();
		expect(exec).toHaveBeenCalledWith('insertText', false, 'bold');
		detach();
		delete (document as { execCommand?: unknown }).execCommand;
	});

	it('ignores pastes outside an inline text editor', () => {
		const root = mountEditor();
		const input = document.createElement('input');
		input.setAttribute('data-inline-editor', '');
		root.appendChild(input);
		const detach = attachInlineRichPaste(root, {
			getEditingElement: () => element,
			cancelEdit: vi.fn(),
			updateElement: vi.fn(),
			enterEdit: vi.fn(),
		});
		const event = pasteEvent({ 'text/plain': '1\t2' });
		input.dispatchEvent(event);
		expect(event.defaultPrevented).toBe(false);
		detach();
	});
});

describe('pending caret offset', () => {
	it('places the caret by body offset once, across paragraphs', () => {
		const root = mountEditor('');
		const editor = root.firstChild as HTMLElement;
		editor.innerHTML = '<div>ab</div><div>cd</div>';
		setPendingCaretOffset('tb1', 4);
		placeCaretAt(editor, null);
		const range = window.getSelection()!.getRangeAt(0);
		expect(range.startContainer.textContent).toBe('cd');
		expect(range.startOffset).toBe(1);
		expect(takePendingCaretOffset()).toBeNull();
	});
});
