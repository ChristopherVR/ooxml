/**
 * The canvas's plain-text inline editor honours the shared inline-editor
 * chords (`mapInlineEditorChord`) before its own Enter rules.
 *
 * It used to check only the mapper's `suppress` result, so Ctrl/Cmd+Shift+Enter
 * (which the mapper resolves to `commit`) fell through to the Shift+Enter
 * branch and inserted a line break. It also committed on Enter while an IME
 * composition was open, when that Enter belongs to the IME.
 *
 * Enter and Shift+Enter now belong to the browser (a new paragraph, a line
 * break), as in every other binding, after the space a soft wrap left before
 * the caret is dropped (shared `spaceBeforeCaretRange`). The editor used to be
 * a `<textarea>` that committed on Enter.
 *
 * The handler is called on the prototype with only the members it touches, as
 * `canvas-element-locks.test.ts` does, so no injected services are needed.
 */
import { afterEach, describe, expect, it, vi } from 'vitest';

import { SlideCanvasComponent } from './slide-canvas.component';

interface EditorHarness {
	editor: HTMLDivElement;
	blur: ReturnType<typeof vi.spyOn>;
	press: (init: KeyboardEventInit, composing?: boolean) => KeyboardEvent;
}

let cleanup: (() => void) | null = null;

afterEach(() => {
	cleanup?.();
	cleanup = null;
});

/** A seeded editor reading "Hello " with the caret after the trailing space. */
function harness(): EditorHarness {
	const editor = document.createElement('div');
	editor.contentEditable = 'true';
	const run = document.createElement('span');
	run.textContent = 'Hello ';
	editor.append(run);
	document.body.append(editor);
	const text = run.firstChild as Text;
	const range = document.createRange();
	range.setStart(text, text.length);
	range.collapse(true);
	document.getSelection()?.removeAllRanges();
	document.getSelection()?.addRange(range);
	const canvas = {
		editCancelled: false,
		emitTextFormat: vi.fn(),
	} as unknown as SlideCanvasComponent;
	const onKeydown = (event: KeyboardEvent) =>
		SlideCanvasComponent.prototype.onEditorKeydown.call(canvas, event);
	editor.addEventListener('keydown', onKeydown);
	const blur = vi.spyOn(editor, 'blur');
	cleanup = () => {
		editor.removeEventListener('keydown', onKeydown);
		blur.mockRestore();
		document.getSelection()?.removeAllRanges();
		editor.remove();
	};
	return {
		editor,
		blur,
		press: (init, composing = false) => {
			const event = new KeyboardEvent('keydown', { bubbles: true, cancelable: true, ...init });
			Object.defineProperty(event, 'isComposing', { value: composing });
			editor.dispatchEvent(event);
			return event;
		},
	};
}

describe('slide canvas inline editor chords', () => {
	it.each([
		['Ctrl+Shift+Enter', { key: 'Enter', ctrlKey: true, shiftKey: true }],
		['Cmd+Shift+Enter', { key: 'Enter', metaKey: true, shiftKey: true }],
		['Ctrl+Enter', { key: 'Enter', ctrlKey: true }],
		['Cmd+Enter', { key: 'Enter', metaKey: true }],
		['Escape', { key: 'Escape' }],
	])('%s leaves the edit instead of inserting a line break', (_name, init) => {
		const { editor, blur, press } = harness();
		const event = press(init);
		expect(event.defaultPrevented).toBeTruthy();
		expect(blur).toHaveBeenCalledOnce();
		expect(editor.textContent).toBe('Hello ');
	});

	it('only cancels Ctrl/Cmd+D, keeping the edit open', () => {
		const { blur, press } = harness();
		const event = press({ key: 'd', ctrlKey: true });
		expect(event.defaultPrevented).toBeTruthy();
		expect(blur).not.toHaveBeenCalled();
	});

	it.each([
		['Enter', { key: 'Enter' }],
		['Ctrl+Enter', { key: 'Enter', ctrlKey: true }],
		['Ctrl+Shift+Enter', { key: 'Enter', ctrlKey: true, shiftKey: true }],
		['Escape', { key: 'Escape' }],
	])('leaves %s to an open IME composition', (_name, init) => {
		const { editor, blur, press } = harness();
		const event = press(init, true);
		expect(event.defaultPrevented).toBeFalsy();
		expect(blur).not.toHaveBeenCalled();
		expect(editor.textContent).toBe('Hello ');
	});

	it.each([
		['Enter', { key: 'Enter' }],
		['Shift+Enter', { key: 'Enter', shiftKey: true }],
	])('lets %s through to the browser after dropping the trailing space', (_name, init) => {
		const { editor, blur, press } = harness();
		const event = press(init);
		expect(event.defaultPrevented).toBeFalsy();
		expect(blur).not.toHaveBeenCalled();
		expect(editor.textContent).toBe('Hello');
	});
});
