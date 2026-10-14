/**
 * The canvas's plain-text inline editor (a `<textarea>`) honours the shared
 * inline-editor chords (`mapInlineEditorChord`) before its own Enter rules.
 *
 * It used to check only the mapper's `suppress` result, so Ctrl/Cmd+Shift+Enter
 * (which the mapper resolves to `commit`) fell through to the Shift+Enter
 * branch and inserted a line break. It also committed on Enter while an IME
 * composition was open, when that Enter belongs to the IME.
 *
 * The handler is called on the prototype with only the members it touches, as
 * `canvas-element-locks.test.ts` does, so no injected services are needed.
 */
import { afterEach, describe, expect, it, vi } from 'vitest';

import { SlideCanvasComponent } from './slide-canvas.component';

interface EditorHarness {
	editor: HTMLTextAreaElement;
	blur: ReturnType<typeof vi.spyOn>;
	trim: ReturnType<typeof vi.fn>;
	press: (init: KeyboardEventInit, composing?: boolean) => KeyboardEvent;
}

let cleanup: (() => void) | null = null;

afterEach(() => {
	cleanup?.();
	cleanup = null;
});

function harness(): EditorHarness {
	const editor = document.createElement('textarea');
	editor.value = 'Hello';
	document.body.append(editor);
	const trim = vi.fn();
	const canvas = {
		editCancelled: false,
		emitTextFormat: vi.fn(),
		trimTrailingSpaceBeforeCaret: trim,
	} as unknown as SlideCanvasComponent;
	const onKeydown = (event: KeyboardEvent) =>
		SlideCanvasComponent.prototype.onEditorKeydown.call(canvas, event);
	editor.addEventListener('keydown', onKeydown);
	const blur = vi.spyOn(editor, 'blur');
	cleanup = () => {
		editor.removeEventListener('keydown', onKeydown);
		blur.mockRestore();
		editor.remove();
	};
	return {
		editor,
		blur,
		trim,
		press: (init, composing = false) => {
			const event = new KeyboardEvent('keydown', { bubbles: true, cancelable: true, ...init });
			Object.defineProperty(event, 'isComposing', { value: composing });
			editor.dispatchEvent(event);
			return event;
		},
	};
}

describe('slide canvas textarea editor chords', () => {
	it.each([
		['Ctrl+Shift+Enter', { key: 'Enter', ctrlKey: true, shiftKey: true }],
		['Cmd+Shift+Enter', { key: 'Enter', metaKey: true, shiftKey: true }],
		['Ctrl+Enter', { key: 'Enter', ctrlKey: true }],
		['Cmd+Enter', { key: 'Enter', metaKey: true }],
	])('%s leaves the edit instead of inserting a line break', (_name, init) => {
		const { blur, trim, press } = harness();
		const event = press(init);
		expect(event.defaultPrevented).toBeTruthy();
		expect(blur).toHaveBeenCalledOnce();
		expect(trim).not.toHaveBeenCalled();
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
		const { blur, trim, press } = harness();
		const event = press(init, true);
		expect(event.defaultPrevented).toBeFalsy();
		expect(blur).not.toHaveBeenCalled();
		expect(trim).not.toHaveBeenCalled();
	});

	it('still lets plain Shift+Enter insert a line break', () => {
		const { blur, trim, press } = harness();
		const event = press({ key: 'Enter', shiftKey: true });
		expect(event.defaultPrevented).toBeFalsy();
		expect(blur).not.toHaveBeenCalled();
		expect(trim).toHaveBeenCalledOnce();
	});

	it('still commits on plain Enter', () => {
		const { blur, press } = harness();
		const event = press({ key: 'Enter' });
		expect(event.defaultPrevented).toBeTruthy();
		expect(blur).toHaveBeenCalledOnce();
	});
});
