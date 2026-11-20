// @vitest-environment happy-dom
import { afterEach, describe, expect, it, vi } from 'vitest';

import { listenForPopupMenuDismiss } from './popup-menu-dismiss';

afterEach(() => {
	document.body.replaceChildren();
});

const press = (target: EventTarget, init: KeyboardEventInit): KeyboardEvent => {
	const event = new KeyboardEvent('keydown', { bubbles: true, cancelable: true, ...init });
	target.dispatchEvent(event);
	return event;
};

describe('listenForPopupMenuDismiss', () => {
	it('closes on Escape wherever focus is and gives focus back to the trigger', () => {
		const trigger = document.createElement('button');
		const item = document.createElement('button');
		document.body.append(trigger, item);
		const onDismiss = vi.fn();
		const editorEscape = vi.fn();
		document.body.addEventListener('keydown', editorEscape);
		const detach = listenForPopupMenuDismiss(document, { onDismiss, trigger });
		item.focus();
		const event = press(item, { key: 'Escape' });
		expect(event.defaultPrevented).toBe(true);
		expect(onDismiss).toHaveBeenCalledOnce();
		expect(document.activeElement).toBe(trigger);
		// Stopped in the capture phase, before the editor's own Escape clears the selection.
		expect(editorEscape).not.toHaveBeenCalled();
		detach();
		document.body.removeEventListener('keydown', editorEscape);
	});

	it('ignores chords, IME Escape and every key after detaching', () => {
		const onDismiss = vi.fn();
		const detach = listenForPopupMenuDismiss(document, { onDismiss });
		press(document.body, { key: 'Escape', ctrlKey: true });
		press(document.body, { key: 'Enter' });
		const composing = new KeyboardEvent('keydown', { key: 'Escape', bubbles: true });
		Object.defineProperty(composing, 'isComposing', { value: true });
		document.body.dispatchEvent(composing);
		expect(onDismiss).not.toHaveBeenCalled();
		detach();
		expect(press(document.body, { key: 'Escape' }).defaultPrevented).toBe(false);
		expect(onDismiss).not.toHaveBeenCalled();
	});
});
