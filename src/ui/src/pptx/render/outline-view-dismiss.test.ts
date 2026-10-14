// @vitest-environment jsdom
import { describe, expect, it, vi } from 'vitest';

import { isOutlineViewDismissKey, listenForOutlineViewDismiss } from './outline-view-dismiss';

describe('isOutlineViewDismissKey', () => {
	it('is bare Escape', () => {
		expect(isOutlineViewDismissKey({ key: 'Escape' })).toBe(true);
	});

	it('leaves chords, other keys and an IME Escape alone', () => {
		expect(isOutlineViewDismissKey({ key: 'Escape', ctrlKey: true })).toBe(false);
		expect(isOutlineViewDismissKey({ key: 'Escape', metaKey: true })).toBe(false);
		expect(isOutlineViewDismissKey({ key: 'Escape', altKey: true })).toBe(false);
		expect(isOutlineViewDismissKey({ key: 'Escape', shiftKey: true })).toBe(false);
		expect(isOutlineViewDismissKey({ key: 'Escape', isComposing: true })).toBe(false);
		expect(isOutlineViewDismissKey({ key: 'Enter' })).toBe(false);
	});
});

describe('listenForOutlineViewDismiss', () => {
	it('closes on Escape from anywhere in the document and stops the key there', () => {
		const onDismiss = vi.fn();
		const detach = listenForOutlineViewDismiss(document, onDismiss);
		const behind = vi.fn();
		document.body.addEventListener('keydown', behind);
		const button = document.createElement('button');
		document.body.append(button);
		try {
			const event = new KeyboardEvent('keydown', {
				key: 'Escape',
				bubbles: true,
				cancelable: true,
			});
			button.dispatchEvent(event);
			expect(onDismiss).toHaveBeenCalledOnce();
			expect(event.defaultPrevented).toBe(true);
			// The editor underneath must not also run its own Escape.
			expect(behind).not.toHaveBeenCalled();

			button.dispatchEvent(new KeyboardEvent('keydown', { key: 'a', bubbles: true }));
			expect(onDismiss).toHaveBeenCalledOnce();
			expect(behind).toHaveBeenCalledOnce();
		} finally {
			detach();
			document.body.removeEventListener('keydown', behind);
			button.remove();
		}
	});

	it('stops listening once detached', () => {
		const onDismiss = vi.fn();
		listenForOutlineViewDismiss(document, onDismiss)();
		document.body.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
		expect(onDismiss).not.toHaveBeenCalled();
	});
});
