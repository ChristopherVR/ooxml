import { afterEach, expect, it, vi } from 'vitest';
import { wireHandleEvents } from './viewer-handle-events';
afterEach(() => document.body.replaceChildren());
it('owns cancelled clicks and resets stale ownership on a new press', () => {
	const viewport = document.createElement('div');
	document.body.append(viewport);
	let pointer: number | undefined;
	const cancel = vi.fn(() => {
		pointer = undefined;
	});
	const dispose = wireHandleEvents(viewport, {
		pointer: () => pointer,
		start: () => {
			pointer = 7;
		},
		move: vi.fn(),
		finish: async () => {
			pointer = undefined;
		},
		cancel,
	});
	const selectionClick = vi.fn();
	viewport.addEventListener('click', selectionClick);
	const dispatchPointer = (type: string) => {
		const event = new Event(type, { bubbles: true });
		Object.defineProperty(event, 'pointerId', { value: 7 });
		viewport.dispatchEvent(event);
	};
	dispatchPointer('pointerdown');
	viewport.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
	expect(cancel).toHaveBeenCalledOnce();
	dispatchPointer('pointerup');
	viewport.click();
	expect(selectionClick).not.toHaveBeenCalled();
	viewport.click();
	expect(selectionClick).toHaveBeenCalledOnce();
	dispatchPointer('pointerdown');
	viewport.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
	// Cancellation may not generate a click. A new press resets stale suppression.
	dispatchPointer('pointerdown');
	viewport.click();
	expect(selectionClick).toHaveBeenCalledTimes(2);
	dispose();
	dispatchPointer('pointerdown');
	expect(pointer).toBeUndefined();
});
