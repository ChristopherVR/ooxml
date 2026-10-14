import { afterEach, expect, it, vi } from 'vitest';
import { trackPointerGesture } from './pointer-gesture';
afterEach(() => document.body.replaceChildren());
it.each(['pointercancel', 'lostpointercapture', 'Escape'])(
	'rolls back on %s and ignores later release',
	(reason) => {
		const owner = document.createElement('div');
		document.body.append(owner);
		const callbacks = { move: vi.fn(), commit: vi.fn(), cancel: vi.fn() };
		const dispose = trackPointerGesture(owner, 2, callbacks);
		const pointer = (type: string, id = 2) => {
			const event = new Event(type, { bubbles: true });
			Object.defineProperty(event, 'pointerId', { value: id });
			owner.dispatchEvent(event);
		};
		pointer('pointermove', 3);
		expect(callbacks.move).not.toHaveBeenCalled();
		pointer('pointermove');
		if (reason === 'Escape')
			document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape' }));
		else pointer(reason);
		pointer('pointerup');
		dispose();
		expect(callbacks.cancel).toHaveBeenCalledOnce();
		expect(callbacks.commit).not.toHaveBeenCalled();
	},
);
