import { describe, expect, it, vi } from 'vitest';

import { createControllerStore } from './controller';

describe('createControllerStore', () => {
	it('keeps the snapshot identity until the state changes', () => {
		const store = createControllerStore({ count: 0 });
		const first = store.getSnapshot();
		expect(store.getSnapshot()).toBe(first);
		store.update((s) => ({ count: s.count + 1 }));
		expect(store.getSnapshot()).not.toBe(first);
		expect(store.getSnapshot().count).toBe(1);
	});

	it('notifies subscribers on change only, and stops after unsubscribe', () => {
		const store = createControllerStore({ count: 0 });
		const listener = vi.fn();
		const unsubscribe = store.subscribe(listener);
		store.set(store.getSnapshot());
		expect(listener).not.toHaveBeenCalled();
		store.set({ count: 1 });
		expect(listener).toHaveBeenCalledTimes(1);
		unsubscribe();
		store.set({ count: 2 });
		expect(listener).toHaveBeenCalledTimes(1);
	});
});
