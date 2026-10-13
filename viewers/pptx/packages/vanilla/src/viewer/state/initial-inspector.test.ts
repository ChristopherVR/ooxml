import { describe, expect, it } from 'vitest';

import { createInitialViewerState, createStore } from './index';

describe('initial inspector visibility (#61)', () => {
	it('starts closed on request and can be opened through the store', () => {
		const store = createStore(createInitialViewerState(false));
		expect(store.get().inspectorOpen).toBe(false);
		store.set({ inspectorOpen: true });
		expect(store.get().inspectorOpen).toBe(true);
		expect(createInitialViewerState().inspectorOpen).toBe(true);
	});
});
