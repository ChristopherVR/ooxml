import { describe, expect, it } from 'vitest';

import { ChromeUiState } from './chrome-ui.svelte';

describe('initial inspector visibility (#61)', () => {
	it('starts closed on request and keeps the panel toggle working', () => {
		const state = new ChromeUiState(false);
		expect(state.inspectorOpen).toBe(false);
		state.toggleInspector();
		expect(state.inspectorOpen).toBe(true);
		expect(new ChromeUiState().inspectorOpen).toBe(true);
	});
});
