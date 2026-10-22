import { describe, expect, it } from 'vitest';

import { autosaveStatusKind } from './viewer-event-payloads';

describe('autosaveStatusKind', () => {
	it('passes the in-flight lifecycle through, from a string or a state object', () => {
		for (const state of ['saving', 'saved', 'error'] as const) {
			expect(autosaveStatusKind(state)).toBe(state);
			expect(autosaveStatusKind({ state })).toBe(state);
		}
	});

	it('reports a disabled or idle engine as idle', () => {
		expect(autosaveStatusKind('disabled')).toBe('idle');
		expect(autosaveStatusKind({ state: 'disabled' })).toBe('idle');
		expect(autosaveStatusKind({ state: 'idle' })).toBe('idle');
	});
});
