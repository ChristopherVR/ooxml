import { describe, expect, it } from 'vitest';

import { insertedLayoutCompletion } from './inserted-slide-layout';

describe('insertedLayoutCompletion', () => {
	const placed = { id: 'new-slide', elements: [] };

	it('replaces the draft while it is untouched', () => {
		expect(insertedLayoutCompletion(placed, placed)).toBe('replace');
	});

	it('keeps an edit made while the layout loaded, and takes only the artwork', () => {
		expect(insertedLayoutCompletion({ ...placed, elements: [{ id: 'shape' }] }, placed)).toBe(
			'artwork',
		);
	});

	it('does nothing once the slide is gone or another one is at its index', () => {
		expect(insertedLayoutCompletion(undefined, placed)).toBe('none');
		expect(insertedLayoutCompletion({ id: 'other' }, placed)).toBe('none');
	});
});
