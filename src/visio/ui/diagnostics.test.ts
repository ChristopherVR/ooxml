import { describe, it, expect } from 'vitest';
import { compatibilityNotes, compatibilityText } from './diagnostics.js';
describe('compatibility summaries', () => {
	it('groups repeated per-shape messages while keeping occurrence counts', () => {
		const notes = compatibilityNotes([
			{ code: 'x', severity: 'warning', message: 'Unsupported style', shapeId: '1' },
			{ code: 'x', severity: 'warning', message: 'Unsupported style', shapeId: '2' },
		]);
		expect(notes).toHaveLength(1);
		expect(compatibilityText(notes[0]!)).toBe('Unsupported style (2 occurrences)');
	});
	it('puts warnings before informational limitations', () => {
		const notes = compatibilityNotes(
			[{ code: 'cache', severity: 'info', message: 'Cached values' }],
			['Text approximation'],
		);
		expect(notes[0]!.message).toBe('Text approximation');
		expect(notes[1]!.message).toBe('Cached values');
	});
});
