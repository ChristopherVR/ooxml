import { describe, expect, it } from 'vitest';

import { createCustomShow, generateCustomShowId } from './custom-shows-helpers';

describe('generateCustomShowId', () => {
	it('returns an unsigned integer string (p:custShow/@id is an xsd:unsignedInt)', () => {
		expect(generateCustomShowId()).toMatch(/^\d+$/u);
	});

	it('returns the next id above the ones in use', () => {
		expect(generateCustomShowId([{ id: '0' }, { id: '4' }, { id: 'legacy' }])).toBe('5');
	});
});

describe('createCustomShow', () => {
	it('creates a show with the given name and slide ids', () => {
		const show = createCustomShow('My Show', ['slide-1', 'slide-2']);
		expect(show.name).toBe('My Show');
		expect(show.slideIds).toStrictEqual(['slide-1', 'slide-2']);
	});

	it('trims whitespace from the name', () => {
		const show = createCustomShow('  Trimmed  ', []);
		expect(show.name).toBe('Trimmed');
	});

	it('assigns a unique numeric id among the existing shows', () => {
		const show1 = createCustomShow('A', []);
		const show2 = createCustomShow('B', [], [show1]);
		expect(show1.id).toMatch(/^\d+$/u);
		expect(show2.id).not.toBe(show1.id);
	});

	it('creates a copy of the slideIds array', () => {
		const ids = ['slide-1', 'slide-2'];
		const show = createCustomShow('Test', ids);
		ids.push('slide-3');
		expect(show.slideIds).toHaveLength(2);
	});

	it('accepts readonly slideIds', () => {
		const ids: readonly string[] = ['slide-a'];
		const show = createCustomShow('Readonly Test', ids);
		expect(show.slideIds).toStrictEqual(['slide-a']);
	});
});
