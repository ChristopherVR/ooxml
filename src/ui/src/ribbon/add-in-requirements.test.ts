import { expect, it } from 'vitest';
import { addInCapabilities, isAddInSetSupported } from './add-in-requirements';

it('advertises the two text selection sets only when a selection adapter is available', () => {
	expect([...(addInCapabilities(true).sets ?? [])]).toEqual([
		['selection', '1.1'],
		['textcoercion', '1.1'],
	]);
	expect(addInCapabilities(false)).toEqual({});
	expect(isAddInSetSupported(true, 'tExtCoercion')).toBe(true);
	expect(isAddInSetSupported(true, 'Selection', '1.1.0')).toBe(true);
	expect(isAddInSetSupported(true, 'Selection', '1.2')).toBe(false);
	expect(isAddInSetSupported(true, 'WordApi')).toBe(false);
	expect(isAddInSetSupported(false, 'Selection')).toBe(false);
});
