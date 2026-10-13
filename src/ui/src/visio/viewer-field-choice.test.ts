import { expect, it } from 'vitest';
import { populateExistingFields } from './viewer-field-choice';
it('offers only core-proven existing fields, retaining insertion and local editing', () => {
	const select = document.createElement('select');
	const source = { start: 0, end: 1, cached: '2', index: '0', sourceKey: 'proof' };
	populateExistingFields(select, [
		source,
		{ ...source, index: '1', editable: false },
		{ ...source, index: '2', editable: true },
		{ ...source, index: '3', editable: true, sourceKey: undefined } as never,
	]);
	expect([...select.options].map((option) => option.value)).toEqual(['', '2']);
});
