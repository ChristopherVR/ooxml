import { describe, expect, it } from 'vitest';
import fixture from './__fixtures__/lookup-native.json';
import { calc } from './test-helpers.js';

describe('native Excel mixed-type lookup corpus', () => {
	it.each(fixture.cases)('$scenario: $formula', ({ formula, cells, type, value }) => {
		expect(calc(formula, cells)).toEqual(type === 'error' ? { error: value } : value);
	});
});
