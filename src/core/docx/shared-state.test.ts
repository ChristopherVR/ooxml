import { describe, expect, it } from 'vitest';
import { sharedWeakMap } from './shared-state';

describe('sharedWeakMap', () => {
	it('gives every bundle copy the same registry, by name', () => {
		const name = `test.shared-state.${Math.random()}`;
		const first = sharedWeakMap<object, number>(name);
		const key = {};
		first.set(key, 1);
		// A second bundle copy of this module asks by the same name and sees the same entries.
		expect(sharedWeakMap<object, number>(name).get(key)).toBe(1);
		expect(sharedWeakMap<object, number>(`${name}.other`).get(key)).toBeUndefined();
		const store = globalThis as Record<symbol, unknown>;
		const symbols = [name, `${name}.other`].map((item) => Symbol.for(`ooxml-core.${item}`));
		expect(store[symbols[0]!]).toBe(first);
		expect(Object.keys(globalThis).some((item) => item.includes(name))).toBe(false);
		for (const symbol of symbols) delete store[symbol];
		expect(store[symbols[0]!]).toBeUndefined();
	});
});
