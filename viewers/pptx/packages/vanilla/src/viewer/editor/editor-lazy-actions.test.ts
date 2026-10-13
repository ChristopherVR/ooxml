import { describe, expect, it } from 'vitest';

import { createLazyActions } from './editor-lazy-actions';

interface Fake {
	greet(name: string): string;
}

interface Nested {
	sections: { addSection(name: string): string; count: number };
	label: string;
	who(): string;
}

describe('createLazyActions', () => {
	it('forwards calls to whatever getActions() currently returns', () => {
		let current: Fake = { greet: (name) => `v1:${name}` };
		const lazy = createLazyActions(() => current);
		expect(lazy.greet('a')).toBe('v1:a');

		current = { greet: (name) => `v2:${name}` };
		expect(lazy.greet('b')).toBe('v2:b');
	});

	it('reuses the same wrapper function across calls to the same method', () => {
		const target: Fake = { greet: (name) => name };
		const lazy = createLazyActions(() => target);
		const first = lazy.greet;
		const second = lazy.greet;
		expect(first).toBe(second);
	});

	it('forwards calls into nested action groups at call time', () => {
		let count = 0;
		const make = (tag: string): Nested => ({
			sections: {
				count: 0,
				addSection(name) {
					count += 1;
					return `${tag}:${name}`;
				},
			},
			label: tag,
			who() {
				return this.label;
			},
		});
		let current = make('v1');
		const lazy = createLazyActions(() => current);
		expect(lazy.sections.addSection('a')).toBe('v1:a');

		current = make('v2');
		expect(lazy.sections.addSection('b')).toBe('v2:b');
		expect(count).toBe(2);
		// Methods still run with the current bag as `this`.
		expect(lazy.who()).toBe('v2');
		expect(lazy.sections).toBe(lazy.sections);
	});
});
