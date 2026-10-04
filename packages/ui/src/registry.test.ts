import { describe, expect, it } from 'vitest';
import { OFFICE_UI_TAGS } from './index.js';
import { defineTitleBar } from './chrome/title-bar.js';
import { definer } from './registry.js';

/** A registry stand-in that records the order tags are defined in. */
function fakeRegistry() {
	const defined = new Map<string, CustomElementConstructor>();
	const registry = {
		defined,
		get: (tag: string) => defined.get(tag),
		define: (tag: string, ctor: CustomElementConstructor) => void defined.set(tag, ctor),
	};
	return registry as typeof registry & CustomElementRegistry;
}

describe('definer', () => {
	it('defines the elements a composite renders before the composite, once', () => {
		const registry = fakeRegistry();
		const make = () => class extends HTMLElement {};
		const inner = definer('office-ui-test-inner', make);
		const outer = definer('office-ui-test-outer', make, [inner]);
		outer(registry);
		outer(registry);
		expect([...registry.defined.keys()]).toEqual(['office-ui-test-inner', 'office-ui-test-outer']);
	});

	it('gives the title bar its search field and switch', () => {
		const registry = fakeRegistry();
		defineTitleBar(registry);
		expect([...registry.defined.keys()].sort()).toEqual([
			'office-ui-search',
			'office-ui-switch',
			'office-ui-title-bar',
		]);
		expect(OFFICE_UI_TAGS).toContain('office-ui-title-bar');
	});
});
