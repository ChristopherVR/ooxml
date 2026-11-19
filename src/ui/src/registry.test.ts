// @vitest-environment jsdom
import { afterEach, describe, expect, it } from 'vitest';
import { OFFICE_UI_TAGS } from './index';
import { defineTitleBar } from './chrome/title-bar';
import { CONTRACT_REVISION, defineElement, definer } from './registry';

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

describe('defineElement', () => {
	const unique = (name: string) => `office-ui-test-${name}-${Math.random().toString(36).slice(2)}`;
	const contract = (ctor: CustomElementConstructor | undefined) =>
		(ctor as unknown as Record<symbol, number> | undefined)?.[
			Symbol.for('office-ui.web-control-contract')
		];

	afterEach(() => document.body.replaceChildren());

	it('stamps the class with the contract revision', () => {
		const tag = unique('stamp');
		defineElement(customElements, tag, class extends HTMLElement {});
		expect(contract(customElements.get(tag))).toBe(CONTRACT_REVISION);
	});
});
