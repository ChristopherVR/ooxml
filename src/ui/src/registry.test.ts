// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest';
import { OFFICE_UI_TAGS } from './index';
import { defineTitleBar } from './chrome/title-bar';
import { html } from 'lit';
import { OfficeElement } from './base';
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

	it('moves properties set before the definition onto the class accessors, in order', () => {
		const tag = unique('early');
		const early = document.createElement(tag) as HTMLElement & Record<string, unknown>;
		early.first = 1;
		early.second = 2;
		early.unrelated = 'kept';
		document.body.append(early);
		const writes: string[] = [];
		class Early extends HTMLElement {
			connected = false;
			connectedCallback() {
				this.connected = true;
			}
			get first() {
				return 10;
			}
			set first(value: number) {
				writes.push(`first=${value} connected=${this.connected}`);
			}
			get second() {
				return 20;
			}
			set second(value: number) {
				writes.push(`second=${value}`);
			}
		}
		defineElement(customElements, tag, Early);
		expect(writes).toEqual(['first=1 connected=true', 'second=2']);
		expect(early.first).toBe(10);
		expect(Object.hasOwn(early, 'first')).toBe(false);
		expect(early.unrelated).toBe('kept');
	});

	it('applies the remaining properties when one setter refuses its value', async () => {
		const tag = unique('refuse');
		const early = document.createElement(tag) as HTMLElement & Record<string, unknown>;
		early.bad = 1;
		early.good = 2;
		document.body.append(early);
		const reported = new Promise((resolve) => {
			const microtask = vi.spyOn(globalThis, 'queueMicrotask').mockImplementationOnce((task) => {
				microtask.mockRestore();
				try {
					task();
				} catch (error) {
					resolve(error);
				}
			});
		});
		let good: unknown;
		defineElement(
			customElements,
			tag,
			class extends HTMLElement {
				set bad(_value: number) {
					throw new Error('refused');
				}
				set good(value: number) {
					good = value;
				}
			},
		);
		expect(good).toBe(2);
		expect(await reported).toMatchObject({ message: 'refused' });
	});

	it('connects from the class defaults, not from the early value', () => {
		const tag = unique('defaults');
		const early = document.createElement(tag) as HTMLElement & Record<string, unknown>;
		early.items = undefined;
		document.body.append(early);
		const seen: unknown[] = [];
		let items: unknown = ['default'];
		defineElement(
			customElements,
			tag,
			class extends HTMLElement {
				connectedCallback() {
					seen.push(this.items);
				}
				get items() {
					return items;
				}
				set items(value: unknown) {
					items = value;
				}
			},
		);
		expect(seen).toEqual([['default']]);
		expect(items).toBeUndefined();
	});

	it('keeps an early value away from a render made in the constructor', () => {
		const tag = unique('constructor-render');
		const early = document.createElement(tag) as HTMLElement & { items: string[] };
		early.items = ['early'];
		document.body.append(early);
		class Rendering extends OfficeElement {
			#items: string[] = [];
			rendered: number[] = [];
			constructor() {
				super();
				void this.shadowRoot;
			}
			get items() {
				return this.#items;
			}
			set items(value: string[]) {
				this.#items = value;
				this.requestUpdate();
			}
			override render() {
				this.rendered.push(this.#items.length);
				return html`${this.items.join()}`;
			}
		}
		defineElement(customElements, tag, Rendering);
		const upgraded = early as unknown as Rendering;
		expect(upgraded.rendered[0]).toBe(0);
		expect(upgraded.shadowRoot!.textContent).toContain('early');
	});

	it('leaves an element created after the definition alone', () => {
		const tag = unique('late');
		let writes = 0;
		class Late extends HTMLElement {
			field = 'own';
			set value(_next: string) {
				writes++;
			}
		}
		defineElement(customElements, tag, Late);
		const late = document.createElement(tag) as Late;
		document.body.append(late);
		expect(writes).toBe(0);
		expect(late.field).toBe('own');
	});
});
