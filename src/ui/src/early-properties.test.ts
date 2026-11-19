// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest';
import { html } from 'lit';
import { OfficeElement } from './base';
import { defineElement } from './registry';

const unique = (name: string) => `office-ui-early-${name}-${Math.random().toString(36).slice(2)}`;
const flush = () => new Promise<void>((resolve) => setTimeout(resolve, 0));

/** A plain element whose `label` property mirrors its `label` attribute, like the editors'. */
function plainLabel(log: string[] = []) {
	return class extends HTMLElement {
		static observedAttributes = ['label'];
		#label = 'default';
		connectedCallback() {
			log.push(`connected label=${this.#label}`);
		}
		attributeChangedCallback(_name: string, _old: string | null, value: string | null) {
			this.label = value ?? 'default';
		}
		get label() {
			return this.#label;
		}
		set label(value: string) {
			this.#label = value;
			log.push(`label=${value}`);
		}
	};
}

/** The same on the shared base class. */
class OfficeLabel extends OfficeElement {
	#label = 'default';
	get label() {
		return this.#label;
	}
	set label(value: string) {
		this.#label = value;
	}
}

type Labelled = HTMLElement & { label: string };
const early = (tag: string, label: string, attribute?: string): Labelled => {
	const element = document.createElement(tag) as Labelled;
	if (attribute !== undefined) element.setAttribute('label', attribute);
	element.label = label;
	return element;
};

afterEach(() => document.body.replaceChildren());

describe('a property set before the definition loads', () => {
	it('is applied, in assignment order, before the element first connects', () => {
		const tag = unique('order');
		const element = document.createElement(tag) as HTMLElement & Record<string, unknown>;
		element.first = 1;
		element.second = 2;
		element.unrelated = 'kept';
		document.body.append(element);
		const log: string[] = [];
		defineElement(
			customElements,
			tag,
			class extends HTMLElement {
				connectedCallback() {
					log.push('connected');
				}
				set first(value: number) {
					log.push(`first=${value}`);
				}
				set second(value: number) {
					log.push(`second=${value}`);
				}
			},
		);
		expect(log).toEqual(['first=1', 'second=2', 'connected']);
		expect(Object.hasOwn(element, 'first')).toBe(false);
		expect(element.unrelated).toBe('kept');
	});

	it('wins over the attribute the element had when it upgraded', () => {
		const tag = unique('attribute');
		const log: string[] = [];
		const element = early(tag, 'property', 'attribute');
		document.body.append(element);
		defineElement(customElements, tag, plainLabel(log));
		expect(element.label).toBe('property');
		expect(log.at(-1)).toBe('connected label=property');
	});

	it('loses to a property written after a detached upgrade', async () => {
		for (const [name, make] of [
			['plain', () => plainLabel()],
			['office', () => class extends OfficeLabel {}],
		] as const) {
			const tag = unique(`later-${name}`);
			const element = early(tag, 'early');
			defineElement(customElements, tag, make());
			customElements.upgrade(element);
			element.label = 'later';
			await flush();
			expect(element.label, name).toBe('later');
			document.body.append(element);
			expect(element.label, name).toBe('later');
		}
	});

	it('loses to an attribute written after a detached upgrade', async () => {
		const tag = unique('later-attribute');
		const element = early(tag, 'early', 'first');
		defineElement(customElements, tag, plainLabel());
		customElements.upgrade(element);
		element.setAttribute('label', 'later');
		await flush();
		expect(element.label).toBe('later');
	});

	it('is applied once the script that upgraded a detached element finishes', async () => {
		const tag = unique('detached');
		const element = early(tag, 'early');
		defineElement(customElements, tag, class extends OfficeLabel {});
		customElements.upgrade(element);
		await flush();
		expect(Object.hasOwn(element, 'label')).toBe(false);
		expect(element.label).toBe('early');
	});

	it('is kept from a render the constructor makes', () => {
		const tag = unique('constructor-render');
		const element = document.createElement(tag) as HTMLElement & { items: string[] };
		element.items = ['early'];
		document.body.append(element);
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
		const upgraded = element as unknown as Rendering;
		expect(upgraded.rendered[0]).toBe(0);
		expect(upgraded.shadowRoot!.textContent).toContain('early');
	});

	it('does not stop the others when its setter refuses it', () => {
		const tag = unique('refuse');
		const element = document.createElement(tag) as HTMLElement & Record<string, unknown>;
		element.bad = 1;
		element.good = 2;
		document.body.append(element);
		const tasks: VoidFunction[] = [];
		const microtask = vi.spyOn(globalThis, 'queueMicrotask').mockImplementation((task) => {
			tasks.push(task);
		});
		let good: unknown;
		try {
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
		} finally {
			microtask.mockRestore();
		}
		expect(good).toBe(2);
		const reported: unknown[] = [];
		for (const task of tasks) {
			try {
				task();
			} catch (error) {
				reported.push(error);
			}
		}
		expect(reported).toMatchObject([{ message: 'refused' }]);
	});
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
