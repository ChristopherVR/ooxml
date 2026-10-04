import { afterEach, beforeAll, describe, expect, it } from 'vitest';
import { registerOfficeUi } from './index.js';

// Products (pptx-viewer's `pptx-ui-*` tags) register subclasses of these elements and rely on more
// than the documented hooks. These tests pin the behaviour the Lit port has to keep.
beforeAll(() => registerOfficeUi());
afterEach(() => document.body.replaceChildren());

type Ctor = new () => HTMLElement;
const base = (tag: string) => customElements.get(tag) as unknown as Ctor;
let tags = 0;
const define = (Class: Ctor): string => {
	const tag = `product-element-${++tags}`;
	customElements.define(tag, Class as unknown as CustomElementConstructor);
	return tag;
};

describe('subclassing the elements', () => {
	it('lets a subclass override an accessor and keep its own private model', () => {
		interface Toasts extends HTMLElement {
			get state(): unknown;
			set state(value: unknown);
		}
		class Product extends (base('office-ui-toasts') as unknown as new () => Toasts) {
			#model: { toasts: unknown[] } = { toasts: [] };
			override get state() {
				return this.#model;
			}
			override set state(value: unknown) {
				const next = value as { toasts: unknown[] };
				this.#model = next;
				super.state = {
					toasts: next.toasts.map((id) => ({ id, severity: 'info', message: `m:${String(id)}` })),
				};
			}
		}
		const el = document.createElement(define(Product)) as Product;
		el.state = { toasts: ['a', 'b'] };
		document.body.append(el);
		expect(el.shadowRoot!.querySelectorAll('.toast')).toHaveLength(2);
		expect(el.shadowRoot!.querySelector('.message')!.textContent).toBe('m:a');
	});

	it('creates without adding attributes and lets a constructor reach the shadow root', () => {
		let root: ShadowRoot | null | undefined;
		class Product extends base('office-ui-checkbox') {
			constructor() {
				super();
				root = this.shadowRoot;
			}
		}
		const tag = define(Product);
		// `document.createElement` throws if a constructor adds attributes.
		expect(() => document.createElement(tag)).not.toThrow();
		expect(root).not.toBeNull();
	});

	it('renders a connected subclass that overrides connectedCallback without super', async () => {
		class Product extends base('office-ui-toasts') {
			connectedCallback(): void {
				/* The product never calls super. */
			}
		}
		const el = document.createElement(define(Product)) as Product & { state: unknown };
		el.state = { toasts: [{ id: '1', severity: 'info', message: 'hello' }] };
		document.body.append(el);
		await Promise.resolve();
		expect(el.shadowRoot!.querySelector('.message')?.textContent).toBe('hello');
	});

	it('renders the title bar when a constructor reads searchField and attribute changes bypass super', () => {
		type Bar = HTMLElement & { state: unknown; searchField: HTMLElement; placement: string };
		class Product extends (base('office-ui-title-bar') as unknown as new () => Bar) {
			constructor() {
				super();
				this.searchField.setAttribute('data-product', '');
			}
			attributeChangedCallback(): void {
				this.state = this.state;
			}
		}
		const tag = define(Product);
		const el = document.createElement(tag) as Bar;
		el.setAttribute('placement', 'belowRibbon');
		el.state = {
			fileName: 'a.pptx',
			quickAccess: { label: 'Quick', items: [{ id: 'save', icon: 'save', label: 'Save' }] },
		};
		document.body.append(el);
		expect(el.placement).toBe('belowRibbon');
		expect(el.shadowRoot!.querySelectorAll('.qat button')).toHaveLength(1);
		expect(el.shadowRoot!.querySelector('.name')).toBeNull();
	});

	it('writes properties that product hooks read as attributes through to the attribute', () => {
		const button = document.createElement('office-ui-button') as HTMLElement & {
			icon: string;
			command: string;
		};
		button.icon = 'copy';
		button.command = 'copy';
		expect(button.getAttribute('icon')).toBe('copy');
		expect(button.getAttribute('command')).toBe('copy');
		const group = document.createElement('office-ui-ribbon-group') as HTMLElement & {
			launcher: string;
		};
		group.launcher = 'font';
		expect(group.getAttribute('launcher')).toBe('font');
	});

	it('keeps a controlled toggle controlled until the host commits', () => {
		const toggle = document.createElement('office-ui-ribbon-toggle') as HTMLElement & {
			checked: boolean;
		};
		toggle.setAttribute('command', 'grid');
		toggle.setAttribute('label', 'Grid');
		document.body.append(toggle);
		const seen: unknown[] = [];
		toggle.addEventListener('office-toggle', (e) => seen.push((e as CustomEvent).detail));
		toggle.shadowRoot!.querySelector('label')!.click();
		expect(seen).toEqual([{ command: 'grid', checked: true }]);
		expect(toggle.checked).toBe(false);
		expect(
			toggle.shadowRoot!.querySelector<HTMLElement & { checked: boolean }>('office-ui-checkbox')!
				.checked,
		).toBe(false);
	});

	it('reads exact text from labels, options and strip items', () => {
		document.body.innerHTML =
			'<office-ui-button label="Hide" icon="check" caret></office-ui-button>' +
			'<office-ui-status-bar></office-ui-status-bar>';
		const button = document.body.firstElementChild!;
		expect(button.shadowRoot!.querySelector('button')!.textContent).toBe('Hide');
		const bar = document.querySelector('office-ui-status-bar') as HTMLElement & { state: unknown };
		bar.state = { toggles: [{ id: 'notes', icon: 'message', label: 'Notes', text: 'Notes' }] };
		expect(bar.shadowRoot!.querySelector('.toggle')!.textContent).toBe('Notes');
	});
});
