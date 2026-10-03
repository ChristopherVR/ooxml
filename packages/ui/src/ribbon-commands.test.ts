import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest';
import { registerIcon } from './icons.js';
import { registerOfficeUi } from './index.js';

beforeAll(() => registerOfficeUi());
afterEach(() => {
	document.body.replaceChildren();
	vi.restoreAllMocks();
});

const listen = (host: Element, type: string) => {
	const spy = vi.fn();
	host.addEventListener(type, (event) => spy((event as CustomEvent).detail));
	return spy;
};

describe('office-ui-button ribbon features', () => {
	function button(attributes: Record<string, string>) {
		const el = document.createElement('office-ui-button');
		for (const [name, value] of Object.entries(attributes)) el.setAttribute(name, value);
		document.body.append(el);
		return { el, inner: el.shadowRoot!.querySelector('button')! };
	}

	it('draws a badge and a trailing caret only when asked', () => {
		const { el, inner } = button({ label: 'Comments', command: 'comments', badge: '3', caret: '' });
		const badge = el.shadowRoot!.querySelector<HTMLElement>('.badge')!;
		expect(badge.hidden).toBe(false);
		expect(badge.textContent).toBe('3');
		expect(inner.querySelector('.label .caret')!.hasAttribute('hidden')).toBe(false);
		el.setAttribute('caret', 'false');
		el.removeAttribute('badge');
		expect(inner.querySelector('.caret')!.hasAttribute('hidden')).toBe(true);
		expect(badge.hidden).toBe(true);
	});

	it('keeps Enter and Space away from host key handlers', () => {
		const { inner } = button({ label: 'New', command: 'new' });
		const host = vi.fn();
		document.body.addEventListener('keydown', host);
		inner.dispatchEvent(
			new KeyboardEvent('keydown', { key: 'Enter', bubbles: true, composed: true }),
		);
		inner.dispatchEvent(new KeyboardEvent('keydown', { key: 'a', bubbles: true, composed: true }));
		expect(host).toHaveBeenCalledOnce();
	});

	it('lets a product subclass keep its id attribute, event, glyph names and caret rule', () => {
		const Base = customElements.get('office-ui-button') as CustomElementConstructor;
		registerIcon('legacy:shape', 'M3 3h14v14H3Z');
		class Legacy extends Base {
			static requestEvent = 'command-request';
			static idAttribute = 'data-ribbon-control';
			static detailKey = 'id';
			iconName(): string | null {
				return `legacy:${this.getAttribute('icon')}`;
			}
			showsCaret(): boolean {
				return this.getAttribute('data-ribbon-control') === 'shapes';
			}
		}
		if (!customElements.get('legacy-ribbon-command'))
			customElements.define('legacy-ribbon-command', Legacy);
		const el = document.createElement('legacy-ribbon-command');
		el.setAttribute('label', 'Shapes');
		el.setAttribute('icon', 'shape');
		el.setAttribute('data-ribbon-control', 'shapes');
		document.body.append(el);
		const root = el.shadowRoot!;
		expect(root.querySelector('svg.glyph')!.hasAttribute('data-painted')).toBe(true);
		expect(root.querySelector('.caret')!.hasAttribute('hidden')).toBe(false);
		const spy = listen(el, 'command-request');
		root.querySelector('button')!.click();
		expect(spy).toHaveBeenCalledWith({ id: 'shapes' });
	});
});

describe('office-ui-ribbon-group collapse and layout', () => {
	function group(children = '') {
		const el = document.createElement('office-ui-ribbon-group');
		el.setAttribute('label', 'Arrange');
		el.innerHTML = children;
		document.body.append(el);
		return el;
	}

	it('shows a launcher for an empty launcher attribute and names it', () => {
		const el = group();
		el.setAttribute('launcher', 'arrange-dialog');
		el.setAttribute('launcher-label', 'More arrange options');
		const launcher = el.shadowRoot!.querySelector<HTMLButtonElement>('.launcher')!;
		expect(launcher.hidden).toBe(false);
		expect(launcher.getAttribute('aria-label')).toBe('More arrange options');
		const spy = listen(el, 'office-command');
		launcher.click();
		expect(spy).toHaveBeenCalledWith({ command: 'arrange-dialog' });
	});

	it('collapses to a face that asks the overflow controller to open it', () => {
		const el = group('<office-ui-button label="Align"></office-ui-button>');
		el.setAttribute('icon', 'alignObjects');
		el.setAttribute('data-collapsed', '');
		const face = el.shadowRoot!.querySelector<HTMLButtonElement>('.face')!;
		expect(face.getAttribute('aria-label')).toBe('Arrange');
		expect(face.getAttribute('aria-expanded')).toBe('false');
		const spy = listen(el, 'office-ribbon-collapse-toggle');
		face.click();
		expect(spy).toHaveBeenCalledOnce();
		el.setAttribute('data-open', '');
		expect(face.getAttribute('aria-expanded')).toBe('true');
	});

	it('marks one-line rows compact and stacks only drop-down galleries', () => {
		vi.spyOn(Element.prototype, 'getBoundingClientRect').mockReturnValue(
			DOMRect.fromRect({ width: 40, height: 24 }),
		);
		const compact = group('<office-ui-button label="Bold"></office-ui-button>');
		expect(compact.hasAttribute('data-compact-row')).toBe(true);
		const stack = group('<div mode="dropdown"></div><div mode="dropdown"></div>');
		expect(stack.hasAttribute('data-stack')).toBe(true);
		const mixed = group('<div mode="dropdown"></div><office-ui-button></office-ui-button>');
		expect(mixed.hasAttribute('data-stack')).toBe(false);
	});
});
