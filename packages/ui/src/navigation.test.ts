import { afterEach, beforeAll, describe, expect, it } from 'vitest';
import { registerOfficeUi, type OfficeTab } from './index.js';

beforeAll(() => registerOfficeUi());
afterEach(() => document.body.replaceChildren());

const make = <T extends HTMLElement>(html: string): T => {
	document.body.innerHTML = html;
	return document.body.firstElementChild as T;
};
type Slider = HTMLElement & { value: number; disabled: boolean };
type Strip = HTMLElement & { tabs: OfficeTab[]; selected: string };

describe('office-ui-zoom-slider', () => {
	it('reflects value into the range and percentage without emitting', () => {
		const el = make<Slider>('<office-ui-zoom-slider value="67"></office-ui-zoom-slider>');
		const events: string[] = [];
		el.addEventListener('input', () => events.push('input'));
		const root = el.shadowRoot!;
		expect(root.querySelector('input')!.value).toBe('67');
		expect(root.querySelector('output')!.value).toBe('67%');
		el.value = 900;
		expect(el.value).toBe(400);
		expect(root.querySelector<HTMLButtonElement>('[aria-label="Zoom in"]')!.disabled).toBe(true);
		expect(events).toEqual([]);
		expect(el.getAttribute('role')).toBe('group');
	});

	it('steps to the next multiple like Office and emits input then change', () => {
		const el = make<Slider>('<office-ui-zoom-slider value="67"></office-ui-zoom-slider>');
		const events: string[] = [];
		for (const type of ['input', 'change'])
			el.addEventListener(type, () => events.push(`${type}:${el.value}`));
		el.shadowRoot!.querySelector<HTMLButtonElement>('[aria-label="Zoom in"]')!.click();
		expect(events).toEqual(['input:70', 'change:70']);
		el.shadowRoot!.querySelector<HTMLButtonElement>('[aria-label="Zoom out"]')!.click();
		el.shadowRoot!.querySelector<HTMLButtonElement>('[aria-label="Zoom out"]')!.click();
		expect(el.value).toBe(50);
	});

	it('emits slider input, hides fit unless requested and blocks disabled use', () => {
		const el = make<Slider>(
			'<office-ui-zoom-slider value="100" fit="Fit page to current window"></office-ui-zoom-slider>',
		);
		const seen: unknown[] = [];
		document.body.addEventListener('office-command', (e) => seen.push((e as CustomEvent).detail));
		const range = el.shadowRoot!.querySelector('input')!;
		range.value = '150';
		range.dispatchEvent(new Event('input'));
		expect(el.value).toBe(150);
		const fit = el.shadowRoot!.querySelector<HTMLButtonElement>('.fit')!;
		expect(fit.getAttribute('aria-label')).toBe('Fit page to current window');
		fit.click();
		expect(seen).toEqual([{ command: 'zoom-fit' }]);
		el.disabled = true;
		el.shadowRoot!.querySelector<HTMLButtonElement>('[aria-label="Zoom out"]')!.click();
		expect(el.value).toBe(150);
		expect(range.disabled).toBe(true);
	});
});

describe('office-ui-tab-strip', () => {
	const tabs = [
		{ id: 'p1', label: 'Process' },
		{ id: 'p2', label: '<img src=x onerror=alert(1)>', title: 'Background' },
		{ id: 'p3', label: 'Notes' },
	];

	it('renders labels as text with tab semantics and a roving tabindex', () => {
		const el = make<Strip>('<office-ui-tab-strip label="Pages"></office-ui-tab-strip>');
		el.tabs = tabs;
		el.selected = 'p1';
		const root = el.shadowRoot!;
		expect(root.querySelector('[role="tablist"]')!.getAttribute('aria-label')).toBe('Pages');
		const items = [...root.querySelectorAll<HTMLButtonElement>('[role="tab"]')];
		expect(items.map((tab) => tab.textContent)).toEqual(tabs.map((tab) => tab.label));
		expect(root.querySelector('img')).toBeNull();
		expect(items[1]!.title).toBe('Background');
		expect(items.map((tab) => tab.tabIndex)).toEqual([0, -1, -1]);
		expect(items[0]!.getAttribute('aria-selected')).toBe('true');
		expect(root.querySelector<HTMLButtonElement>('[aria-label="Previous"]')!.disabled).toBe(true);
	});

	it('selects by click, step buttons and keys, emitting only for user changes', () => {
		const el = make<Strip>('<office-ui-tab-strip></office-ui-tab-strip>');
		const seen: string[] = [];
		el.addEventListener('office-tab-select', (e) => seen.push((e as CustomEvent).detail.id));
		el.tabs = tabs;
		el.selected = 'p1';
		const root = el.shadowRoot!;
		root.querySelectorAll<HTMLButtonElement>('[role="tab"]')[2]!.click();
		root.querySelector<HTMLButtonElement>('[aria-label="Previous"]')!.click();
		const list = root.querySelector('[role="tablist"]')!;
		list.dispatchEvent(new KeyboardEvent('keydown', { key: 'Home', bubbles: true }));
		list.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowRight', bubbles: true }));
		expect(seen).toEqual(['p3', 'p2', 'p1', 'p2']);
		expect(root.activeElement?.textContent).toBe(tabs[1]!.label);
		el.selected = 'p3';
		expect(seen).toHaveLength(4);
	});

	it('lets the host cancel a selection and ignores input when disabled', () => {
		const el = make<Strip>('<office-ui-tab-strip></office-ui-tab-strip>');
		el.tabs = tabs;
		el.selected = 'p1';
		el.addEventListener('office-tab-select', (e) => e.preventDefault(), { once: true });
		el.shadowRoot!.querySelectorAll<HTMLButtonElement>('[role="tab"]')[1]!.click();
		expect(el.selected).toBe('p1');
		el.setAttribute('disabled', '');
		el.shadowRoot!.querySelector<HTMLButtonElement>('[aria-label="Next"]')!.click();
		expect(el.selected).toBe('p1');
		el.setAttribute('next-label', 'Next page');
		expect(el.shadowRoot!.querySelector('[aria-label="Next page"]')).not.toBeNull();
	});
});

describe('office-ui-button keyboard shortcuts', () => {
	it('forwards keyshortcuts to aria-keyshortcuts', () => {
		const el = make(
			'<office-ui-button label="Undo" icon="undo" keyshortcuts="Control+Z"></office-ui-button>',
		);
		const button = el.shadowRoot!.querySelector('button')!;
		expect(button.getAttribute('aria-keyshortcuts')).toBe('Control+Z');
		expect(el.shadowRoot!.querySelector('svg path')).not.toBeNull();
		el.removeAttribute('keyshortcuts');
		expect(button.hasAttribute('aria-keyshortcuts')).toBe(false);
	});
});

describe('office-ui-status-bar end slot', () => {
	it('assigns slot="end" children to the trailing slot', () => {
		const el = make(
			'<office-ui-status-bar><span id="a">Page 1</span><office-ui-zoom-slider id="z" slot="end"></office-ui-zoom-slider></office-ui-status-bar>',
		);
		const end = el.shadowRoot!.querySelector<HTMLSlotElement>('slot[name="end"]')!;
		expect(end.assignedElements().map((node) => node.id)).toEqual(['z']);
		const main = el.shadowRoot!.querySelector<HTMLSlotElement>('slot:not([name])')!;
		expect(main.assignedElements().map((node) => node.id)).toEqual(['a']);
	});
});
