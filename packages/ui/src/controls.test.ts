import { registerOfficeUi } from './index.js';

beforeAll(() => registerOfficeUi());
afterEach(() => document.body.replaceChildren());

const make = <T extends HTMLElement>(html: string): T => {
	document.body.innerHTML = html;
	return document.body.firstElementChild as T;
};
const key = (el: Element, k: string, init: KeyboardEventInit = {}) =>
	el.dispatchEvent(
		new KeyboardEvent('keydown', { key: k, bubbles: true, cancelable: true, ...init }),
	);

describe('office-ui-button', () => {
	it('renders label, icon and aria state from attributes', () => {
		const el = make(
			'<office-ui-button label="Bold" icon="check" command="bold" pressed="true"></office-ui-button>',
		);
		const button = el.shadowRoot!.querySelector('button')!;
		expect(button.textContent).toBe('Bold');
		expect(button.getAttribute('aria-pressed')).toBe('true');
		expect(el.shadowRoot!.querySelector('svg path')).not.toBeNull();
		el.setAttribute('pressed', 'false');
		expect(button.getAttribute('aria-pressed')).toBe('false');
		el.removeAttribute('pressed');
		expect(button.hasAttribute('aria-pressed')).toBe(false);
	});

	it('emits one bubbling, composed office-command and nothing when disabled', () => {
		const el = make<HTMLElement & { disabled: boolean }>(
			'<office-ui-button label="Go" command="go"></office-ui-button>',
		);
		const seen: unknown[] = [];
		document.body.addEventListener('office-command', (e) => seen.push((e as CustomEvent).detail));
		el.shadowRoot!.querySelector('button')!.click();
		expect(seen).toEqual([{ command: 'go' }]);
		el.disabled = true;
		el.shadowRoot!.querySelector('button')!.click();
		expect(seen).toHaveLength(1);
		expect(el.shadowRoot!.querySelector('button')!.disabled).toBe(true);
	});

	it('icon-only buttons keep an accessible name', () => {
		const el = make(
			'<office-ui-button label="Copy" icon="copy" icon-only command="c"></office-ui-button>',
		);
		const button = el.shadowRoot!.querySelector('button')!;
		expect(button.getAttribute('aria-label')).toBe('Copy');
		expect(button.title).toBe('Copy');
	});
});

describe('office-ui-checkbox and switch', () => {
	it('exposes role, aria-checked and toggles on click and Space with input then change', () => {
		const el = make<HTMLElement & { checked: boolean }>(
			'<office-ui-checkbox aria-label="Opt"></office-ui-checkbox>',
		);
		const order: string[] = [];
		el.addEventListener('input', () => order.push('input'));
		el.addEventListener('change', () => order.push('change'));
		expect(el.getAttribute('role')).toBe('checkbox');
		expect(el.getAttribute('aria-checked')).toBe('false');
		expect(el.tabIndex).toBe(0);
		el.click();
		expect(el.checked).toBe(true);
		expect(el.getAttribute('aria-checked')).toBe('true');
		key(el, ' ');
		expect(el.checked).toBe(false);
		expect(order).toEqual(['input', 'change', 'input', 'change']);
	});

	it('property writes are silent and disabled blocks activation', () => {
		const el = make<HTMLElement & { checked: boolean; disabled: boolean }>(
			'<office-ui-checkbox></office-ui-checkbox>',
		);
		const spy = vi.fn();
		el.addEventListener('change', spy);
		el.checked = true;
		expect(spy).not.toHaveBeenCalled();
		el.disabled = true;
		el.click();
		key(el, ' ');
		expect(el.checked).toBe(true);
		expect(el.tabIndex).toBe(-1);
		expect(el.getAttribute('aria-disabled')).toBe('true');
	});

	it('switch uses role=switch and also toggles on Enter', () => {
		const el = make<HTMLElement & { checked: boolean }>('<office-ui-switch></office-ui-switch>');
		expect(el.getAttribute('role')).toBe('switch');
		key(el, 'Enter');
		expect(el.checked).toBe(true);
		const box = make<HTMLElement & { checked: boolean }>(
			'<office-ui-checkbox></office-ui-checkbox>',
		);
		key(box, 'Enter');
		expect(box.checked).toBe(false);
	});
});

describe('office-ui-select', () => {
	type Sel = HTMLElement & { options: unknown[]; value: string; selectedIndex: number };
	const options = [
		{ value: 'a', label: 'Alpha' },
		{ value: 'b', label: 'Beta', disabled: true },
		{ value: 'c', label: 'Gamma' },
	];
	const setup = (): { el: Sel; trigger: HTMLButtonElement; list: HTMLElement } => {
		const el = make<Sel>('<office-ui-select aria-label="Greek"></office-ui-select>');
		el.options = options;
		const root = el.shadowRoot!;
		return { el, trigger: root.querySelector('button')!, list: root.querySelector('ul')! };
	};

	it('has combobox/listbox/option semantics', () => {
		const { el, trigger, list } = setup();
		el.value = 'c';
		expect(trigger.getAttribute('role')).toBe('combobox');
		expect(trigger.getAttribute('aria-label')).toBe('Greek');
		expect(trigger.getAttribute('aria-expanded')).toBe('false');
		expect(list.getAttribute('role')).toBe('listbox');
		const items = [...list.querySelectorAll('li')];
		expect(items.map((i) => i.getAttribute('role'))).toEqual(['option', 'option', 'option']);
		expect(items.map((i) => i.getAttribute('aria-selected'))).toEqual(['false', 'false', 'true']);
		expect(items[1]!.getAttribute('aria-disabled')).toBe('true');
		expect(trigger.textContent).toContain('Gamma');
		expect(el.selectedIndex).toBe(2);
	});

	it('navigates with arrows skipping disabled options, chooses with Enter, emits input then change', () => {
		const { el, trigger } = setup();
		const events: string[] = [];
		el.addEventListener('input', () => events.push('input'));
		el.addEventListener('change', () => events.push('change'));
		key(trigger, 'ArrowDown');
		expect(trigger.getAttribute('aria-expanded')).toBe('true');
		key(trigger, 'ArrowDown');
		expect(trigger.getAttribute('aria-activedescendant')).toMatch(/-2$/);
		key(trigger, 'Enter');
		expect(el.value).toBe('c');
		expect(trigger.getAttribute('aria-expanded')).toBe('false');
		expect(events).toEqual(['input', 'change']);
	});

	it('Escape closes without changing, Home/End jump, typeahead picks', () => {
		const { el, trigger } = setup();
		key(trigger, 'ArrowDown');
		key(trigger, 'End');
		expect(trigger.getAttribute('aria-activedescendant')).toMatch(/-2$/);
		key(trigger, 'Home');
		expect(trigger.getAttribute('aria-activedescendant')).toMatch(/-0$/);
		key(trigger, 'Escape');
		expect(trigger.getAttribute('aria-expanded')).toBe('false');
		expect(el.value).toBe('');
		key(trigger, 'g');
		expect(el.value).toBe('c');
	});

	it('property writes are silent; disabled does not open', () => {
		const { el, trigger } = setup();
		const spy = vi.fn();
		el.addEventListener('change', spy);
		el.selectedIndex = 0;
		expect(el.value).toBe('a');
		expect(spy).not.toHaveBeenCalled();
		el.setAttribute('disabled', '');
		trigger.click();
		expect(trigger.getAttribute('aria-expanded')).toBe('false');
	});
});

describe('office-ui-ribbon-group and toolbar', () => {
	it('group is a labelled role=group', () => {
		const el = make('<office-ui-ribbon-group label="Clipboard"></office-ui-ribbon-group>');
		expect(el.getAttribute('role')).toBe('group');
		expect(el.getAttribute('aria-label')).toBe('Clipboard');
		expect(el.shadowRoot!.querySelector('.caption')!.textContent).toBe('Clipboard');
		el.setAttribute('label', 'Font');
		expect(el.getAttribute('aria-label')).toBe('Font');
	});

	it('toolbar moves focus with arrows, Home and End, skipping disabled items', () => {
		const el = make(
			'<office-ui-toolbar aria-label="Tools"><office-ui-button label="a" command="a"></office-ui-button>' +
				'<office-ui-button label="b" command="b" disabled></office-ui-button>' +
				'<office-ui-ribbon-group label="g"><office-ui-button label="c" command="c"></office-ui-button></office-ui-ribbon-group></office-ui-toolbar>',
		);
		expect(el.getAttribute('role')).toBe('toolbar');
		const [a, , c] = [...el.querySelectorAll<HTMLElement>('office-ui-button')];
		const focused: Element[] = [];
		for (const b of [a!, c!]) b.focus = () => void focused.push(b);
		key(a!, 'ArrowRight');
		key(c!, 'ArrowRight');
		key(a!, 'End');
		key(c!, 'Home');
		key(a!, 'ArrowLeft');
		expect(focused).toEqual([c, a, c, a, c]);
	});
});

describe('office-ui-status-bar and item', () => {
	it('bar is a named group; items render label and value', () => {
		const bar = make(
			'<office-ui-status-bar><office-ui-status-item label="Page" value="2 of 5"></office-ui-status-item></office-ui-status-bar>',
		);
		expect(bar.getAttribute('role')).toBe('group');
		expect(bar.getAttribute('aria-label')).toBe('Status');
		const item = bar.querySelector('office-ui-status-item') as HTMLElement & { value: string };
		expect(item.shadowRoot!.textContent).toContain('Page');
		item.value = '3 of 5';
		expect(item.shadowRoot!.textContent).toContain('3 of 5');
	});

	it('interactive items are buttons that emit office-status-activate', () => {
		const item = make(
			'<office-ui-status-item id="zoom" label="Zoom" value="100%" interactive></office-ui-status-item>',
		);
		const seen: unknown[] = [];
		document.body.addEventListener('office-status-activate', (e) =>
			seen.push((e as CustomEvent).detail),
		);
		item.shadowRoot!.querySelector('button')!.click();
		expect(seen).toEqual([{ id: 'zoom' }]);
	});
});

describe('office-ui-icon', () => {
	it('is decorative unless labelled', () => {
		const el = make('<office-ui-icon name="check"></office-ui-icon>');
		expect(el.getAttribute('aria-hidden')).toBe('true');
		expect(el.shadowRoot!.querySelector('path')).not.toBeNull();
		el.setAttribute('label', 'Done');
		expect(el.getAttribute('role')).toBe('img');
		expect(el.getAttribute('aria-label')).toBe('Done');
		expect(el.hasAttribute('aria-hidden')).toBe(false);
	});
});
