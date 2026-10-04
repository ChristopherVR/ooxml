import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest';
import type { OfficeTitleBarState } from './controls.js';
import { registerOfficeUi } from './index.js';

// Adapted from pptx-viewer's title-bar tests (`packages/shared/src/web-components/title-bar.test.ts`).
beforeAll(() => registerOfficeUi());
afterEach(() => document.body.replaceChildren());

type TitleBar = HTMLElement & {
	state: OfficeTitleBarState;
	placement: 'titleBar' | 'belowRibbon';
	searchField: HTMLElement & { value: string };
};

const STATE: OfficeTitleBarState = {
	appMark: 'V',
	fileName: 'Plan.vsdx',
	status: 'Saved to this PC',
	autosave: { enabled: true, label: 'AutoSave', stateLabel: 'On', toggleLabel: 'Toggle AutoSave' },
	quickAccess: {
		label: 'Quick Access Toolbar',
		items: [
			{ id: 'save', icon: 'save', label: 'Save', title: 'Save (Ctrl+S)' },
			{ id: 'undo', icon: 'undo', label: 'Undo', disabled: true },
			{ id: 'redo', icon: 'redo', label: 'Redo' },
		],
	},
	search: {
		placeholder: 'Search',
		label: 'Search commands',
		heading: 'Commands',
		empty: 'No commands',
		commands: [
			{ id: 'bold', label: 'Bold', category: 'home' },
			{ id: 'border', label: 'Borders' },
			{ id: 'paste', label: 'Paste' },
		],
		content: (query) => `Find "${query}"`,
	},
};

function make(state: OfficeTitleBarState = STATE, placement?: string): TitleBar {
	const bar = document.createElement('office-ui-title-bar') as TitleBar;
	if (placement) bar.setAttribute('placement', placement);
	bar.state = state;
	document.body.append(bar);
	return bar;
}
const $ = <T extends Element = HTMLElement>(bar: TitleBar, selector: string) =>
	bar.shadowRoot!.querySelector<T>(selector)!;
const listen = (host: HTMLElement, type: string) => {
	const spy = vi.fn();
	host.addEventListener(type, (event) => spy((event as CustomEvent).detail));
	return spy;
};
const type = (bar: TitleBar, text: string) => {
	bar.searchField.value = text;
	bar.searchField.dispatchEvent(new Event('input', { bubbles: true }));
};

describe('office-ui-title-bar', () => {
	it('renders the translated state and hides absent parts', () => {
		const bar = make();
		expect($(bar, '.mark').textContent).toBe('V');
		expect($(bar, '.name').textContent).toBe('Plan.vsdx');
		expect($(bar, '.status').textContent).toBe('Saved to this PC');
		expect($(bar, '.autosave').hidden).toBe(false);
		bar.state = { fileName: 'Bare' };
		for (const part of ['.mark', '.autosave', '.qat', '.status', '.dot', '.search'])
			expect($(bar, part).hidden, part).toBe(true);
	});

	it('keys Quick Access buttons by id with labels, ScreenTips and disabled state', () => {
		const bar = make();
		const buttons = [...bar.shadowRoot!.querySelectorAll<HTMLButtonElement>('.qat button')];
		expect(buttons.map((b) => b.dataset.command)).toEqual(['save', 'undo', 'redo']);
		expect(buttons[0]!.getAttribute('aria-label')).toBe('Save');
		expect(buttons[0]!.title).toBe('Save (Ctrl+S)');
		expect(buttons[1]!.hasAttribute('title')).toBe(false);
		expect(buttons[1]!.disabled).toBe(true);
		expect(buttons[0]!.querySelector('path')).not.toBeNull();
		// Reordering keeps the same nodes.
		bar.state = {
			...STATE,
			quickAccess: {
				label: 'QAT',
				items: [...STATE.quickAccess!.items].reverse(),
				showLabels: true,
			},
		};
		const after = [...bar.shadowRoot!.querySelectorAll<HTMLButtonElement>('.qat button')];
		expect(after[0]).toBe(buttons[2]);
		expect(after[0]!.querySelector('small')!.textContent).toBe('Redo');
	});

	it('emits office-command for a Quick Access button and roves with the arrows', () => {
		const bar = make();
		const command = listen(bar, 'office-command');
		const [save, , redo] = bar.shadowRoot!.querySelectorAll<HTMLButtonElement>('.qat button');
		redo!.click();
		expect(command).toHaveBeenCalledWith({ command: 'redo' });
		expect(save!.tabIndex).toBe(0);
		save!.focus();
		save!.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowRight', bubbles: true }));
		// Disabled Undo is skipped.
		expect(bar.shadowRoot!.activeElement).toBe(redo);
		expect(redo!.tabIndex).toBe(0);
	});

	it('keeps AutoSave controlled and requests a toggle', () => {
		const bar = make();
		const toggle = listen(bar, 'office-autosave-toggle');
		const sw = $<HTMLElement & { checked: boolean }>(bar, '.switch');
		expect(sw.checked).toBe(true);
		sw.checked = false;
		sw.dispatchEvent(new Event('change', { bubbles: true }));
		expect(toggle).toHaveBeenCalledTimes(1);
		expect(sw.checked).toBe(true);
	});

	it('searches commands locally and commits one request', () => {
		const bar = make();
		const search = listen(bar, 'office-command-search');
		type(bar, 'bol');
		const options = [...bar.shadowRoot!.querySelectorAll('[role="option"]')];
		expect(options.map((o) => o.firstChild!.textContent)).toEqual(['Bold']);
		expect($(bar, '.results .content').textContent).toBe('Find "bol"');
		bar.searchField.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
		expect(search).toHaveBeenCalledWith({ query: 'bol', command: 'bold' });
		expect($(bar, '.results').hidden).toBe(true);
		type(bar, 'zzz');
		expect($(bar, '.empty').textContent).toBe('No commands');
		$(bar, '.results .content').dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));
		expect(search).toHaveBeenLastCalledWith({ query: 'zzz' });
	});

	it('prefers the host matcher and caps the list', () => {
		const many = Array.from({ length: 12 }, (_, i) => ({ id: `c${i}`, label: `C${i}` }));
		const bar = make({ ...STATE, search: { ...STATE.search!, match: () => many } });
		type(bar, 'anything');
		expect(bar.shadowRoot!.querySelectorAll('[role="option"]')).toHaveLength(8);
	});

	it('renders only the strip below the ribbon and hides itself when empty', () => {
		const bar = make(STATE, 'belowRibbon');
		expect(bar.shadowRoot!.querySelector('.search')).toBeNull();
		expect(bar.shadowRoot!.querySelector('.qat')).not.toBeNull();
		expect(bar.hasAttribute('data-empty')).toBe(false);
		bar.state = { fileName: '' };
		expect(bar.hasAttribute('data-empty')).toBe(true);
		bar.placement = 'titleBar';
		expect(bar.shadowRoot!.querySelector('.search')).not.toBeNull();
		expect(bar.hasAttribute('data-empty')).toBe(false);
	});

	it('lets a product subclass rename events and route buttons', () => {
		const Base = customElements.get('office-ui-title-bar')!;
		class Product extends (Base as unknown as { new (): TitleBar }) {
			static autosaveEvent = 'toggle-autosave';
			static searchEvent = 'command-search';
			activate(id: string): void {
				this.dispatchEvent(new CustomEvent(id, { bubbles: true }));
			}
			rendered(): void {
				this.toggleAttribute('data-product-bar', true);
			}
		}
		customElements.define('product-title-bar', Product as unknown as CustomElementConstructor);
		const bar = document.createElement('product-title-bar') as TitleBar;
		bar.state = STATE;
		document.body.append(bar);
		expect(bar.hasAttribute('data-product-bar')).toBe(true);
		const save = listen(bar, 'save');
		$<HTMLButtonElement>(bar, '.qat button').click();
		expect(save).toHaveBeenCalled();
		const toggle = listen(bar, 'toggle-autosave');
		$(bar, '.switch').dispatchEvent(new Event('change'));
		expect(toggle).toHaveBeenCalled();
	});
});
