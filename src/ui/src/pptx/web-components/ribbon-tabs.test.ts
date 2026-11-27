// @vitest-environment jsdom
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest';

import { buildRibbonTabRow } from '../render';
import type { RibbonTabRowInput, RibbonTabRowState, RibbonTabSelectDetail } from '../render';
import { registerPptxWebControls } from './index';

beforeAll(registerPptxWebControls);
afterEach(() => document.body.replaceChildren());

const input: RibbonTabRowInput = {
	translate: (key) => key.replace(/^pptx\.ribbon\.tab\./, ''),
	active: 'home',
};

type TabRow = HTMLElement & {
	state: RibbonTabRowState;
	tabButtons(): HTMLButtonElement[];
	setInstalledAddInTabs(tabs: readonly { id: string; label: string; groups: [] }[]): void;
};

function mount(overrides: Partial<RibbonTabRowInput> = {}): TabRow {
	const host = document.createElement('pptx-ui-ribbon-tabs') as TabRow;
	document.body.append(host);
	host.state = buildRibbonTabRow({ ...input, ...overrides });
	return host;
}

describe('buildRibbonTabRow', () => {
	it('lists the fixed tabs, then the add-in tabs, then the contextual tabs', () => {
		const state = buildRibbonTabRow({
			...input,
			addInTabs: [{ id: 'tools', label: 'Tools', groups: [] }],
			contextualTabs: ['shapeFormat'],
			active: 'shapeFormat',
		});
		expect(state.items.map((tab) => `${tab.kind}:${tab.id}`)).toEqual([
			'file:file',
			'tab:home',
			'tab:insert',
			'tab:draw',
			'tab:design',
			'tab:transitions',
			'tab:animations',
			'tab:slideShow',
			'tab:record',
			'tab:review',
			'tab:view',
			'tab:help',
			'addIn:tools',
			'contextual:shapeFormat',
		]);
		expect(state.items.at(-1)).toMatchObject({ contextual: true, label: 'shapeFormat' });
		expect(state.selected).toBe('shapeFormat');
	});

	it('hides host and user hidden tabs, but the user cannot hide File', () => {
		const state = buildRibbonTabRow({
			...input,
			hiddenActions: ['draw'],
			hiddenTabIds: ['file', 'review'],
		});
		const ids = state.items.map((tab) => tab.id);
		expect(ids).toContain('file');
		expect(ids).not.toContain('draw');
		expect(ids).not.toContain('review');
		expect(buildRibbonTabRow({ ...input, hiddenActions: ['file'] }).items[0]!.id).toBe('home');
	});

	it('draws ScreenTips through the resolver and selects nothing for an unknown tab', () => {
		const off = buildRibbonTabRow({ ...input, screenTip: () => undefined, active: 'missing' });
		expect(off.items[1]!.title).toBeUndefined();
		expect(off.selected).toBe('');
		expect(buildRibbonTabRow(input).items[1]!.title).toBe('home');
	});
});

describe('pptx-ui-ribbon-tabs', () => {
	it('draws the state on the shared tab row, with the markers the bindings used', () => {
		const host = mount({
			addInTabs: [{ id: 'tools', label: 'Tools', groups: [] }],
			contextualTabs: ['pictureFormat'],
		});
		const tabs = host.tabButtons();
		expect(tabs.map((tab) => tab.textContent)).toContain('Tools');
		const home = tabs.find((tab) => tab.dataset.tab === 'home')!;
		expect(home.getAttribute('aria-selected')).toBe('true');
		expect(tabs.find((tab) => tab.dataset.tab === 'tools')!.dataset.ribbonAddInTab).toBe('tools');
		const picture = tabs.find((tab) => tab.dataset.tab === 'pictureFormat')!;
		expect(picture.dataset.ribbonContextualTab).toBe('pictureFormat');
		expect(picture.hasAttribute('data-contextual')).toBe(true);
		expect(home.hasAttribute('data-ribbon-contextual-tab')).toBe(false);
	});

	it('reports a chosen tab with its kind and leaves the selection to the binding', () => {
		const host = mount({ contextualTabs: ['chartDesign'] });
		const picked: RibbonTabSelectDetail[] = [];
		host.addEventListener('tab-select', (event) =>
			picked.push((event as CustomEvent<RibbonTabSelectDetail>).detail),
		);
		const byId = (id: string) => host.tabButtons().find((tab) => tab.dataset.tab === id)!;
		byId('chartDesign').click();
		byId('file').click();
		expect(picked).toEqual([
			{ tab: 'chartDesign', kind: 'contextual' },
			{ tab: 'file', kind: 'file' },
		]);
		expect(byId('home').getAttribute('aria-selected')).toBe('true');
	});

	it('marks the installed panel as the sole active ribbon panel across bindings', async () => {
		const host = mount();
		host.setInstalledAddInTabs([{ id: 'installed-add-in.notes', label: 'Notes', groups: [] }]);
		const tab = await vi.waitFor(() => {
			const notes = host.tabButtons().find((button) => button.textContent === 'Notes');
			if (!notes) throw new Error('Missing installed tab');
			return notes;
		});
		tab.click();
		expect(host).toHaveProperty('selected', 'installed-add-in.notes');
		expect(host.hasAttribute('data-manifest-tab-active')).toBe(true);
		expect(host.querySelector('[data-add-in]')).toHaveProperty('hidden', false);
		host.state = buildRibbonTabRow(input);
		expect(host).toHaveProperty('selected', 'installed-add-in.notes');
		host.state = buildRibbonTabRow({
			...input,
			active: 'pictureFormat',
			contextualTabs: ['pictureFormat'],
		});
		expect(host).toHaveProperty('selected', 'pictureFormat');
		expect(host.hasAttribute('data-manifest-tab-active')).toBe(false);
		expect(host.querySelector('[data-add-in]')).toHaveProperty('hidden', true);
		host
			.tabButtons()
			.find((button) => button.textContent === 'Notes')
			?.click();
		host
			.tabButtons()
			.find((button) => button.dataset.tab === 'home')
			?.click();
		expect(host.hasAttribute('data-manifest-tab-active')).toBe(false);
		expect(host.querySelector('[data-add-in]')).toHaveProperty('hidden', true);
	});

	it('asks to collapse on a double click on a tab other than File', () => {
		const host = mount();
		let toggles = 0;
		host.addEventListener('collapse-toggle', () => toggles++);
		const byId = (id: string) => host.tabButtons().find((tab) => tab.dataset.tab === id)!;
		byId('view').dispatchEvent(new MouseEvent('dblclick', { bubbles: true, composed: true }));
		byId('file').dispatchEvent(new MouseEvent('dblclick', { bubbles: true, composed: true }));
		expect(toggles).toBe(1);
	});
});
