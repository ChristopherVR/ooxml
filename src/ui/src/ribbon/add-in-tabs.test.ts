import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest';
import { registerOfficeUi } from '../index';
import { RIBBON_ADD_IN_EVENT, syncRibbonAddIns, type RibbonAddInTab } from './add-in-tabs';

beforeAll(() => registerOfficeUi());
afterEach(() => document.body.replaceChildren());

type Ribbon = HTMLElement & { selected: string; updateComplete: Promise<unknown> };

function make(): Ribbon {
	const ribbon = document.createElement('office-ui-ribbon') as Ribbon;
	for (const [id, label] of [
		['home', 'Home'],
		['help', 'Help'],
	] as const) {
		const panel = document.createElement('div');
		panel.dataset.ribbonTab = id;
		panel.dataset.label = label;
		ribbon.append(panel);
	}
	document.body.append(ribbon);
	return ribbon;
}

const pdf = (run = vi.fn()): RibbonAddInTab => ({
	id: 'pdf',
	label: 'PDF',
	keytip: 'B',
	groups: [
		{
			label: 'Create PDF',
			commands: [
				{ id: 'create', label: 'Create PDF', icon: 'save', run },
				{ id: 'preferences', label: 'Preferences', size: 'small' },
				{ id: 'locked', label: 'Locked', size: 'small', disabled: true, run },
				{ id: 'a', label: 'A', size: 'small' },
				{ id: 'b', label: 'B', size: 'small' },
				{
					id: 'share',
					label: 'Share',
					items: [{ id: 'share-mail', label: 'By mail', run }],
				},
			],
		},
	],
});

const tabs = (ribbon: Ribbon) =>
	[...ribbon.shadowRoot!.querySelectorAll('[role="tab"]')].map((tab) => tab.textContent!.trim());
const click = (el: Element, command: string) =>
	el.dispatchEvent(
		new CustomEvent('office-command', { detail: { command }, bubbles: true, composed: true }),
	);

describe('ribbon add-in tabs', () => {
	it('adds a tab after the product tabs, with groups, columns of three and a drop-down', async () => {
		const ribbon = make();
		expect(syncRibbonAddIns(ribbon, [pdf()], { panelClass: 'panel' })).toEqual(['pdf']);
		await ribbon.updateComplete;
		expect(tabs(ribbon)).toEqual(['Home', 'Help', 'PDF']);
		// The product's selection stays.
		expect(ribbon.selected).toBe('home');
		const panel = ribbon.querySelector<HTMLElement>('[data-add-in]')!;
		expect(panel.className).toBe('panel');
		expect(panel.dataset.tabKeytip).toBe('B');
		const group = panel.querySelector('office-ui-ribbon-group')!;
		expect(group.getAttribute('label')).toBe('Create PDF');
		expect([...group.children].map((child) => child.localName)).toEqual([
			'office-ui-button',
			'office-ui-ribbon-stack',
			'office-ui-ribbon-stack',
			'office-ui-menu-button',
		]);
		expect(
			[...group.querySelectorAll('office-ui-ribbon-stack')].map((stack) => stack.children.length),
		).toEqual([3, 1]);
		expect(group.querySelector('[data-add-in-command="create"]')!.getAttribute('variant')).toBe(
			'stacked',
		);
		expect(group.querySelector('[data-add-in-command="locked"]')!.hasAttribute('disabled')).toBe(
			true,
		);
		expect(group.querySelector('office-ui-menu-item')!.getAttribute('label')).toBe('By mail');
	});

	it('runs the command, announces it past the ribbon and hides it from the product router', () => {
		const ribbon = make();
		const run = vi.fn();
		const product = vi.fn();
		const heard = vi.fn();
		ribbon.addEventListener('office-command', product);
		document.body.addEventListener(RIBBON_ADD_IN_EVENT, (event) =>
			heard((event as CustomEvent).detail),
		);
		syncRibbonAddIns(ribbon, [pdf(run)]);
		const panel = ribbon.querySelector<HTMLElement>('[data-add-in]')!;
		click(panel.querySelector('[data-add-in-command="create"]')!, 'create');
		expect(run).toHaveBeenCalledTimes(1);
		expect(heard).toHaveBeenLastCalledWith({ tab: 'pdf', command: 'create' });
		click(panel.querySelector('office-ui-menu-item')!, 'share-mail');
		expect(heard).toHaveBeenLastCalledWith({ tab: 'pdf', command: 'share-mail' });
		// Disabled and unknown commands do nothing.
		click(panel, 'locked');
		click(panel, 'missing');
		expect(run).toHaveBeenCalledTimes(2);
		expect(heard).toHaveBeenCalledTimes(2);
		expect(product).not.toHaveBeenCalled();
	});

	it('replaces the add-in tabs and never a product tab', async () => {
		const ribbon = make();
		const home = ribbon.querySelector('[data-ribbon-tab="home"]');
		syncRibbonAddIns(ribbon, [pdf()]);
		const other: RibbonAddInTab = { id: 'sign', label: 'Sign', groups: [] };
		expect(
			syncRibbonAddIns(ribbon, [
				other,
				{ ...other, label: 'Again' },
				{ id: 'home', label: 'Fake Home', groups: [] },
				{ id: '', label: 'Nameless', groups: [] },
			]),
		).toEqual(['sign']);
		await ribbon.updateComplete;
		expect(tabs(ribbon)).toEqual(['Home', 'Help', 'Sign']);
		expect(ribbon.querySelector('[data-ribbon-tab="home"]')).toBe(home);
		expect(syncRibbonAddIns(ribbon, [])).toEqual([]);
		expect(ribbon.querySelector('[data-add-in]')).toBeNull();
	});
});
