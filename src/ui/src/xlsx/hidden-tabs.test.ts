// @vitest-environment jsdom
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest';
import { EXCEL_RIBBON_TAB_IDS, commandsOnlyInTabs, ribbonTabs } from 'ooxml-core/xlsx/ui';
import type { EditorCore } from './editor-core';
import { defineXlsxEditor, type ExcelCustomization, type XlsxEditorElement } from './index';
import { resolveExcelCustomization } from './customization';

beforeAll(() => defineXlsxEditor());
afterEach(() => document.body.replaceChildren());

const NO_FILE: ExcelCustomization = { ribbon: { hiddenTabs: ['file'] } };

function mount(customization?: ExcelCustomization) {
	const element = document.createElement('xlsx-editor') as XlsxEditorElement;
	if (customization) element.customization = customization;
	document.body.append(element);
	element.newWorkbook();
	return element;
}
const core = (element: XlsxEditorElement) => (element as unknown as { core: EditorCore }).core;
type Ribbon = HTMLElement & {
	noFile: boolean;
	selected: string;
	fileButton(): HTMLButtonElement | null;
	updateComplete: Promise<unknown>;
};
const ribbon = (element: XlsxEditorElement) =>
	element.shadowRoot!.querySelector('office-ui-ribbon') as Ribbon;
const backstage = (element: XlsxEditorElement) =>
	element.shadowRoot!.querySelector('office-ui-backstage') as HTMLElement & { open: boolean };
/** Lets the queued render and the tab row's panel observer run. */
const settle = () => new Promise((resolve) => setTimeout(resolve));
const tabIds = async (element: XlsxEditorElement) => {
	await settle();
	await ribbon(element).updateComplete;
	return [...ribbon(element).shadowRoot!.querySelectorAll<HTMLElement>('[role="tab"]')].map(
		(tab) => tab.dataset.tab,
	);
};

describe('<xlsx-editor> ribbon.hiddenTabs', () => {
	it('names every native tab the ribbon builds', () => {
		mount();
		expect(['file', ...ribbonTabs().map((tab) => tab.id)]).toEqual([...EXCEL_RIBBON_TAB_IDS]);
	});

	it('hides File and refuses the commands that only open the backstage', async () => {
		expect([...resolveExcelCustomization(NO_FILE).commands]).toEqual([
			'file.info',
			'file.save-as',
			'file.options',
		]);
		const element = mount(NO_FILE);
		await settle();
		expect(ribbon(element).noFile).toBe(true);
		expect(ribbon(element).fileButton()).toBeNull();
		expect(await tabIds(element)).toEqual([
			'home',
			'insert',
			'page-layout',
			'formulas',
			'data',
			'review',
			'view',
			'help',
		]);
		ribbon(element).dispatchEvent(new CustomEvent('office-ribbon-file'));
		const { commands } = core(element);
		for (const id of ['file.info', 'file.save-as', 'file.options']) {
			expect(commands.isHidden(id)).toBe(true);
			expect(await commands.run(id)).toBe(false);
		}
		expect(backstage(element).open).toBe(false);
		expect(element.shadowRoot!.querySelector('.xve-empty')?.textContent).not.toContain('File');
	});

	it('shows no File key tip, and F opens nothing', async () => {
		const element = mount(NO_FILE);
		await settle();
		const grid = element.shadowRoot!.querySelector<HTMLElement>('[part~="grid"]')!;
		grid.dispatchEvent(new KeyboardEvent('keydown', { key: 'F10', bubbles: true, composed: true }));
		const badges = [...ribbon(element).querySelectorAll('.xve-keytip')].map((b) => b.textContent);
		expect(badges[0]).toBe('H');
		expect(badges).not.toContain('F');
		ribbon(element).dispatchEvent(new KeyboardEvent('keydown', { key: 'f', bubbles: true }));
		expect(backstage(element).open).toBe(false);
	});

	it('keeps Save, Ctrl+S, Print and file-command working', async () => {
		const element = mount(NO_FILE);
		const seen = vi.fn((event: Event) => event.preventDefault());
		element.addEventListener('file-command', seen);
		await core(element).commands.run('file.save');
		await core(element).commands.run('file.print');
		const grid = element.shadowRoot!.querySelector<HTMLElement>('[part~="grid"]')!;
		grid.dispatchEvent(
			new KeyboardEvent('keydown', {
				key: 's',
				ctrlKey: true,
				bubbles: true,
				composed: true,
				cancelable: true,
			}),
		);
		expect(seen.mock.calls.map(([event]) => (event as CustomEvent).detail.command)).toEqual([
			'save',
			'print',
			'save',
		]);
	});

	it('closes an open backstage when File is hidden at runtime and returns focus to the ribbon', async () => {
		const element = mount();
		await settle();
		ribbon(element).fileButton()!.click();
		expect(backstage(element).open).toBe(true);
		element.customization = NO_FILE;
		expect(backstage(element).open).toBe(false);
		expect(
			(ribbon(element).shadowRoot!.activeElement as HTMLElement | null)?.getAttribute('role'),
		).toBe('tab');
		element.customization = null;
		await settle();
		expect(ribbon(element).noFile).toBe(false);
		expect(core(element).commands.isHidden('file.info')).toBe(false);
		ribbon(element).fileButton()!.click();
		expect(backstage(element).open).toBe(true);
	});

	it('hides other native tabs and selects the first one left', async () => {
		const element = mount({ ribbon: { hiddenTabs: ['home', 'formulas'] } });
		expect(await tabIds(element)).toEqual([
			'insert',
			'page-layout',
			'data',
			'review',
			'view',
			'help',
		]);
		expect(ribbon(element).selected).toBe('insert');
		element.customization = {};
		expect(await tabIds(element)).toContain('home');
	});

	it('leaves commands that live only on a hidden tab out of Tell me', async () => {
		expect([...commandsOnlyInTabs(ribbonTabs(), new Set(['formulas']))]).toContain(
			'formulas.recent',
		);
		expect(commandsOnlyInTabs(ribbonTabs(), new Set(['formulas']))).not.toContain('home.autosum');
		expect(commandsOnlyInTabs(ribbonTabs(), new Set()).size).toBe(0);
		const element = mount({ ribbon: { hiddenTabs: ['formulas'] } });
		await settle();
		const bar = element.shadowRoot!.querySelector('office-ui-title-bar') as HTMLElement & {
			state: { search: { match(query: string): { id: string }[] } };
		};
		const found = (query: string) => bar.state.search.match(query).map((entry) => entry.id);
		expect(found('recently used')).not.toContain('formulas.recent');
		expect(found('autosum')).toContain('home.autosum');
		element.customization = null;
		await settle();
		expect(found('recently used')).toContain('formulas.recent');
	});
});
