// @vitest-environment jsdom
import { afterEach, describe, expect, it } from 'vitest';
import { createWorkbook } from 'ooxml-core/xlsx';
import { createBackstage, type BackstageHost, type EditorOptions } from './index';
import { registerRibbonTabs, resetRibbonTabs } from 'ooxml-core/xlsx/ui';
import { shellFixture, spyCommand } from '../test-support/shell';

function setup(workbook = createWorkbook()) {
	const { core } = shellFixture(workbook);
	const calls: unknown[][] = [];
	let hidden: string[] = [];
	const options: EditorOptions = {
		locale: 'en',
		theme: 'auto',
		author: 'Ada',
		calculation: 'automatic',
		iteration: undefined,
	};
	const host: BackstageHost = {
		ctx: core.ctx,
		fileName: () => 'Book1.xlsx',
		saveState: () => 'saved',
		fileCommand: (...args) => calls.push(['file', ...args]),
		newFromTemplate: (id) => calls.push(['template', id]),
		close: () => calls.push(['close']),
		options: () => options,
		setOption: (key, value) => calls.push(['option', key, value]),
		setProperty: (key, value) => calls.push(['property', key, value]),
		hiddenActions: () => hidden,
		setHiddenActions: (ids) => {
			hidden = ids;
			calls.push(['hidden', ids]);
		},
	};
	const backstage = createBackstage(host);
	core.ctx.root.append(backstage.element, backstage.optionsDialog);
	return { core, backstage, calls, el: backstage.element };
}
type OptionsDialog = HTMLElement & { category: string; open: boolean };
const dialogTabs = (dialog: HTMLElement) =>
	[...dialog.shadowRoot!.querySelectorAll('[role="tab"]')].map((tab) => tab.textContent);
const dialogField = <T extends HTMLElement>(dialog: HTMLElement, key: string) =>
	dialog.shadowRoot!.querySelector<T>(`[data-key="${key}"] input, [data-key="${key}"] select`)!;
const dialogAction = (dialog: HTMLElement, action: 'ok' | 'cancel' | 'reset') =>
	dialog.shadowRoot!.querySelector<HTMLButtonElement>(`[data-action="${action}"]`)!;
const button = (root: HTMLElement, text: string) =>
	[...root.querySelectorAll<HTMLButtonElement>('.xve-backstage-content button')].find(
		(b) => b.textContent?.trim() === text,
	)!;

afterEach(() => resetRibbonTabs());

describe('File backstage', () => {
	it('shows Info with properties, the xls compatibility note and the loader warnings', () => {
		const workbook = createWorkbook();
		workbook.format = 'xls';
		workbook.warnings.push('Pivot tables are kept but not shown.');
		const { backstage, el, calls } = setup(workbook);
		backstage.open();
		expect(el.getAttribute('part')).toBe('backstage');
		const notes = [...el.querySelectorAll('.xve-compatibility-notes li')].map(
			(li) => li.textContent,
		);
		expect(notes[0]).toContain('Excel 97-2003');
		expect(notes[1]).toBe('Pivot tables are kept but not shown.');
		const title = el.querySelector<HTMLInputElement>('input[aria-label="Title"]')!;
		title.value = 'Q3';
		title.dispatchEvent(new Event('change'));
		expect(calls).toContainEqual(['property', 'title', 'Q3']);
	});

	it('creates blank and template workbooks from New', () => {
		const { backstage, el, calls } = setup();
		backstage.open('new');
		el.querySelector<HTMLButtonElement>('[aria-label="Blank workbook"]')!.click();
		el.querySelector<HTMLButtonElement>('[aria-label="Invoice"]')!.click();
		expect(calls).toContainEqual(['file', 'new', undefined]);
		expect(calls).toContainEqual(['template', 'invoice']);
	});

	it('saves as xlsx or csv with the chosen name', () => {
		const { backstage, el, calls } = setup();
		backstage.open('saveAs');
		el.querySelector<HTMLInputElement>('input[aria-label="File name"]')!.value = 'Report';
		button(el, 'Save').click();
		backstage.open('saveAs');
		el.querySelector<HTMLInputElement>('input[aria-label="File name"]')!.value = 'Values';
		el.querySelector<HTMLSelectElement>('select[aria-label="Format"]')!.value = 'csv';
		button(el, 'Save').click();
		expect(calls).toContainEqual(['file', 'saveAs', 'Report.xlsx']);
		expect(calls).toContainEqual(['file', 'exportCsv', 'Values.csv']);
	});

	it('offers Export and Options for language, theme, author and calculation', () => {
		const { backstage, el, calls } = setup();
		backstage.open('export');
		button(el, 'Export as CSV').click();
		expect(calls).toContainEqual(['file', 'exportCsv', undefined]);
		backstage.open('options');
		const dialog = backstage.optionsDialog as OptionsDialog;
		dialog.category = 'formulas';
		const calc = dialogField<HTMLSelectElement>(dialog, 'calculation');
		calc.value = 'manual';
		calc.dispatchEvent(new Event('change'));
		dialog.category = 'language';
		const language = dialogField<HTMLSelectElement>(dialog, 'locale');
		language.value = 'zh-CN';
		language.dispatchEvent(new Event('change'));
		expect(calls.filter((call) => call[0] === 'option')).toEqual([]);
		dialogAction(dialog, 'ok').click();
		expect(calls).toContainEqual(['option', 'calculation', 'manual']);
		expect(calls).toContainEqual(['option', 'locale', 'zh-CN']);
	});

	it('opens Options as a dialog over the workbook, in the Excel Options categories', () => {
		const { backstage, el } = setup();
		backstage.open('info');
		backstage.open('options');
		const dialog = backstage.optionsDialog as OptionsDialog;
		expect(backstage.isOpen).toBe(false);
		expect(el.querySelector('[data-backstage-page="options"]')).toBeNull();
		expect(dialog.open).toBe(true);
		expect(dialogTabs(dialog)).toEqual([
			'General',
			'Formulas',
			'Data',
			'Proofing',
			'Save',
			'Language',
			'Accessibility',
			'Advanced',
			'Customize Ribbon',
			'Quick Access Toolbar',
			'Add-ins',
			'Trust Center',
		]);
		dialog.category = 'data';
		expect(dialog.shadowRoot!.querySelector('.unavailable')!.textContent).toContain(
			'Not available yet',
		);
		dialog.category = 'formulas';
		expect(dialogField<HTMLInputElement>(dialog, 'r1c1').disabled).toBe(true);
	});

	it('resets the Excel options to their defaults from Reset all options', () => {
		const { backstage, calls } = setup();
		backstage.open('options');
		const dialog = backstage.optionsDialog as OptionsDialog;
		dialog.category = 'formulas';
		const calc = dialogField<HTMLSelectElement>(dialog, 'calculation');
		calc.value = 'manual';
		calc.dispatchEvent(new Event('change'));
		dialogAction(dialog, 'reset').click();
		expect(dialogField<HTMLSelectElement>(dialog, 'calculation').value).toBe('automatic');
		dialogAction(dialog, 'ok').click();
		expect(calls.filter((call) => call[0] === 'option')).toEqual([]);
	});

	it('hides ribbon commands from Customize Ribbon', () => {
		const { backstage, calls, core } = setup();
		core.commands.registerAll([
			spyCommand('home.bold', { label: 'Bold' }),
			spyCommand('home.italic', { label: 'Italic' }),
		]);
		registerRibbonTabs([
			{
				id: 'home',
				label: 'Home',
				groups: [
					{
						id: 'font',
						label: 'Font',
						controls: [
							{ kind: 'toggle', command: 'home.bold' },
							{ kind: 'toggle', command: 'home.italic' },
						],
					},
				],
			},
		]);
		backstage.open('customize');
		const boxes = backstage.optionsDialog.querySelectorAll<HTMLInputElement>(
			'.xve-customize-row input',
		);
		expect(boxes).toHaveLength(2);
		boxes[0]!.checked = false;
		boxes[0]!.dispatchEvent(new Event('change'));
		expect(calls).toContainEqual(['hidden', ['home.bold']]);
	});

	it('turns iterative calculation on and off from the Formulas options', async () => {
		const { backstage, calls } = setup();
		backstage.open('options');
		const dialog = backstage.optionsDialog as OptionsDialog & { updateComplete: Promise<unknown> };
		dialog.category = 'formulas';
		await dialog.updateComplete;
		const toggle = dialogField<HTMLInputElement>(dialog, 'iterate');
		const count = dialogField<HTMLInputElement>(dialog, 'iterationCount');
		const delta = dialogField<HTMLInputElement>(dialog, 'iterationDelta');
		expect(toggle.checked).toBe(false);
		expect([count.disabled, delta.disabled]).toEqual([true, true]);
		toggle.click();
		await dialog.updateComplete;
		expect(count.disabled).toBe(false);
		count.value = '50';
		count.dispatchEvent(new Event('change'));
		delta.value = '0.01';
		delta.dispatchEvent(new Event('change'));
		dialogAction(dialog, 'ok').click();
		expect(calls.at(-1)).toEqual(['option', 'iteration', { count: 50, delta: 0.01 }]);
		backstage.open('options');
		dialog.category = 'formulas';
		await dialog.updateComplete;
		expect(dialogField<HTMLInputElement>(dialog, 'iterate').checked).toBe(false);
		dialogField<HTMLInputElement>(dialog, 'iterate').click();
		dialogAction(dialog, 'ok').click();
		expect(calls.at(-1)).toEqual(['option', 'iteration', { count: 100, delta: 0.001 }]);
	});

	it('follows the display language', () => {
		const { backstage, el, core } = setup();
		core.setLocale('fr');
		backstage.relocalize();
		backstage.open();
		expect(el.querySelector('h2')!.textContent).toBe('Informations');
		expect(el.getAttribute('back-label')).toBe('Retour au classeur');
	});

	it('draws the navigation with the shared element and routes its events', async () => {
		const { backstage, el, calls } = setup();
		backstage.open();
		await (el as HTMLElement & { updateComplete: Promise<unknown> }).updateComplete;
		const nav = [...el.shadowRoot!.querySelectorAll<HTMLButtonElement>('[data-backstage-item]')];
		expect(nav.map((item) => item.dataset.backstageItem)).toEqual([
			'info',
			'new',
			'open',
			'save',
			'saveAs',
			'print',
			'share',
			'export',
			'options',
		]);
		expect(backstage.isOpen).toBe(true);
		nav.find((item) => item.dataset.backstageItem === 'print')!.click();
		expect(el.querySelector<HTMLElement>('[data-backstage-page="print"]')!.hidden).toBe(false);
		expect(el.querySelector<HTMLElement>('[data-backstage-page="info"]')!.hidden).toBe(true);
		nav.find((item) => item.dataset.backstageItem === 'save')!.click();
		expect(calls).toContainEqual(['close']);
		expect(calls).toContainEqual(['file', 'save']);
		calls.length = 0;
		el.shadowRoot!.querySelector<HTMLButtonElement>('[data-backstage="back"]')!.click();
		expect(calls).toEqual([['close']]);
		expect(backstage.isOpen).toBe(false);
	});
});
