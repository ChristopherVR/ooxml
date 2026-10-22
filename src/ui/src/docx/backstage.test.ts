// @vitest-environment jsdom
import { createDocument } from 'ooxml-core/docx';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { createRibbon } from './ribbon';
import { createBackstage, type BackstageHandlers } from './backstage';
import { documentStats } from 'ooxml-core/docx';
import { localizeElement } from './localization';
import { EMPTY_WORD_CUSTOMIZATION, resolveWordCustomization } from './customization';

afterEach(() => (document.body.innerHTML = ''));

function setup(overrides: Partial<BackstageHandlers> = {}) {
	const model = createDocument();
	model.blocks = [{ type: 'paragraph', id: 'p', runs: [{ text: 'Hello there' }] }];
	const calls: Array<[string, string | undefined]> = [];
	const options: Array<[string, string]> = [];
	const properties: Array<[string, string]> = [];
	const handlers: BackstageHandlers = {
		fileCommand: (command, name) => calls.push([command, name]),
		close: vi.fn(),
		summary: () => ({
			fileName: 'Report.docx',
			model,
			words: 2,
			stats: documentStats(model),
			saveState: 'dirty',
		}),
		options: () => ({ locale: 'en', theme: 'auto', author: 'Ann' }),
		setProperty: (key, value) => properties.push([key, value]),
		setOption: (key, value) => options.push([key, value]),
		ribbon: () => undefined,
		hiddenActions: () => [],
		setHiddenActions: () => {},
		customization: () => EMPTY_WORD_CUSTOMIZATION,
		regionContext: (region) => ({
			region,
			editor: document.body,
			fileName: 'Report.docx',
			locale: 'en',
			canEdit: true,
			close: () => undefined,
		}),
		...overrides,
	};
	const backstage = createBackstage(handlers);
	document.body.append(backstage.element, backstage.optionsDialog);
	return { backstage, handlers, calls, options, properties, model };
}
type OptionsDialog = HTMLElement & { category: string; open: boolean };
const tabs = (dialog: HTMLElement) =>
	[...dialog.shadowRoot!.querySelectorAll('[role="tab"]')].map((tab) => tab.textContent);
const field = <T extends HTMLElement>(dialog: HTMLElement, key: string) =>
	dialog.shadowRoot!.querySelector<T>(`[data-key="${key}"] input, [data-key="${key}"] select`)!;
type Updating = HTMLElement & { updateComplete: Promise<unknown> };
/** The navigation buttons of the shared backstage, once it has rendered. */
const navButtons = async (root: HTMLElement) => {
	await (root as Updating).updateComplete;
	return [...root.shadowRoot!.querySelectorAll<HTMLButtonElement>('button')];
};
const nav = async (root: HTMLElement) => (await navButtons(root)).map((b) => b.textContent?.trim());
const click = (root: HTMLElement, label: string) =>
	[...root.querySelectorAll<HTMLButtonElement>('.dve-backstage-content button')]
		.find((b) => b.textContent === label)!
		.click();

describe('File backstage', () => {
	it('starts at Home with working local file actions', () => {
		const { backstage, calls } = setup();
		backstage.open();
		expect(backstage.element.querySelector('h2')?.textContent).toBe('Home');
		click(backstage.element, 'Open');
		expect(calls).toEqual([['open', undefined]]);
	});

	it('makes the editor inert while open and restores focus and prior inert state', () => {
		const opener = document.createElement('button');
		const disabledRegion = document.createElement('div');
		disabledRegion.inert = true;
		document.body.append(opener, disabledRegion);
		const { backstage } = setup();
		opener.focus();
		backstage.open('info');
		expect(opener.inert).toBe(true);
		backstage.close();
		expect(document.activeElement).toBe(opener);
		expect(opener.inert).toBe(false);
		expect(disabledRegion.inert).toBe(true);
	});
	it("lists Word's file actions in order", async () => {
		const { backstage } = setup();
		expect(await nav(backstage.element)).toEqual([
			'←',
			'Home',
			'Info',
			'New',
			'Open',
			'Save',
			'Save As',
			'Print',
			'Export',
			'Options',
		]);
	});

	it('Info shows the file, its state and the properties panel', () => {
		const { backstage } = setup();
		backstage.open('info');
		const text = backstage.element.textContent!;
		expect(text).toContain('Report.docx');
		expect(text).toContain('Unsaved changes');
		expect(text).toContain('Characters (with spaces)');
		expect(backstage.element.querySelector('.dve-backstage-properties')!.textContent).toContain(
			'11',
		);
		expect(text).toContain('Letter'.toUpperCase());
	});

	it('Save runs the save command and closes', async () => {
		const { backstage, calls, handlers } = setup();
		backstage.open();
		(await navButtons(backstage.element)).find((b) => b.textContent?.trim() === 'Save')!.click();
		expect(calls).toEqual([['save', undefined]]);
		expect(handlers.close).toHaveBeenCalled();
	});

	it('Save As sends the chosen file name, and ignores an empty one', () => {
		const { backstage, calls } = setup();
		backstage.open('saveAs');
		const name = backstage.element.querySelector<HTMLInputElement>(
			'input[aria-label="File name"]',
		)!;
		expect(name.value).toBe('Report.docx');
		name.value = 'Final draft';
		click(backstage.element, 'Save');
		expect(calls).toEqual([['saveAs', 'Final draft']]);
		backstage.open('saveAs');
		const again = backstage.element.querySelector<HTMLInputElement>(
			'input[aria-label="File name"]',
		)!;
		again.value = '   ';
		click(backstage.element, 'Save');
		expect(calls).toHaveLength(1);
	});

	it('Export offers PDF (via print), DOCX and plain text', () => {
		const { backstage, calls } = setup();
		backstage.open('export');
		click(backstage.element, 'Create PDF');
		backstage.open('export');
		click(backstage.element, 'Save a copy as DOCX');
		backstage.open('export');
		click(backstage.element, 'Export as plain text');
		expect(calls.map(([command]) => command)).toEqual(['print', 'export', 'exportText']);
	});

	it('Print shows paper facts and prints', () => {
		const { backstage, calls } = setup();
		backstage.open('print');
		expect(backstage.element.textContent).toContain('Paper size');
		click(backstage.element, 'Print');
		expect(calls).toEqual([['print', undefined]]);
	});

	it('Options opens the shared dialog over the document, in the Word Options categories', () => {
		const { backstage } = setup();
		backstage.open('info');
		backstage.open('options');
		const dialog = backstage.optionsDialog as OptionsDialog;
		expect(backstage.isOpen).toBe(false);
		expect(dialog.open).toBe(true);
		expect(tabs(dialog)).toEqual([
			'General',
			'Display',
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
		expect(dialog.getAttribute('heading')).toBe('Options');
	});

	it('Options reads the current values and applies them on OK', () => {
		const { backstage, options } = setup();
		backstage.open('options');
		const dialog = backstage.optionsDialog as OptionsDialog;
		const name = field<HTMLInputElement>(dialog, 'author');
		expect(name.value).toBe('Ann');
		name.value = 'Bea';
		name.dispatchEvent(new Event('input'));
		dialog.category = 'language';
		const language = field<HTMLSelectElement>(dialog, 'locale');
		expect(language.value).toBe('en');
		language.value = 'de';
		language.dispatchEvent(new Event('change'));
		expect(options).toEqual([]);
		dialog.shadowRoot!.querySelector<HTMLButtonElement>('[data-action="ok"]')!.click();
		expect(options).toEqual([
			['locale', 'de'],
			['author', 'Bea'],
		]);
		expect(dialog.open).toBe(false);
	});

	it('Options offers only the display settings the editor reports, and marks the rest', () => {
		const { backstage } = setup({
			options: () => ({
				locale: 'en',
				theme: 'dark',
				author: 'Ann',
				spellCheck: true,
				showRuler: false,
			}),
		});
		backstage.open('options');
		const dialog = backstage.optionsDialog as OptionsDialog;
		expect(field<HTMLSelectElement>(dialog, 'theme').value).toBe('dark');
		dialog.category = 'display';
		expect(field<HTMLInputElement>(dialog, 'showRuler')).toBeTruthy();
		expect(dialog.shadowRoot!.querySelector('[data-key="showMarks"]')).toBeNull();
		dialog.category = 'save';
		expect(dialog.shadowRoot!.querySelector('.unavailable')!.textContent).toContain(
			'Not available yet',
		);
		dialog.category = 'general';
		expect(field<HTMLInputElement>(dialog, 'miniToolbar').disabled).toBe(true);
	});

	it('follows the display language', async () => {
		const { backstage } = setup();
		backstage.element.dataset.editorLocale = 'fr';
		backstage.relocalize();
		backstage.open('saveAs');
		expect(backstage.element.querySelector('h2')!.textContent).toBe('Enregistrer sous');
		localizeElement(backstage.element, 'fr');
		expect(await nav(backstage.element)).toContain('Exporter');
	});
});

describe('Customize Ribbon', () => {
	it('lists every command with a checkbox and reports the unchecked ones as hidden', () => {
		const ribbon = createRibbon();
		document.body.append(ribbon);
		const hidden: string[][] = [];
		const { backstage } = setup({
			ribbon: () => ribbon,
			hiddenActions: () => ['italic'],
			setHiddenActions: (ids) => hidden.push(ids),
		});
		backstage.open('customize');
		const dialog = backstage.optionsDialog as OptionsDialog;
		expect(dialog.category).toBe('ribbon');
		const boxes = [...dialog.querySelectorAll<HTMLInputElement>('.dve-customize-row input')];
		const named = (label: string) => boxes.find((b) => b.parentElement!.textContent === label)!;
		expect(boxes.length).toBeGreaterThan(40);
		expect(named('Bold').checked).toBe(true);
		expect(named('Italic').checked).toBe(false);
		named('Bold').click();
		expect(hidden.at(-1)).toEqual(expect.arrayContaining(['bold', 'italic']));
		dialog.querySelector<HTMLButtonElement>('.dve-options-action')!.click();
		expect(hidden.at(-1)).toEqual([]);
	});

	it('leaves out the commands a host capability removed', () => {
		const ribbon = createRibbon();
		document.body.append(ribbon);
		const hidden: string[][] = [];
		const { backstage } = setup({
			ribbon: () => ribbon,
			setHiddenActions: (ids) => hidden.push(ids),
			customization: () =>
				resolveWordCustomization({ capabilities: { print: false, copy: false } }),
		});
		backstage.open('customize');
		const dialog = backstage.optionsDialog as OptionsDialog;
		const labels = [...dialog.querySelectorAll('.dve-customize-row')].map((row) => row.textContent);
		expect(labels).toContain('Bold');
		expect(labels).not.toContain('Print');
		expect(labels).not.toContain('Copy');
		dialog.querySelector<HTMLButtonElement>('.dve-options-action')!.click();
		expect(hidden.at(-1)).toEqual([]);
	});

	it('puts the ribbon back when the dialog is cancelled', () => {
		const ribbon = createRibbon();
		document.body.append(ribbon);
		let hiddenNow: string[] = ['italic'];
		const { backstage } = setup({
			ribbon: () => ribbon,
			hiddenActions: () => hiddenNow,
			setHiddenActions: (ids) => (hiddenNow = ids),
		});
		backstage.open('customize');
		const dialog = backstage.optionsDialog as OptionsDialog;
		dialog.querySelector<HTMLInputElement>('.dve-customize-row input')!.click();
		expect(hiddenNow).not.toEqual(['italic']);
		dialog.shadowRoot!.querySelector<HTMLButtonElement>('[data-action="cancel"]')!.click();
		expect(hiddenNow).toEqual(['italic']);
	});

	it('shows the Customize Ribbon page in the display language', () => {
		const ribbon = createRibbon();
		document.body.append(ribbon);
		const { backstage } = setup({
			ribbon: () => ribbon,
		});
		backstage.element.dataset.editorLocale = 'fr';
		backstage.relocalize();
		backstage.open('customize');
		const dialog = backstage.optionsDialog as OptionsDialog;
		expect(tabs(dialog)).toContain('Personnaliser le ruban');
		expect(dialog.shadowRoot!.querySelector('[data-action="cancel"]')!.textContent).toContain(
			'Annuler',
		);
		const text = dialog.textContent!;
		expect(text).toContain('Réinitialiser toutes les personnalisations');
		expect(text).toContain('Choisissez les commandes affichées dans le ruban.');
	});
});

it('edits document properties from Info', () => {
	const s = setup();
	s.model.properties = { title: 'Report' };
	s.backstage.open('info');
	const title = s.backstage.element.querySelector<HTMLInputElement>('[aria-label="Title"]')!;
	expect(title.value).toBe('Report');
	const author = s.backstage.element.querySelector<HTMLInputElement>('[aria-label="Author"]')!;
	author.value = 'Bo';
	author.dispatchEvent(new Event('change'));
	expect(s.properties).toEqual([['creator', 'Bo']]);
});
