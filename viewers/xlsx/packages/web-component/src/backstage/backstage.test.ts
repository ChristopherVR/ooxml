// @vitest-environment jsdom
import { afterEach, describe, expect, it } from 'vitest';
import { createWorkbook } from '@christophervr/xlsx-core';
import { createBackstage, type BackstageHost, type EditorOptions } from './index';
import { registerRibbonTabs, resetRibbonTabs } from '../ribbon/parts';
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
	core.ctx.root.append(backstage.element);
	return { core, backstage, calls, el: backstage.element };
}
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
		const calc = el.querySelector<HTMLSelectElement>('select[aria-label="Workbook calculation"]')!;
		calc.value = 'manual';
		calc.dispatchEvent(new Event('change'));
		const language = el.querySelector<HTMLSelectElement>('select[aria-label="Display language"]')!;
		language.value = 'zh-CN';
		language.dispatchEvent(new Event('change'));
		expect(calls).toContainEqual(['option', 'calculation', 'manual']);
		expect(calls).toContainEqual(['option', 'locale', 'zh-CN']);
	});

	it('hides ribbon commands from Customize Ribbon', () => {
		const { backstage, el, calls, core } = setup();
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
		const boxes = el.querySelectorAll<HTMLInputElement>('.xve-customize-row input');
		expect(boxes).toHaveLength(2);
		boxes[0]!.checked = false;
		boxes[0]!.dispatchEvent(new Event('change'));
		expect(calls).toContainEqual(['hidden', ['home.bold']]);
	});

	it('follows the display language', () => {
		const { backstage, el, core } = setup();
		core.setLocale('fr');
		backstage.relocalize();
		backstage.open();
		expect(el.querySelector('h2')!.textContent).toBe('Informations');
		expect(el.querySelector('.xve-backstage-back')!.textContent).toBe('Retour au classeur');
	});
});
