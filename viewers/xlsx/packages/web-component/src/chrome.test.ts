// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import { createEditSession, createWorkbook } from '@christophervr/xlsx-core';
import { createStatusBar, statisticsText } from './status-bar';
import { createTitleBar } from './title-bar';
import { buildPrintHtml, printRange } from './print';
import { createTemplateWorkbook } from './backstage';
import { flush, shellFixture, spyCommand } from './test-support/shell';
import { withExtension, saveExtension } from './file-commands';

function numbers() {
	const workbook = createWorkbook();
	const session = createEditSession(workbook);
	session.setCellInput(0, 0, 0, '2');
	session.setCellInput(0, 1, 0, '4');
	session.setCellInput(0, 2, 0, 'text');
	return workbook;
}

describe('status bar', () => {
	it('shows Average, Count and Sum for a multi-cell selection, none for one cell', () => {
		const { core } = shellFixture(numbers());
		core.selection.set({ ranges: [{ start: { row: 0, col: 0 }, end: { row: 2, col: 0 } }] });
		expect(statisticsText(core.ctx, new Set(['average', 'count', 'sum']))).toBe(
			'Average: 3    Count: 3    Sum: 6',
		);
		expect(statisticsText(core.ctx, new Set(['min', 'max', 'numericCount']))).toBe(
			'Numerical Count: 2    Minimum: 2    Maximum: 4',
		);
		core.selection.set({ active: { row: 0, col: 0 } });
		expect(statisticsText(core.ctx, new Set(['sum']))).toBe('');
	});

	it('shows the mode, read-only badge, notes button and zoom', () => {
		const { core } = shellFixture(numbers());
		const notes: string[] = [];
		const bar = createStatusBar(core.ctx, {
			showNotes: () => notes.push('open'),
			setZoom: () => undefined,
		});
		expect(bar.element.getAttribute('part')).toBe('status-bar');
		expect(bar.element.querySelector('.xve-status-mode')!.textContent).toBe('Ready');
		core.workbook!.warnings.push('Pivot tables are not shown');
		core.setReadOnly(true);
		bar.refresh();
		expect((bar.element.querySelector('.xve-status-badge') as HTMLElement).hidden).toBe(false);
		const button = bar.element.querySelector<HTMLButtonElement>('.xve-status-notes')!;
		expect(button.textContent).toBe('1 compatibility note');
		button.click();
		expect(notes).toEqual(['open']);
		expect(bar.element.querySelector('.xve-zoom-percent')!.textContent).toBe('100%');
	});
});

describe('title bar', () => {
	it('shows the file name and save state, toggles Editing/Viewing and runs undo', async () => {
		const { core } = shellFixture();
		const undo = spyCommand('edit.undo');
		core.commands.register(undo);
		const modes: boolean[] = [];
		const bar = createTitleBar(core.ctx, {
			save: () => undefined,
			setReadOnly: (on) => modes.push(on),
			isHidden: () => false,
			revealControl: () => false,
		});
		bar.setFileName('Budget.xlsx');
		bar.setSaveState('dirty');
		expect(bar.element.querySelector('.xve-filename')!.textContent).toBe('Budget.xlsx');
		expect(bar.element.querySelector('.xve-save-state')!.textContent).toBe('Unsaved changes');
		const select = bar.element.querySelector<HTMLSelectElement>('.xve-mode-select')!;
		select.value = 'viewing';
		select.dispatchEvent(new Event('change'));
		expect(modes).toEqual([true]);
		bar.element.querySelector<HTMLButtonElement>('[aria-label="Undo"]')!.click();
		await flush();
		expect(undo.runs).toHaveLength(1);
		core.setLocale('es');
		bar.relocalize();
		expect(bar.element.querySelector('.xve-save-state')!.textContent).toBe('Cambios sin guardar');
	});
});

describe('printing and templates', () => {
	it('prints the used range as an escaped table with merges', () => {
		const workbook = numbers();
		const session = createEditSession(workbook);
		session.setCellInput(0, 3, 0, '<b>&');
		session.setCellInput(0, 0, 1, 'Wide');
		session.merge(0, { start: { row: 0, col: 1 }, end: { row: 0, col: 2 } }, 'merge');
		expect(printRange(workbook, 0)).toEqual({ start: { row: 0, col: 0 }, end: { row: 3, col: 2 } });
		const html = buildPrintHtml(workbook, 0, 'Book<1>');
		expect(html).toContain('<title>Book&#60;1&#62;</title>');
		expect(html).toContain('&#60;b&#62;&#38;');
		expect(html).toContain('colspan="2"');
		expect(html.match(/<tr /g)).toHaveLength(4);
	});

	it('generates the templates with working formulas', () => {
		const t = (key: string) => key;
		const budget = createTemplateWorkbook('budget', t);
		expect(budget.sheets[0]!.name).toBe('Budget');
		expect(budget.sheets[0]!.rows.get(9)?.get(1)?.value).toBe(1450);
		const invoice = createTemplateWorkbook('invoice', t);
		expect(invoice.sheets[0]!.rows.get(14)?.get(3)?.value).toBeCloseTo(2045 * 1.15);
		const todo = createTemplateWorkbook('todo', t);
		expect(todo.sheets[0]!.rows.get(8)?.get(1)?.value).toBe(3);
	});

	it('keeps .xlsm and turns other sources into .xlsx on save', () => {
		expect(saveExtension('Macro.xlsm')).toBe('xlsm');
		expect(saveExtension('Old.xls')).toBe('xlsx');
		expect(withExtension('Data.csv', saveExtension('Data.csv'))).toBe('Data.xlsx');
		expect(withExtension('  ', 'xlsx')).toBe('Book1.xlsx');
	});
});
