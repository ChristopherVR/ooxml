// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import { createEditSession, createWorkbook } from 'ooxml-core/xlsx';
import { createStatusBar, statisticsText } from './status-bar';
import { createTitleBar } from './title-bar';
import { createRibbonActions } from './ribbon-actions';
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

	it('shows the mode, read-only badge, notes button and zoom', async () => {
		const { core } = shellFixture(numbers());
		const notes: string[] = [];
		const bar = createStatusBar(core.ctx, {
			showNotes: () => notes.push('open'),
			setZoom: () => undefined,
		});
		document.body.append(bar.element);
		type Updating = HTMLElement & { updateComplete: Promise<unknown> };
		const drawn = async () => {
			await (bar.element as Updating).updateComplete;
			return bar.element.shadowRoot!;
		};
		expect(bar.element.getAttribute('part')).toBe('status-bar');
		// Excel's order: statistics, then the view buttons, then the zoom slider.
		expect(bar.element.querySelector('.xve-status-stats')!.slot).toBe('summary');
		expect((await drawn()).querySelector('[data-item="mode"]')!.textContent).toBe('Ready');
		core.workbook!.warnings.push('Pivot tables are not shown');
		core.setReadOnly(true);
		bar.refresh();
		const root = await drawn();
		expect(root.querySelector('[data-item="readonly"]')!.textContent).toBe('Read-only');
		const button = root.querySelector<HTMLButtonElement>('[data-id="notes"]')!;
		expect(button.hidden).toBe(false);
		expect(button.textContent?.trim()).toBe('1 compatibility note');
		button.click();
		expect(notes).toEqual(['open']);
		const slider = bar.element.querySelector<Updating & { value: number }>(
			'office-ui-zoom-slider',
		)!;
		await slider.updateComplete;
		expect(slider.value).toBe(100);
		expect(slider.shadowRoot!.querySelector('output')!.textContent).toBe('100%');
		expect(root.querySelector<HTMLButtonElement>('[data-id="pageLayout"]')!.disabled).toBe(true);
	});
});

describe('title bar', () => {
	it('shows the file name and save state and runs undo', async () => {
		const { core } = shellFixture();
		const undo = spyCommand('edit.undo');
		core.commands.register(undo);
		const bar = createTitleBar(core.ctx, {
			save: () => undefined,
			isHidden: () => false,
			revealControl: () => false,
		});
		document.body.append(bar.element);
		const drawn = async () => {
			await (bar.element as HTMLElement & { updateComplete: Promise<unknown> }).updateComplete;
			return bar.element.shadowRoot!;
		};
		bar.setFileName('Budget.xlsx');
		bar.setSaveState('dirty');
		let root = await drawn();
		expect(root.querySelector('.name')!.textContent).toBe('Budget.xlsx');
		expect(root.querySelector('.status')!.textContent).toBe('Unsaved changes');
		// Editing mode, Comments and Share moved to the ribbon tab row (ribbon actions).
		expect(bar.element.querySelector('select, button')).toBeNull();
		root.querySelector<HTMLButtonElement>('[aria-label="Undo"]')!.click();
		await flush();
		expect(undo.runs).toHaveLength(1);
		core.setLocale('es');
		bar.relocalize();
		root = await drawn();
		expect(root.querySelector('.status')!.textContent).toBe('Cambios sin guardar');
	});

	it('searches commands from the shared search and runs the chosen one', async () => {
		const { core } = shellFixture();
		const undo = spyCommand('edit.undo');
		core.commands.register(undo);
		const bar = createTitleBar(core.ctx, {
			save: () => undefined,
			isHidden: () => false,
			revealControl: () => false,
		});
		document.body.append(bar.element);
		bar.element.dispatchEvent(
			new CustomEvent('office-command-search', { detail: { query: 'undo', command: 'edit.undo' } }),
		);
		await flush();
		expect(undo.runs).toHaveLength(1);
	});
});

describe('ribbon actions', () => {
	it('toggles Editing/Viewing, shows Comments and runs Share, labelled and translated', async () => {
		const { core } = shellFixture();
		const share = spyCommand('file.share');
		const comments = spyCommand('review.show-comments');
		core.commands.register(share);
		core.commands.register(comments);
		const modes: boolean[] = [];
		let hidden = false;
		const actions = createRibbonActions(core.ctx, {
			setReadOnly: (on) => modes.push(on),
			isHidden: (id) => hidden && id === 'file.share',
			collaboration: () => ({ active: true, status: 'connected', people: [] }) as never,
		});
		const node = actions.element;
		expect(node.slot).toBe('actions');
		const select = node.querySelector<HTMLSelectElement>('.xve-mode-select')!;
		const commentsButton = node.querySelector<HTMLButtonElement>('.xve-comments-button')!;
		const shareButton = node.querySelector<HTMLButtonElement>('.xve-share-button')!;
		// Excel's order: the editing mode just left of Comments, then Share.
		expect([...node.children].map((child) => child.className)).toEqual([
			'xve-mode',
			'xve-icon-button xve-comments-button',
			'xve-share-button',
		]);
		expect(select.getAttribute('aria-label')).toBe('Editing mode');
		expect(select.value).toBe('editing');
		expect(commentsButton.textContent).toBe('Comments');
		expect(shareButton.textContent).toBe('Share');
		expect(shareButton.getAttribute('aria-pressed')).toBe('true');
		select.value = 'viewing';
		select.dispatchEvent(new Event('change'));
		expect(modes).toEqual([true]);
		commentsButton.click();
		shareButton.click();
		await flush();
		expect(comments.runs).toHaveLength(1);
		expect(share.runs).toHaveLength(1);
		hidden = true;
		actions.refresh();
		expect(shareButton.hidden).toBe(true);
		core.setLocale('fr');
		actions.relocalize();
		expect(select.options[0]!.textContent).not.toBe('Editing');
		expect(commentsButton.textContent).not.toBe('Comments');
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
