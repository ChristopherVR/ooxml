// @vitest-environment jsdom
import { createWorkbook, getCell, styleAt } from 'ooxml-core/xlsx';
import { afterEach, describe, expect, it } from 'vitest';
import { clipState, clipboardCommands } from 'ooxml-core/xlsx/ui';
import {
	clickButton,
	createTestContext,
	dialogEl,
	inputByLabel,
	pressKey,
	setValue,
} from '../commands/test-support.js';
import { registerToolDialogs } from './register-tools.js';

afterEach(() => (document.body.innerHTML = ''));

function setup(grid?: { zoom: number }) {
	const ctx = createTestContext(
		createWorkbook({ sheets: ['A', 'B', 'C'] }),
		grid ? { grid: { zoom: () => grid.zoom, setZoom: (z: number) => (grid.zoom = z) } } : {},
	);
	registerToolDialogs(ctx);
	ctx.commands.registerAll(clipboardCommands());
	return ctx;
}
const radio = (root: HTMLElement, label: string) => inputByLabel(root, label);

describe('Paste Special', () => {
	it('copies all except destination borders through the shared dialog', async () => {
		const ctx = setup();
		const session = ctx.session()!;
		const source = { start: { row: 0, col: 0 }, end: { row: 0, col: 0 } };
		const dest = { start: { row: 0, col: 2 }, end: { row: 0, col: 2 } };
		session.setCellValue(0, 0, 0, 2);
		session.setCellValue(0, 0, 2, 9);
		session.applyStyle(0, [source], {
			font: { bold: true },
			border: { bottom: { style: 'thin', color: { rgb: 'FFFF0000' } } },
		});
		session.applyStyle(0, [dest], {
			border: { bottom: { style: 'double', color: { rgb: 'FF0000FF' } } },
		});
		clipState(ctx).payload = session.copy(0, source);
		ctx.select('C1');
		const result = ctx.commands.run('home.paste-special');
		const dialog = dialogEl(ctx, 'paste-special');
		radio(dialog, 'All except borders').click();
		clickButton(dialog, 'OK');
		await result;
		const workbook = ctx.workbook()!;
		const cell = getCell(workbook.sheets[0]!, 0, 2);
		expect(cell?.value).toBe(2);
		expect(styleAt(workbook, cell?.styleId).font.bold).toBe(true);
		expect(styleAt(workbook, cell?.styleId).border.bottom).toEqual({
			style: 'double',
			color: { rgb: 'FF0000FF' },
		});
		session.undo();
		expect(getCell(workbook.sheets[0]!, 0, 2)?.value).toBe(9);
	});
	it('multiplies a selection and retains destination formulas with Values', async () => {
		const ctx = setup();
		const s = ctx.session()!;
		s.setCellInput(0, 0, 0, '=1+1');
		s.setRangeValues(0, { row: 0, col: 2 }, [[10, 'text', 10]]);
		s.setCellInput(0, 0, 4, '=5+5');
		clipState(ctx).payload = s.copy(0, { start: { row: 0, col: 0 }, end: { row: 0, col: 0 } });
		ctx.select('C1:E1');
		const result = ctx.commands.run('home.paste-special');
		const dialog = dialogEl(ctx, 'paste-special');
		radio(dialog, 'Values').click();
		radio(dialog, 'Multiply').click();
		clickButton(dialog, 'OK');
		await result;
		const read = () => [2, 3, 4].map((col) => getCell(ctx.workbook()!.sheets[0]!, 0, col)?.value);
		expect(read()).toEqual([20, 'text', 20]);
		expect(getCell(ctx.workbook()!.sheets[0]!, 0, 4)?.formula).toBe('(5+5)*2');
		s.undo();
		expect(read()).toEqual([10, 'text', 10]);
	});
	it('combines values, transpose and skip blanks in one undo step', async () => {
		const ctx = setup();
		const s = ctx.session()!;
		s.setCellInput(0, 1, 0, '=""');
		s.setCellValue(0, 2, 0, 0);
		clipState(ctx).payload = s.copy(0, { start: { row: 0, col: 0 }, end: { row: 2, col: 0 } });
		s.setRangeValues(0, { row: 0, col: 2 }, [[9, 9, 9]]);
		ctx.select('C1');
		const result = ctx.commands.run('home.paste-special');
		const dialog = dialogEl(ctx, 'paste-special');
		radio(dialog, 'Values').click();
		inputByLabel(dialog, 'Transpose').click();
		inputByLabel(dialog, 'Skip blanks').click();
		clickButton(dialog, 'OK');
		await result;
		const read = () => [2, 3, 4].map((col) => getCell(ctx.workbook()!.sheets[0]!, 0, col)?.value);
		expect(read()).toEqual([9, '', 0]);
		expect(getCell(ctx.workbook()!.sheets[0]!, 0, 3)?.formula).toBeUndefined();
		s.undo();
		expect(read()).toEqual([9, 9, 9]);
	});
	it('pastes values only, and Cancel pastes nothing', async () => {
		const ctx = setup();
		const s = ctx.session()!;
		s.setCellInput(0, 0, 0, '2');
		s.setCellInput(0, 0, 1, '=A1*2');
		clipState(ctx).payload = s.copy(0, { start: { row: 0, col: 0 }, end: { row: 0, col: 1 } });
		ctx.select('A3');
		const cancelled = ctx.dialogs.open('paste-special');
		clickButton(dialogEl(ctx, 'paste-special'), 'Cancel');
		await cancelled;
		expect(getCell(ctx.workbook()!.sheets[0]!, 2, 1)).toBeUndefined();
		const result = ctx.commands.run('home.paste-special');
		const dialog = dialogEl(ctx, 'paste-special');
		expect(radio(dialog, 'All except borders').disabled).toBe(false);
		radio(dialog, 'Values').click();
		clickButton(dialog, 'OK');
		expect(await result).toBe(true);
		const pasted = getCell(ctx.workbook()!.sheets[0]!, 2, 1);
		expect(pasted?.value).toBe(4);
		expect(pasted?.formula).toBeUndefined();
	});
});

describe('Page Setup', () => {
	it('applies orientation, margins, header and print options in one step', async () => {
		const ctx = setup();
		const result = ctx.dialogs.open('page-setup', { tab: 'margins' });
		const dialog = dialogEl(ctx, 'page-setup');
		radio(dialog, 'Landscape').click();
		setValue(inputByLabel(dialog, 'Top'), '1.5');
		setValue(inputByLabel(dialog, 'Header:'), 'Budget');
		setValue(inputByLabel(dialog, 'Print area'), 'A1:D20');
		inputByLabel(dialog, 'Gridlines').click();
		clickButton(dialog, 'OK');
		await result;
		const ws = ctx.workbook()!.sheets[0]!;
		expect(ws.pageSetup).toMatchObject({
			orientation: 'landscape',
			header: '&CBudget',
			scale: 100,
		});
		expect(ws.pageSetup?.margins?.top).toBe(1.5);
		expect(ws.pageSetup?.printArea).toEqual({
			start: { row: 0, col: 0 },
			end: { row: 19, col: 3 },
		});
		expect(ws.printOptions?.gridLines).toBe(true);
		ctx.session()!.undo();
		expect(ctx.workbook()!.sheets[0]!.pageSetup).toBeUndefined();
		expect(ctx.workbook()!.sheets[0]!.printOptions?.gridLines).toBeUndefined();
	});

	it('keeps the dialog open on a bad margin; Escape cancels', async () => {
		const ctx = setup();
		void ctx.dialogs.open('page-setup');
		const dialog = dialogEl(ctx, 'page-setup');
		setValue(inputByLabel(dialog, 'Left'), '-1');
		clickButton(dialog, 'OK');
		await Promise.resolve();
		expect(ctx.toasts).toHaveLength(1);
		pressKey(inputByLabel(dialog, 'Left'), 'Escape');
		expect(ctx.root.querySelector('[data-dialog="page-setup"]')).toBeNull();
		expect(ctx.workbook()!.sheets[0]!.pageSetup).toBeUndefined();
	});
});

describe('Zoom', () => {
	it('starts on the current zoom and applies a preset or a custom value', async () => {
		const grid = { zoom: 100 };
		const ctx = setup(grid);
		const a = ctx.dialogs.open('zoom');
		const dialog = dialogEl(ctx, 'zoom');
		expect(radio(dialog, '100%').checked).toBe(true);
		radio(dialog, '75%').click();
		clickButton(dialog, 'OK');
		expect(await a).toBe(75);
		expect(grid.zoom).toBe(75);
		const b = ctx.dialogs.open('zoom');
		setValue(inputByLabel(dialogEl(ctx, 'zoom'), 'Percent'), '130');
		pressKey(inputByLabel(dialogEl(ctx, 'zoom'), 'Percent'), 'Enter');
		expect(await b).toBe(130);
		const c = ctx.dialogs.open('zoom');
		clickButton(dialogEl(ctx, 'zoom'), 'Cancel');
		await c;
		expect(grid.zoom).toBe(130);
	});
});

describe('Move or Copy', () => {
	it('moves the active sheet to the end', async () => {
		const ctx = setup();
		const result = ctx.dialogs.open('move-copy-sheet');
		const dialog = dialogEl(ctx, 'move-copy-sheet');
		dialog.querySelector<HTMLElement>('[data-value="end"]')!.click();
		clickButton(dialog, 'OK');
		expect(await result).toBe(2);
		expect(ctx.workbook()!.sheets.map((s) => s.name)).toEqual(['B', 'C', 'A']);
		expect(ctx.activeSheet()).toBe(2);
	});

	it('copies a sheet before another one', async () => {
		const ctx = setup();
		ctx.setActiveSheet(2);
		const result = ctx.dialogs.open('move-copy-sheet');
		const dialog = dialogEl(ctx, 'move-copy-sheet');
		dialog.querySelector<HTMLElement>('[data-value="0"]')!.click();
		inputByLabel(dialog, 'Create a copy').click();
		clickButton(dialog, 'OK');
		expect(await result).toBe(0);
		expect(ctx.workbook()!.sheets.map((s) => s.name)).toEqual(['C (2)', 'A', 'B', 'C']);
	});

	it('Cancel leaves the order alone', async () => {
		const ctx = setup();
		const result = ctx.dialogs.open('move-copy-sheet');
		clickButton(dialogEl(ctx, 'move-copy-sheet'), 'Cancel');
		await result;
		expect(ctx.workbook()!.sheets.map((s) => s.name)).toEqual(['A', 'B', 'C']);
	});
});
