// @vitest-environment jsdom
import { createWorkbook, getCell } from '@christophervr/xlsx-core';
import { afterEach, describe, expect, it } from 'vitest';
import { clipState } from '../commands/clipboard.js';
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
	return ctx;
}
const radio = (root: HTMLElement, label: string) => inputByLabel(root, label);

describe('Paste Special', () => {
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
		const result = ctx.dialogs.open('paste-special');
		const dialog = dialogEl(ctx, 'paste-special');
		expect(radio(dialog, 'All except borders').disabled).toBe(true);
		radio(dialog, 'Values').click();
		clickButton(dialog, 'OK');
		expect(await result).toBe('values');
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
