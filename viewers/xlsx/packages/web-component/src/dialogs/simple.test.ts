// @vitest-environment jsdom
import { createWorkbook, getCell } from '@christophervr/xlsx-core';
import { afterEach, describe, expect, it } from 'vitest';
import {
	clickButton,
	createTestContext,
	dialogEl,
	inputByLabel,
	pressKey,
	setValue,
	tick,
} from '../commands/test-support.js';
import { registerSimpleDialogs } from './simple.js';

afterEach(() => (document.body.innerHTML = ''));

function setup() {
	const ctx = createTestContext(createWorkbook({ sheets: ['Sheet1', 'Data'] }));
	registerSimpleDialogs(ctx);
	return ctx;
}

describe('rename sheet dialog', () => {
	it('opens with the current name and OK renames', async () => {
		const ctx = setup();
		const result = ctx.dialogs.open('rename-sheet');
		const dialog = dialogEl(ctx, 'rename-sheet');
		const input = inputByLabel(dialog, 'Sheet name');
		expect(input.value).toBe('Sheet1');
		setValue(input, 'Budget');
		clickButton(dialog, 'OK');
		expect(await result).toBe('Budget');
		expect(ctx.workbook()?.sheets[0]?.name).toBe('Budget');
		expect(ctx.root.querySelector('[data-dialog]')).toBeNull();
	});

	it('keeps the dialog open on an invalid or duplicate name', async () => {
		const ctx = setup();
		void ctx.dialogs.open('rename-sheet');
		const dialog = dialogEl(ctx, 'rename-sheet');
		setValue(inputByLabel(dialog, 'Sheet name'), 'Data');
		clickButton(dialog, 'OK');
		await tick();
		expect(ctx.root.querySelector('[data-dialog="rename-sheet"]')).not.toBeNull();
		expect(ctx.toasts.length).toBe(1);
		expect(ctx.workbook()?.sheets[0]?.name).toBe('Sheet1');
	});

	it('Cancel and Escape change nothing', async () => {
		const ctx = setup();
		const first = ctx.dialogs.open('rename-sheet');
		setValue(inputByLabel(dialogEl(ctx, 'rename-sheet'), 'Sheet name'), 'X');
		clickButton(dialogEl(ctx, 'rename-sheet'), 'Cancel');
		expect(await first).toBeUndefined();
		const second = ctx.dialogs.open('rename-sheet');
		pressKey(inputByLabel(dialogEl(ctx, 'rename-sheet'), 'Sheet name'), 'Escape');
		expect(await second).toBeUndefined();
		expect(ctx.workbook()?.sheets[0]?.name).toBe('Sheet1');
	});

	it('Enter in the field confirms', async () => {
		const ctx = setup();
		const result = ctx.dialogs.open('rename-sheet');
		const input = inputByLabel(dialogEl(ctx, 'rename-sheet'), 'Sheet name');
		setValue(input, 'Q1');
		pressKey(input, 'Enter');
		expect(await result).toBe('Q1');
	});
});

describe('number prompts', () => {
	it('sets the row height of the selected rows', async () => {
		const ctx = setup();
		ctx.select('A2:B3');
		const result = ctx.dialogs.open('row-height');
		const dialog = dialogEl(ctx, 'row-height');
		setValue(inputByLabel(dialog, 'Row height:'), '30');
		clickButton(dialog, 'OK');
		await result;
		const ws = ctx.workbook()!.sheets[0]!;
		expect(ws.rowInfo.get(1)?.height).toBe(30);
		expect(ws.rowInfo.get(2)?.height).toBe(30);
		expect(ctx.session()?.canUndo()).toBe(true);
	});

	it('rejects a column width out of range', async () => {
		const ctx = setup();
		void ctx.dialogs.open('column-width');
		const dialog = dialogEl(ctx, 'column-width');
		setValue(inputByLabel(dialog, 'Column width:'), '300');
		clickButton(dialog, 'OK');
		await tick();
		expect(dialogEl(ctx, 'column-width')).toBeTruthy();
		expect(ctx.workbook()!.sheets[0]!.columns).toEqual([]);
	});

	it('sets the default width as one undo step', async () => {
		const ctx = setup();
		const result = ctx.dialogs.open('default-width');
		const dialog = dialogEl(ctx, 'default-width');
		setValue(inputByLabel(dialog, 'Standard column width:'), '12');
		clickButton(dialog, 'OK');
		await result;
		expect(ctx.workbook()!.sheets[0]!.defaultColWidth).toBe(12);
		ctx.session()!.undo();
		expect(ctx.workbook()!.sheets[0]!.defaultColWidth).toBeUndefined();
	});
});

describe('comment and unhide dialogs', () => {
	it('adds a comment by the author', async () => {
		const ctx = setup();
		ctx.select('B2');
		const result = ctx.dialogs.open('comment');
		const dialog = dialogEl(ctx, 'comment');
		setValue(inputByLabel<HTMLTextAreaElement>(dialog, 'Comment'), 'Check this');
		clickButton(dialog, 'OK');
		await result;
		expect(ctx.workbook()!.sheets[0]!.comments).toEqual([
			{ address: { row: 1, col: 1 }, author: 'Tester', text: 'Check this' },
		]);
	});

	it('unhides the chosen sheet', async () => {
		const ctx = setup();
		ctx.session()!.setSheetState(1, 'hidden');
		const result = ctx.dialogs.open('unhide-sheet');
		const dialog = dialogEl(ctx, 'unhide-sheet');
		expect(dialog.querySelectorAll('[role="option"]').length).toBe(1);
		clickButton(dialog, 'OK');
		expect(await result).toBe('1');
		expect(ctx.workbook()!.sheets[1]!.state).toBe('visible');
		expect(getCell(ctx.workbook()!.sheets[0]!, 0, 0)).toBeUndefined();
	});
});
