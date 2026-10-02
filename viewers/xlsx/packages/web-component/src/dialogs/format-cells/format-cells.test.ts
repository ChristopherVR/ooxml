// @vitest-environment jsdom
import { BUILTIN_NUMBER_FORMATS, createWorkbook, getCell, styleAt } from '@christophervr/xlsx-core';
import { afterEach, describe, expect, it } from 'vitest';
import {
	clickButton,
	createTestContext,
	dialogEl,
	inputByLabel,
	pressKey,
	setValue,
} from '../../commands/test-support.js';
import { registerFormatCellsDialogs } from './index.js';
import { detectCategory, buildCode, DEFAULT_OPTIONS } from './number-categories.js';

afterEach(() => (document.body.innerHTML = ''));

function setup() {
	const ctx = createTestContext(createWorkbook());
	registerFormatCellsDialogs(ctx);
	ctx.session()!.setCellValue(0, 0, 0, 1234.5);
	return ctx;
}

const styleOf = (ctx: ReturnType<typeof setup>, row = 0, col = 0) => {
	const wb = ctx.workbook()!;
	return styleAt(wb, getCell(wb.sheets[0]!, row, col)?.styleId);
};

const option = (root: HTMLElement, value: string) =>
	root.querySelector<HTMLElement>(`[role="option"][data-value="${value}"]`)!;
const sample = (root: HTMLElement) => root.querySelector<HTMLElement>('[data-sample]')!.textContent;

describe('number categories', () => {
	it('detects the category of known codes and falls back to Custom', () => {
		expect(detectCategory('General').category).toBe('general');
		expect(detectCategory('0.00%')).toMatchObject({
			category: 'percentage',
			options: { decimals: 2 },
		});
		expect(detectCategory('#,##0.0').options).toMatchObject({ decimals: 1, thousands: true });
		expect(detectCategory(BUILTIN_NUMBER_FORMATS[44] ?? '').category).toBe('accounting');
		expect(detectCategory('m/d/yyyy').category).toBe('date');
		expect(detectCategory('0.0"kg"').category).toBe('custom');
		expect(buildCode('number', { ...DEFAULT_OPTIONS, decimals: 0, negative: 2 })).toBe('0_);(0)');
	});
});

describe('Format Cells: Number tab', () => {
	it('updates the sample live and applies the code on OK', async () => {
		const ctx = setup();
		const result = ctx.dialogs.open('format-cells');
		const dialog = dialogEl(ctx, 'format-cells');
		expect(sample(dialog)).toBe('1234.5');
		option(dialog, 'number').click();
		expect(sample(dialog)).toBe('1234.50');
		setValue(inputByLabel(dialog, 'Decimal places:'), '1');
		expect(sample(dialog)).toBe('1234.5');
		clickButton(dialog, 'OK');
		expect(await result).toBe(true);
		expect(styleOf(ctx).numFmt).toBe('0.0');
		expect(ctx.session()!.undoLabel()).toBe('Format Cells');
	});

	it('edits a custom code', async () => {
		const ctx = setup();
		const result = ctx.dialogs.open('format-cells');
		const dialog = dialogEl(ctx, 'format-cells');
		option(dialog, 'custom').click();
		setValue(inputByLabel(dialog, 'Type:'), '0.000');
		expect(sample(dialog)).toBe('1234.500');
		clickButton(dialog, 'OK');
		await result;
		expect(styleOf(ctx).numFmt).toBe('0.000');
	});

	it('opens on the detected or requested category', async () => {
		const ctx = setup();
		ctx
			.session()!
			.applyStyle(0, [{ start: { row: 0, col: 0 }, end: { row: 0, col: 0 } }], { numFmt: '0.00%' });
		void ctx.dialogs.open('format-cells');
		expect(option(dialogEl(ctx, 'format-cells'), 'percentage').getAttribute('aria-selected')).toBe(
			'true',
		);
		clickButton(dialogEl(ctx, 'format-cells'), 'Cancel');
		const result = ctx.dialogs.open('format-cells', { category: 'accounting' });
		clickButton(dialogEl(ctx, 'format-cells'), 'OK');
		await result;
		expect(styleOf(ctx).numFmt).toBe(BUILTIN_NUMBER_FORMATS[44]);
	});

	it('Cancel and Escape change nothing', async () => {
		const ctx = setup();
		const first = ctx.dialogs.open('format-cells');
		option(dialogEl(ctx, 'format-cells'), 'percentage').click();
		clickButton(dialogEl(ctx, 'format-cells'), 'Cancel');
		expect(await first).toBeUndefined();
		const second = ctx.dialogs.open('format-cells');
		pressKey(dialogEl(ctx, 'format-cells').querySelector('[role="tab"]')!, 'Escape');
		expect(await second).toBeUndefined();
		expect(styleOf(ctx).numFmt).toBe('General');
		expect(ctx.session()!.undoLabel()).not.toBe('Format Cells');
	});

	it('OK without changes leaves the workbook alone', async () => {
		const ctx = setup();
		const before = ctx.workbook()!.styles.length;
		const result = ctx.dialogs.open('format-cells');
		clickButton(dialogEl(ctx, 'format-cells'), 'OK');
		await result;
		expect(ctx.workbook()!.styles.length).toBe(before);
	});
});

describe('Format Cells: Alignment and Protection tabs', () => {
	it('applies only the changed alignment fields and merges', async () => {
		const ctx = setup();
		ctx.select('A1:B2');
		const result = ctx.dialogs.open('format-cells', { tab: 'alignment' });
		const dialog = dialogEl(ctx, 'format-cells');
		setValue(inputByLabel<HTMLSelectElement>(dialog, 'Horizontal:'), 'center');
		setValue(inputByLabel(dialog, 'Degrees'), '-45');
		inputByLabel(dialog, 'Wrap text').click();
		inputByLabel(dialog, 'Merge cells').click();
		clickButton(dialog, 'OK');
		await result;
		expect(styleOf(ctx).alignment).toEqual({
			horizontal: 'center',
			textRotation: 135,
			wrapText: true,
		});
		expect(ctx.workbook()!.sheets[0]!.merges).toHaveLength(1);
		ctx.session()!.undo();
		expect(ctx.workbook()!.sheets[0]!.merges).toHaveLength(0);
		expect(styleOf(ctx).alignment).toBeUndefined();
	});

	it('unlocks cells', async () => {
		const ctx = setup();
		const result = ctx.dialogs.open('format-cells', { tab: 'protection' });
		const dialog = dialogEl(ctx, 'format-cells');
		expect(
			dialog
				.querySelector<HTMLElement>('[role="tab"][data-tab="protection"]')
				?.getAttribute('aria-selected'),
		).toBe('true');
		inputByLabel(dialog, 'Locked').click();
		clickButton(dialog, 'OK');
		await result;
		expect(styleOf(ctx).protection).toEqual({ locked: false });
	});

	it('moves between tabs with the arrow keys', () => {
		const ctx = setup();
		void ctx.dialogs.open('format-cells');
		const dialog = dialogEl(ctx, 'format-cells');
		const first = dialog.querySelector<HTMLElement>('[role="tab"]')!;
		pressKey(first, 'ArrowRight');
		expect(
			dialog.querySelector('[role="tab"][aria-selected="true"]')?.getAttribute('data-tab'),
		).toBe('alignment');
	});
});
