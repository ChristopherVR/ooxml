// @vitest-environment jsdom
import { createWorkbook } from '@christophervr/xlsx-core';
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
import { registerRuleDialogs } from './register-rules.js';

afterEach(() => (document.body.innerHTML = ''));

function setup() {
	const ctx = createTestContext(createWorkbook());
	registerRuleDialogs(ctx);
	ctx.select('B2:B5');
	return ctx;
}
const dvs = (ctx: ReturnType<typeof setup>) => ctx.workbook()!.sheets[0]!.dataValidations;

describe('data validation dialog', () => {
	it('sets a whole number between rule with messages', async () => {
		const ctx = setup();
		const result = ctx.dialogs.open('data-validation');
		const d = dialogEl(ctx, 'data-validation');
		setValue(inputByLabel<HTMLSelectElement>(d, 'Allow:'), 'whole');
		setValue(inputByLabel(d, 'Minimum:'), '1');
		setValue(inputByLabel(d, 'Maximum:'), '=C1');
		setValue(inputByLabel(d, 'Title:'), 'Hint');
		setValue(inputByLabel<HTMLSelectElement>(d, 'Style:'), 'warning');
		clickButton(d, 'OK');
		await result;
		expect(dvs(ctx)).toHaveLength(1);
		expect(dvs(ctx)[0]).toMatchObject({
			type: 'whole',
			operator: 'between',
			formula1: '1',
			formula2: 'C1',
			promptTitle: 'Hint',
			errorStyle: 'warning',
			ranges: [{ start: { row: 1, col: 1 }, end: { row: 4, col: 1 } }],
		});
	});

	it('quotes a typed list, keeps a reference, and reloads it', async () => {
		const ctx = setup();
		let result = ctx.dialogs.open('data-validation');
		let d = dialogEl(ctx, 'data-validation');
		setValue(inputByLabel<HTMLSelectElement>(d, 'Allow:'), 'list');
		setValue(inputByLabel(d, 'Source:'), 'Yes,No');
		clickButton(d, 'OK');
		await result;
		expect(dvs(ctx)[0]).toMatchObject({ type: 'list', formula1: '"Yes,No"', showDropDown: true });
		result = ctx.dialogs.open('data-validation');
		d = dialogEl(ctx, 'data-validation');
		expect(inputByLabel(d, 'Source:').value).toBe('Yes,No');
		setValue(inputByLabel(d, 'Source:'), '=$D$1:$D$3');
		clickButton(d, 'OK');
		await result;
		expect(dvs(ctx)[0]?.formula1).toBe('$D$1:$D$3');
	});

	it('stores dates as serials', async () => {
		const ctx = setup();
		const result = ctx.dialogs.open('data-validation');
		const d = dialogEl(ctx, 'data-validation');
		setValue(inputByLabel<HTMLSelectElement>(d, 'Allow:'), 'date');
		setValue(inputByLabel<HTMLSelectElement>(d, 'Data:'), 'greaterThan');
		setValue(inputByLabel(d, 'Minimum:'), '2024-01-01');
		clickButton(d, 'OK');
		await result;
		expect(dvs(ctx)[0]).toMatchObject({ type: 'date', operator: 'greaterThan', formula1: '45292' });
	});

	it('Clear All removes the validation; invalid input keeps the dialog open', async () => {
		const ctx = setup();
		ctx
			.session()!
			.setDataValidation(
				0,
				{ ranges: [], type: 'decimal', operator: 'lessThan', formula1: '5' },
				{ start: { row: 1, col: 1 }, end: { row: 4, col: 1 } },
			);
		void ctx.dialogs.open('data-validation');
		let d = dialogEl(ctx, 'data-validation');
		setValue(inputByLabel(d, 'Minimum:'), '');
		clickButton(d, 'OK');
		await tick();
		expect(ctx.toasts).toHaveLength(1);
		clickButton(d, 'Clear All');
		clickButton(d, 'OK');
		await tick();
		expect(dvs(ctx)).toEqual([]);
		void ctx.dialogs.open('data-validation');
		d = dialogEl(ctx, 'data-validation');
		setValue(inputByLabel<HTMLSelectElement>(d, 'Allow:'), 'custom');
		pressKey(inputByLabel(d, 'Formula:'), 'Escape');
		await tick();
		clickButton(
			(void ctx.dialogs.open('data-validation'), dialogEl(ctx, 'data-validation')),
			'Cancel',
		);
		expect(dvs(ctx)).toEqual([]);
	});
});
