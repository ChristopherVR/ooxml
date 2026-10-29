// @vitest-environment jsdom
import { createWorkbook, getCell } from 'ooxml-core/xlsx';
import { afterEach, describe, expect, it } from 'vitest';
import {
	clickButton,
	createTestContext,
	dialogEl,
	inputByLabel,
	setValue,
} from '../commands/test-support';
import { registerToolDialogs } from './register-tools';

afterEach(() => (document.body.innerHTML = ''));

function seeded() {
	const ctx = createTestContext(createWorkbook());
	registerToolDialogs(ctx);
	const s = ctx.session()!;
	s.setCellValue(0, 0, 0, 'Day');
	s.setCellValue(0, 0, 1, 'Sales');
	for (let i = 1; i <= 8; i++) {
		s.setCellValue(0, i, 0, i);
		s.setCellValue(0, i, 1, 10 + 4 * i);
	}
	ctx.select('A1:B9');
	return ctx;
}

describe('Forecast Sheet dialog', () => {
	it('proposes the ranges of the selection and a forecast end', () => {
		const ctx = seeded();
		void ctx.dialogs.open('forecast-sheet');
		const dialog = dialogEl(ctx, 'forecast-sheet');
		expect((inputByLabel(dialog, 'Timeline range:') as HTMLInputElement).value).toBe('$A$1:$A$9');
		expect((inputByLabel(dialog, 'Values range:') as HTMLInputElement).value).toBe('$B$1:$B$9');
		expect((inputByLabel(dialog, 'Forecast end:') as HTMLInputElement).value).toBe('10');
	});

	it('creates the forecast sheet with a chart on OK', async () => {
		const ctx = seeded();
		const done = ctx.dialogs.open('forecast-sheet');
		const dialog = dialogEl(ctx, 'forecast-sheet');
		setValue(inputByLabel(dialog, 'Forecast end:'), '11');
		clickButton(dialog, 'OK');
		expect(await done).toBe(1);
		const wb = ctx.workbook()!;
		expect(wb.sheets).toHaveLength(2);
		const sheet = wb.sheets[1]!;
		expect(getCell(sheet, 0, 0)?.value).toBe('Timeline');
		expect(getCell(sheet, 9, 0)?.value).toBe(9);
		expect(getCell(sheet, 11, 0)?.value).toBe(11);
		expect(getCell(sheet, 10, 2)?.formula).toMatch(/^FORECAST\.ETS\(A11,/);
		expect(sheet.drawings).toHaveLength(1);
	});

	it('explains a forecast end before the data and stays open', () => {
		const ctx = seeded();
		void ctx.dialogs.open('forecast-sheet');
		const dialog = dialogEl(ctx, 'forecast-sheet');
		setValue(inputByLabel(dialog, 'Forecast end:'), '4');
		clickButton(dialog, 'OK');
		expect(ctx.toasts.at(-1)?.message).toBe(
			'The forecast end must come after the last timeline point.',
		);
		expect(ctx.root.querySelector('[data-dialog="forecast-sheet"]')).not.toBeNull();
		expect(ctx.workbook()!.sheets).toHaveLength(1);
	});

	it('rejects a range it cannot read and a missing end', () => {
		const ctx = seeded();
		void ctx.dialogs.open('forecast-sheet');
		const dialog = dialogEl(ctx, 'forecast-sheet');
		setValue(inputByLabel(dialog, 'Values range:'), 'nope');
		clickButton(dialog, 'OK');
		expect(ctx.toasts.at(-1)?.message).toBe('Reference is not valid.');
		setValue(inputByLabel(dialog, 'Values range:'), 'B1:B9');
		setValue(inputByLabel(dialog, 'Forecast end:'), '');
		clickButton(dialog, 'OK');
		expect(ctx.toasts.at(-1)?.message).toBe('Enter the forecast end as a date or a number.');
	});
});
