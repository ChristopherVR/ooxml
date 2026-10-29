// @vitest-environment jsdom
import { createWorkbook, getCell } from 'ooxml-core/xlsx';
import { afterEach, describe, expect, it, vi } from 'vitest';
import {
	clickButton,
	createTestContext,
	dialogEl,
	inputByLabel,
	setValue,
} from '../commands/test-support';
import { registerToolDialogs } from './register-tools';

afterEach(() => (document.body.innerHTML = ''));

function setup() {
	const ctx = createTestContext(createWorkbook());
	registerToolDialogs(ctx);
	return ctx;
}
type Ctx = ReturnType<typeof setup>;
const v = (ctx: Ctx, row: number, col: number) =>
	getCell(ctx.workbook()!.sheets[0]!, row, col)?.value;

describe('Goal Seek', () => {
	const seeded = () => {
		const ctx = setup();
		const s = ctx.session()!;
		s.setCellValue(0, 0, 0, 2);
		s.setCellInput(0, 1, 0, '=A1*3+4');
		ctx.select('A2');
		return ctx;
	};

	it('solves and keeps the result on OK of the status dialog', async () => {
		const ctx = seeded();
		const done = ctx.dialogs.open('goal-seek');
		const dialog = dialogEl(ctx, 'goal-seek');
		expect((inputByLabel(dialog, 'Set cell:') as HTMLInputElement).value).toBe('A2');
		setValue(inputByLabel(dialog, 'To value:'), '25');
		setValue(inputByLabel(dialog, 'By changing cell:'), '$A$1');
		clickButton(dialog, 'OK');

		const status = await vi.waitFor(() => dialogEl(ctx, 'goal-seek-status'));
		expect(status.textContent).toContain('found a solution');
		expect(status.textContent).toContain('Current value: 25');
		clickButton(status, 'OK');
		expect(await done).toBe(true);
		expect(v(ctx, 0, 0) as number).toBeCloseTo(7, 3);
	});

	it('restores the original value when the status dialog is cancelled', async () => {
		const ctx = seeded();
		const done = ctx.dialogs.open('goal-seek');
		const dialog = dialogEl(ctx, 'goal-seek');
		setValue(inputByLabel(dialog, 'To value:'), '25');
		setValue(inputByLabel(dialog, 'By changing cell:'), 'A1');
		clickButton(dialog, 'OK');

		const status = await vi.waitFor(() => dialogEl(ctx, 'goal-seek-status'));
		clickButton(status, 'Cancel');
		expect(await done).toBe(false);
		expect(v(ctx, 0, 0)).toBe(2);
		expect(v(ctx, 1, 0)).toBe(10);
	});

	it('explains an invalid request and stays open', () => {
		const ctx = seeded();
		void ctx.dialogs.open('goal-seek');
		const dialog = dialogEl(ctx, 'goal-seek');
		setValue(inputByLabel(dialog, 'By changing cell:'), 'A1');
		clickButton(dialog, 'OK');
		expect(ctx.toasts.at(-1)?.message).toBe('The To value must be a number.');
		expect(ctx.root.querySelector('[data-dialog="goal-seek"]')).not.toBeNull();
	});
});

describe('Data Table', () => {
	it('creates a one-variable table from the selection', async () => {
		const ctx = setup();
		const s = ctx.session()!;
		s.setCellValue(0, 0, 0, 3);
		s.setCellInput(0, 0, 2, '=A1*10');
		[1, 2].forEach((n, i) => s.setCellValue(0, 1 + i, 1, n));
		ctx.select('B1:C3');
		const done = ctx.dialogs.open('data-table');
		const dialog = dialogEl(ctx, 'data-table');
		setValue(inputByLabel(dialog, 'Column input cell:'), 'A1');
		clickButton(dialog, 'OK');
		expect(await done).toBe(true);
		expect(v(ctx, 1, 2)).toBe(10);
		expect(v(ctx, 2, 2)).toBe(20);
	});

	it('reports a missing input cell', () => {
		const ctx = setup();
		ctx.select('B1:C3');
		void ctx.dialogs.open('data-table');
		clickButton(dialogEl(ctx, 'data-table'), 'OK');
		expect(ctx.toasts.at(-1)?.message).toBe(
			'Enter a row input cell, a column input cell, or both.',
		);
	});
});

describe('Scenario Manager', () => {
	it('adds a scenario with its values and shows it', async () => {
		const ctx = setup();
		const s = ctx.session()!;
		s.setCellValue(0, 0, 0, 5);
		ctx.select('A1');
		void ctx.dialogs.open('scenario-manager');
		const manager = dialogEl(ctx, 'scenario-manager');
		clickButton(manager, 'Add...');
		const edit = await vi.waitFor(() => dialogEl(ctx, 'scenario-edit'));
		setValue(inputByLabel(edit, 'Scenario name:'), 'Big');
		clickButton(edit, 'OK');
		const values = await vi.waitFor(() => dialogEl(ctx, 'scenario-values'));
		setValue(inputByLabel(values, 'A1:'), '99');
		clickButton(values, 'OK');
		await vi.waitFor(() => expect(s.workbook.sheets[0]!.scenarios?.list).toHaveLength(1));
		expect(s.workbook.sheets[0]!.scenarios?.list[0]).toMatchObject({
			name: 'Big',
			cells: [{ ref: 'A1', value: '99' }],
		});
		await vi.waitFor(() => expect(manager.querySelectorAll('[role="option"]')).toHaveLength(1));
		(manager.querySelector('[role="option"]') as HTMLElement).click();
		clickButton(manager, 'Show');
		expect(v(ctx, 0, 0)).toBe(99);
		expect(ctx.root.querySelector('[data-dialog="scenario-manager"]')).not.toBeNull();
	});

	it('builds a summary sheet and merges scenarios from another sheet', async () => {
		const ctx = setup();
		const s = ctx.session()!;
		s.setCellValue(0, 0, 0, 5);
		s.setCellInput(0, 1, 0, '=A1*2');
		s.addScenario(0, { name: 'Big', cells: 'A1', values: { A1: '50' } });
		s.addSheet('Other');
		s.addScenario(1, { name: 'Huge', cells: 'A1', values: { A1: '500' } });
		ctx.select('A1');
		void ctx.dialogs.open('scenario-manager');
		const manager = dialogEl(ctx, 'scenario-manager');
		clickButton(manager, 'Merge...');
		const merge = await vi.waitFor(() => dialogEl(ctx, 'scenario-merge'));
		clickButton(merge, 'OK');
		await vi.waitFor(() =>
			expect(s.workbook.sheets[0]!.scenarios?.list.map((x) => x.name)).toEqual(['Big', 'Huge']),
		);
		clickButton(manager, 'Summary...');
		const summary = await vi.waitFor(() => dialogEl(ctx, 'scenario-summary'));
		setValue(inputByLabel(summary, 'Result cells:'), 'A2');
		clickButton(summary, 'OK');
		await vi.waitFor(() => expect(s.workbook.sheets[0]!.name).toBe('Scenario Summary'));
		const sheet = s.workbook.sheets[0]!;
		expect(getCell(sheet, 2, 4)?.value).toBe('Big');
		expect(getCell(sheet, 2, 5)?.value).toBe('Huge');
		expect(getCell(sheet, 7, 4)?.value).toBe(100);
		expect(getCell(sheet, 7, 5)?.value).toBe(1000);
	});
});
