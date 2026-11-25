import { afterEach, expect, it, vi } from 'vitest';
import { getCell, parseRange, pivotChartButtons } from 'ooxml-core/xlsx';
import { installCommands } from '../commands';
import {
	clickButton,
	createTestContext,
	dialogEl,
	inputByLabel,
	setValue,
} from '../commands/test-support';

afterEach(() => document.body.replaceChildren());

const setup = () => {
	const ctx = createTestContext();
	const session = ctx.session()!;
	[
		['Category', 'Region', 'Amount'],
		['Fruit', 'East', 10],
		['Vegetables', 'West', 20],
		['Fruit', 'West', 30],
	].forEach((row, r) => row.forEach((value, c) => session.setCellInput(0, r, c, String(value))));
	ctx.select('A1:C4');
	return { ctx, session, stop: installCommands(ctx) };
};

it('changes the recommendation preview and inserts the chosen report as one undo step', async () => {
	const { ctx, session, stop } = setup();
	try {
		const done = ctx.commands.run('insert.recommended-pivot-tables');
		const dialog = await vi.waitFor(() => dialogEl(ctx, 'recommended-pivot-tables'));
		expect(dialog.querySelector('table')?.textContent).toContain('Fruit');
		setValue(inputByLabel<HTMLSelectElement>(dialog, 'PivotTable'), '1');
		expect(dialog.querySelector('table')?.textContent).toContain('West50');
		clickButton(dialog, 'Insert');
		expect(await done).toBe(true);
		expect(session.workbook.sheets[1]?.pivotTables?.[0]?.rowField).toBe('Region');
		expect(getCell(session.workbook.sheets[1]!, 4, 1)?.value).toBe(50);
		expect(session.undo()).toBe(true);
		expect(session.workbook.sheets).toHaveLength(1);
	} finally {
		stop();
	}
});

it('creates from a named table, edits the shared field filter and hides the chart buttons', async () => {
	const { ctx, session, stop } = setup();
	try {
		const table = session.createTable(0, parseRange('A1:C4')!, true);
		const done = ctx.commands.run('insert.pivot-chart');
		const dialog = await vi.waitFor(() => dialogEl(ctx, 'create-pivot-chart'));
		expect(inputByLabel(dialog, 'Source Range:').value).toBe(table.name);
		clickButton(dialog, 'OK');
		expect(await done).toBe(true);
		let chart = session.workbook.sheets[1]?.drawings[0];
		if (chart?.kind !== 'chart') throw new Error('Missing chart');
		expect(chart.series[0]?.values).toEqual([40, 20]);
		const filter = ctx.dialogs.open('pivot-chart-field', {
			chart,
			role: 'axis',
			field: 'Category',
		});
		const field = await vi.waitFor(() => dialogEl(ctx, 'pivot-field-filter'));
		const item = [...field.querySelectorAll('label')].find(
			(label) => label.textContent === 'Fruit',
		);
		item?.querySelector('input')?.click();
		clickButton(field, 'OK');
		expect(await filter).toBe(true);
		chart = session.workbook.sheets[1]?.drawings[0];
		if (chart?.kind !== 'chart') throw new Error('Missing chart');
		expect(chart.series[0]?.values).toEqual([20]);
		expect(await ctx.commands.run('chart.pivot-field-buttons')).toBe(true);
		chart = session.workbook.sheets[1]?.drawings[0];
		if (chart?.kind !== 'chart') throw new Error('Missing chart');
		expect(pivotChartButtons(chart).size).toBe(0);
	} finally {
		stop();
	}
});
