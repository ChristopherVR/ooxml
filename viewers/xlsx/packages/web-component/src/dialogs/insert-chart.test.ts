// @vitest-environment jsdom
import { createWorkbook, loadXlsx, saveXlsx } from '@christophervr/xlsx-core';
import { afterEach, describe, expect, it } from 'vitest';
import {
	clickButton,
	createTestContext,
	dialogEl,
	inputByLabel,
	pressKey,
	setValue,
} from '../commands/test-support.js';
import { buildChart } from './insert-chart-model.js';
import { registerToolDialogs } from './register-tools.js';

afterEach(() => (document.body.innerHTML = ''));

function setup() {
	const ctx = createTestContext(createWorkbook());
	registerToolDialogs(ctx);
	const s = ctx.session()!;
	s.setRangeValues(0, { row: 0, col: 0 }, [
		['Month', 'Sales', 'Cost'],
		['Jan', 10, 4],
		['Feb', 20, 8],
		['Mar', 30, 9],
	]);
	ctx.select('A1');
	return ctx;
}

describe('buildChart', () => {
	it('names series from the header row and labels categories from the first column', () => {
		const ctx = setup();
		const chart = buildChart(
			ctx.workbook()!,
			0,
			{ start: { row: 0, col: 0 }, end: { row: 3, col: 2 } },
			'column',
		);
		expect(chart.series).toHaveLength(2);
		expect(chart.series[0]).toMatchObject({
			name: 'Sales',
			nameRef: 'Sheet1!$B$1',
			valuesRef: 'Sheet1!$B$2:$B$4',
			categoriesRef: 'Sheet1!$A$2:$A$4',
			values: [10, 20, 30],
			categories: ['Jan', 'Feb', 'Mar'],
		});
		expect(chart.anchor.from).toMatchObject({ row: 0, col: 4 });
	});
});

describe('Insert Chart dialog', () => {
	it('previews the selection and adds a chart of the chosen type', async () => {
		const ctx = setup();
		const result = ctx.dialogs.open('insert-chart', { type: 'column' });
		const dialog = dialogEl(ctx, 'insert-chart');
		expect(dialog.querySelector('.xve-preview svg')).not.toBeNull();
		dialog.querySelector<HTMLButtonElement>('[data-type="line"]')!.click();
		expect(dialog.querySelector('[data-type="line"]')!.getAttribute('aria-pressed')).toBe('true');
		setValue(inputByLabel(dialog, 'Chart title'), 'Q1');
		clickButton(dialog, 'OK');
		await result;
		const drawings = ctx.workbook()!.sheets[0]!.drawings;
		expect(drawings).toHaveLength(1);
		expect(drawings[0]).toMatchObject({ kind: 'chart', chartType: 'line', title: 'Q1' });
		ctx.session()!.undo();
		expect(ctx.workbook()!.sheets[0]!.drawings).toHaveLength(0);
	});

	it('Cancel and Escape add nothing', async () => {
		const ctx = setup();
		const a = ctx.dialogs.open('insert-chart');
		clickButton(dialogEl(ctx, 'insert-chart'), 'Cancel');
		expect(await a).toBeUndefined();
		const b = ctx.dialogs.open('insert-chart');
		pressKey(dialogEl(ctx, 'insert-chart'), 'Escape');
		expect(await b).toBeUndefined();
		expect(ctx.workbook()!.sheets[0]!.drawings).toHaveLength(0);
	});

	it('changes the type of the chart under the active cell', async () => {
		const ctx = setup();
		const s = ctx.session()!;
		s.addChart(
			0,
			buildChart(
				ctx.workbook()!,
				0,
				{ start: { row: 0, col: 0 }, end: { row: 3, col: 2 } },
				'column',
			),
		);
		ctx.selection.set({ drawing: 0 });
		const result = ctx.dialogs.open('insert-chart', { change: true });
		const dialog = dialogEl(ctx, 'insert-chart');
		dialog.querySelector<HTMLButtonElement>('[data-type="pie"]')!.click();
		clickButton(dialog, 'OK');
		await result;
		expect(ctx.workbook()!.sheets[0]!.drawings).toHaveLength(1);
		expect(ctx.workbook()!.sheets[0]!.drawings[0]).toMatchObject({ chartType: 'pie' });
	});
});

describe('drawings round trip', () => {
	it('keeps an added chart and picture through save and load', async () => {
		const ctx = setup();
		const s = ctx.session()!;
		s.addChart(
			0,
			buildChart(
				ctx.workbook()!,
				0,
				{ start: { row: 0, col: 0 }, end: { row: 3, col: 2 } },
				'column',
			),
		);
		const png = Uint8Array.from(
			atob(
				'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==',
			),
			(c) => c.charCodeAt(0),
		);
		s.addImage(0, png, 'image/png', {
			from: { row: 6, col: 0, rowOffset: 0, colOffset: 0 },
			ext: { cx: 9525, cy: 9525 },
		});
		const loaded = await loadXlsx(await saveXlsx(ctx.workbook()!));
		const kinds = loaded.sheets[0]!.drawings.map((d) => d.kind).sort();
		expect(kinds).toEqual(['chart', 'image']);
	});
});
