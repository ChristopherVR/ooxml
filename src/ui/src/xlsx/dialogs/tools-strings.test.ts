// @vitest-environment jsdom
// Every string the tool dialogs translate has an entry in the editor dictionary.
import { createWorkbook } from 'ooxml-core/xlsx';
import { afterEach, describe, expect, it } from 'vitest';
import { STRINGS } from 'ooxml-core/xlsx/ui';
import { clickButton, createTestContext, waitForDialog } from '../commands/test-support';
import { buildChart } from 'ooxml-core/xlsx/ui';
import { registerToolDialogs } from './register-tools';

afterEach(() => (document.body.innerHTML = ''));

const FRAME = new Set(['OK', 'Cancel', 'Close']);

describe('tool dialog strings', () => {
	it('are all translated', async () => {
		const ctx = createTestContext(createWorkbook());
		const used = new Set<string>();
		const base = ctx.t;
		ctx.t = (key, vars) => {
			used.add(key);
			return base(key, vars);
		};
		registerToolDialogs(ctx);
		const s = ctx.session()!;
		s.setRangeValues(0, { row: 0, col: 0 }, [
			['a', 1],
			['b', 2],
		]);
		s.addChart(
			0,
			buildChart(
				ctx.workbook()!,
				0,
				{ start: { row: 0, col: 0 }, end: { row: 1, col: 1 } },
				'column',
			),
		);
		const newChart = ctx.dialogs.open('insert-chart');
		clickButton(await waitForDialog(ctx, 'insert-chart'), 'Cancel');
		await newChart;
		const names = [
			'paste-special',
			'page-setup',
			'zoom',
			'move-copy-sheet',
			'protect-sheet',
			'insert-cells',
			'delete-cells',
			'remove-duplicates',
			'text-to-columns',
			'symbol',
			'fill-series',
			'create-table',
			'subtotal',
			'consolidate',
			'advanced-filter',
		];
		for (const name of names) void ctx.dialogs.open(name);
		ctx.select('E2');
		const changeChart = ctx.dialogs.open('insert-chart', { change: true });
		clickButton(await waitForDialog(ctx, 'insert-chart'), 'Cancel');
		await changeChart;
		const missing = [...used].filter((key) => !FRAME.has(key) && !(key in STRINGS.en));
		expect(missing).toEqual([]);
	});
});
