// @vitest-environment jsdom
// Every string the tool dialogs translate has an entry in TOOLS_STRINGS.
import { createWorkbook } from '@christophervr/xlsx-core';
import { afterEach, describe, expect, it } from 'vitest';
import { TOOLS_STRINGS } from '../commands/i18n/tools.js';
import { createTestContext } from '../commands/test-support.js';
import { buildChart } from './insert-chart-model.js';
import { registerToolDialogs } from './register-tools.js';

afterEach(() => (document.body.innerHTML = ''));

const FRAME = new Set(['OK', 'Cancel', 'Close']);

describe('tool dialog strings', () => {
	it('are all translated', () => {
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
		const names = [
			'insert-chart',
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
		];
		for (const name of names) void ctx.dialogs.open(name);
		ctx.select('E2');
		void ctx.dialogs.open('insert-chart', { change: true });
		const missing = [...used].filter((key) => !FRAME.has(key) && !(key in TOOLS_STRINGS));
		expect(missing).toEqual([]);
		for (const [key, texts] of Object.entries(TOOLS_STRINGS)) {
			expect(
				texts.every((t) => t.length > 0),
				key,
			).toBe(true);
			expect(texts.join('').includes(String.fromCharCode(0x2014)), key).toBe(false);
		}
	});
});
