import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { createCalcEngine } from '../formula/index.js';
import { loadXlsx } from '../read/index.js';
import { saveXlsx } from '../write/index.js';
import { createWorkbook } from '../workbook.js';
import { createConditionalFormatEvaluator } from './cf-evaluator.js';

const native = JSON.parse(
	readFileSync(new URL('./__fixtures__/excel-databar-geometry.json', import.meta.url), 'utf8'),
) as {
	cases: {
		id: string;
		xml: string;
		axisFraction: number | null;
		bars: ({ start: number; fraction: number } | null)[];
	}[];
};
const seed = saveXlsx(createWorkbook());
const near = (actual: number, expected: number) =>
	expect(Math.abs(actual - expected)).toBeLessThan(0.015);

describe('Data-bar geometry measured from native Excel PDF vectors', () => {
	it.each(native.cases)('$id', async (recorded) => {
		const zip = await JSZip.loadAsync(await seed);
		zip.file('xl/worksheets/sheet1.xml', recorded.xml);
		const workbook = await loadXlsx(await zip.generateAsync({ type: 'uint8array' }));
		const calc = createCalcEngine(workbook);
		const evaluator = createConditionalFormatEvaluator(workbook, 0, (formula, at) =>
			calc.evaluate(formula, at),
		);
		for (let i = 0; i < recorded.bars.length; i++) {
			const view = evaluator.at(i + 1, 1)!.dataBar!;
			const printed = recorded.bars[i];
			near(view.fraction, printed?.fraction ?? 0);
			if (printed) near(view.start!, printed.start);
			if (recorded.axisFraction === null) expect(view.axis).toBeUndefined();
			else {
				near(view.axis!.fraction, recorded.axisFraction);
				expect(view.axis!.color).toBe('#FF00FF');
			}
		}
	});
	it('retains explicit legacy minimum/maximum lengths through save and reload', async () => {
		const workbook = createWorkbook();
		workbook.sheets[0]!.conditionalFormats.push({
			ranges: [{ start: { row: 0, col: 0 }, end: { row: 1, col: 0 } }],
			rules: [
				{
					type: 'dataBar',
					priority: 1,
					min: { type: 'num', value: '0' },
					max: { type: 'num', value: '100' },
					color: { rgb: '00FF00' },
					minLength: 20,
					maxLength: 80,
				},
			],
		});
		const reloaded = await loadXlsx(await saveXlsx(workbook));
		expect(reloaded.sheets[0]!.conditionalFormats[0]!.rules[0]).toMatchObject({
			minLength: 20,
			maxLength: 80,
		});
	});
});
