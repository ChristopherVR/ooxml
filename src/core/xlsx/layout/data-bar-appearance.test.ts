import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { createCalcEngine } from '../formula/index';
import { loadXlsx } from '../read/index';
import { saveXlsx } from '../write/index';
import { createWorkbook } from '../workbook';
import { createConditionalFormatEvaluator } from './cf-evaluator';

const native = JSON.parse(
	readFileSync(
		new URL('../edit/__fixtures__/excel-databar-clipboard.json', import.meta.url),
		'utf8',
	),
) as {
	cases: { variant: number; before: string }[];
};
const seed = saveXlsx(createWorkbook());

describe('Data-bar appearance from native Excel 16.0 workbooks', () => {
	it.each([
		{
			variant: 0,
			gradient: true,
			direction: 'leftToRight',
			negative: '#0000FF',
			border: '#FF0000',
			negativeBorder: '#FFFF00',
			hideValue: false,
		},
		{
			variant: 1,
			gradient: false,
			direction: 'rightToLeft',
			negative: '#0000FF',
			border: '#FF0000',
			negativeBorder: '#FFFF00',
			hideValue: true,
		},
		{
			variant: 2,
			gradient: false,
			direction: 'leftToRight',
			negative: '#00FF00',
			border: undefined,
			negativeBorder: undefined,
			hideValue: false,
		},
	])('resolves variant $variant through the real calculation engine', async (expected) => {
		const zip = await JSZip.loadAsync(await seed);
		zip.file(
			'xl/worksheets/sheet1.xml',
			native.cases.find((c) => c.variant === expected.variant)!.before,
		);
		const workbook = await loadXlsx(await zip.generateAsync({ type: 'uint8array' }));
		const calc = createCalcEngine(workbook);
		const evaluator = createConditionalFormatEvaluator(workbook, 0, (formula, at) =>
			calc.evaluate(formula, at),
		);
		const positive = evaluator.at(4, 0)!;
		const negative = evaluator.at(1, 0)!;
		expect(positive.dataBar).toMatchObject({
			color: '#00FF00',
			gradient: expected.gradient,
			direction: expected.direction,
		});
		expect(negative.dataBar).toMatchObject({
			color: expected.negative,
			negative: true,
			gradient: expected.gradient,
			direction:
				expected.variant === 2
					? expected.direction
					: expected.direction === 'leftToRight'
						? 'rightToLeft'
						: 'leftToRight',
		});
		expect(positive.dataBar?.borderColor).toBe(expected.border);
		expect(negative.dataBar?.borderColor).toBe(expected.negativeBorder);
		expect(Boolean(positive.hideValue)).toBe(expected.hideValue);
		expect(Boolean(negative.hideValue)).toBe(expected.hideValue);
	});
});
