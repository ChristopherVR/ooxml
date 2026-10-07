import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { createCalcEngine } from '../formula/index';
import { getCell } from '../cells';
import { loadXlsx } from '../read/index';
import { saveXlsx } from '../write/index';
import { createWorkbook } from '../workbook';
import { createConditionalFormatEvaluator } from './cf-evaluator';

interface NativeGeometry {
	cases: {
		id: string;
		xml: string;
		styles?: string;
		readingOrder?: number;
		axisFraction: number | null;
		bars: ({ start: number; fraction: number } | null)[];
	}[];
}
const native = [
	'excel-databar-geometry.json',
	'excel-databar-lengths.json',
	'excel-databar-context.json',
].flatMap(
	(name) =>
		(
			JSON.parse(
				readFileSync(new URL(`./__fixtures__/${name}`, import.meta.url), 'utf8'),
			) as NativeGeometry
		).cases,
);
const seed = saveXlsx(createWorkbook());
const near = (actual: number, expected: number) =>
	expect(Math.abs(actual - expected)).toBeLessThan(0.015);

describe('Data-bar geometry measured from native Excel PDF vectors', () => {
	it.each(native)('$id', async (recorded) => {
		const zip = await JSZip.loadAsync(await seed);
		zip.file('xl/worksheets/sheet1.xml', recorded.xml);
		if (recorded.styles) zip.file('xl/styles.xml', recorded.styles);
		const workbook = await loadXlsx(await zip.generateAsync({ type: 'uint8array' }));
		if (recorded.styles) {
			const style = workbook.styles[getCell(workbook.sheets[0]!, 1, 1)?.styleId ?? 0]!;
			expect(style.alignment?.readingOrder ?? 0).toBe(
				recorded.readingOrder === -5004 ? 2 : recorded.readingOrder === -5003 ? 1 : 0,
			);
		}
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
