import { describe, expect, it } from 'vitest';
import { putCell } from '../cells.js';
import type { ConditionalRule, DifferentialStyle, Workbook } from '../model.js';
import { createWorkbook } from '../workbook.js';
import { createConditionalFormatEvaluator } from './cf-evaluator.js';

const RED: DifferentialStyle = {
	fill: { type: 'pattern', pattern: 'solid', bgColor: { rgb: 'FFFF0000' } },
};
const BOLD = { font: { bold: true } };

/** Column A (rows 0..9) holds 1..10 under `rules`. */
function sheetWith(rules: ConditionalRule[]): Workbook {
	const wb = createWorkbook();
	const sheet = wb.sheets[0]!;
	for (let i = 0; i < 10; i++) putCell(sheet, i, 0, { value: i + 1 });
	sheet.conditionalFormats.push({
		ranges: [{ start: { row: 0, col: 0 }, end: { row: 9, col: 0 } }],
		rules,
	});
	return wb;
}

const evaluate = () => true;

describe('Stop If True on every rule type', () => {
	// Excel 16 (x3.ps1/x5.ps1): top 3 with Stop If True, then "=TRUE" bold: A10 is red and not
	// bold, A1 is bold.
	it('stops after a matching top-10 rule', () => {
		const wb = sheetWith([
			{ type: 'top10', rank: 3, style: RED, priority: 1, stopIfTrue: true },
			{ type: 'expression', formula: 'TRUE', style: BOLD, priority: 2 },
		]);
		const ev = createConditionalFormatEvaluator(wb, 0, evaluate);
		expect(ev.at(9, 0)?.style?.fill).toBeDefined();
		expect(ev.at(9, 0)?.style?.font?.bold).toBeUndefined();
		expect(ev.at(0, 0)?.style?.font?.bold).toBe(true);
	});

	// Excel 16, files saved with each rule first and stopIfTrue="1", then "=TRUE" bold: above
	// average holds back the bold rule on A9; a data bar, an icon set and a colour scale do not.
	it('stops after an average rule but not after graphic rules', () => {
		const later: ConditionalRule = {
			type: 'expression',
			formula: 'TRUE',
			style: BOLD,
			priority: 9,
		};
		const bold = (rule: ConditionalRule) =>
			createConditionalFormatEvaluator(sheetWith([rule, later]), 0, evaluate).at(8, 0)?.style?.font
				?.bold;
		expect(
			bold({ type: 'aboveAverage', style: RED, priority: 1, stopIfTrue: true }),
		).toBeUndefined();
		expect(
			bold({ type: 'notContainsBlanks', style: RED, priority: 1, stopIfTrue: true }),
		).toBeUndefined();
		const graphic: ConditionalRule[] = [
			{
				type: 'colorScale',
				thresholds: [{ type: 'min' }, { type: 'max' }],
				colors: [{ rgb: 'FFFFFFFF' }, { rgb: 'FF000000' }],
				priority: 1,
				stopIfTrue: true,
			},
			{
				type: 'dataBar',
				min: { type: 'min' },
				max: { type: 'max' },
				color: { rgb: 'FF638EC6' },
				priority: 1,
				stopIfTrue: true,
			},
			{
				type: 'iconSet',
				iconSet: '3Arrows',
				thresholds: [
					{ type: 'percent', value: '0' },
					{ type: 'percent', value: '33' },
				],
				priority: 1,
				stopIfTrue: true,
			},
		];
		for (const rule of graphic) expect(bold(rule), rule.type).toBe(true);
	});
});

describe('icon set thresholds with gte="0"', () => {
	// Excel 16 (x3.ps1): 3 arrows with "> 4" and "> 7": 4 gets the first icon, 5 and 7 the
	// second, 8 the third.
	it('uses > instead of >= when gte is false', () => {
		const wb = sheetWith([
			{
				type: 'iconSet',
				iconSet: '3Arrows',
				thresholds: [
					{ type: 'percent', value: '0' },
					{ type: 'num', value: '4', gte: false },
					{ type: 'num', value: '7', gte: false },
				],
				priority: 1,
			},
		]);
		const ev = createConditionalFormatEvaluator(wb, 0, evaluate);
		expect([3, 4, 6, 7].map((r) => ev.at(r, 0)?.icon?.index)).toEqual([0, 1, 1, 2]);
	});

	it('keeps >= by default', () => {
		const wb = sheetWith([
			{
				type: 'iconSet',
				iconSet: '3Arrows',
				thresholds: [
					{ type: 'percent', value: '0' },
					{ type: 'num', value: '4' },
					{ type: 'num', value: '7' },
				],
				priority: 1,
			},
		]);
		const ev = createConditionalFormatEvaluator(wb, 0, evaluate);
		expect([2, 3, 6].map((r) => ev.at(r, 0)?.icon?.index)).toEqual([0, 1, 2]);
	});
});
