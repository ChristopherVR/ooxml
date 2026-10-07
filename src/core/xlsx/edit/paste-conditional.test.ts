import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { parseRange, formatRange, rangeContains } from '../address.js';
import { getCell } from '../cells.js';
import type { ConditionalRule, Fill } from '../model.js';
import { createWorkbook } from '../workbook.js';
import { createCalcEngine } from '../formula/index.js';
import { createConditionalFormatEvaluator } from '../layout/cf-evaluator.js';
import { createEditSession } from './session.js';
import type { PasteOptions } from './types.js';
import { saveXlsx } from '../write/index.js';
import { loadXlsx } from '../read/index.js';

interface NativeRule {
	type: number;
	formula1: string;
	operator: number | null;
	priority: number;
	ranges: string;
	fill: number;
	bold: boolean;
	stopIfTrue: boolean;
}
interface NativeCase extends Required<PasteOptions> {
	source: string;
	rules: NativeRule[];
	cells: { row: number; col: number; value: number | null; fill: number; priorities: number[] }[];
}
const fixture = JSON.parse(
	readFileSync(new URL('./__fixtures__/excel-paste-cf.json', import.meta.url), 'utf8'),
) as { cases: NativeCase[]; colorScale: { cells: { ref: string; value: number; fill: number }[] } };
const range = (ref: string) => parseRange(ref)!;
const color = (rgb: string) => {
	const n = parseInt(rgb.slice(-6), 16);
	return ((n & 255) << 16) | (n & 0xff00) | ((n >> 16) & 255);
};
const fillColor = (fill: Fill | undefined) =>
	color(fill?.type === 'pattern' ? (fill.fgColor?.rgb ?? 'FFFFFF') : 'FFFFFF');
const setup = () => {
	const workbook = createWorkbook();
	const sheet = workbook.sheets[0]!;
	const session = createEditSession(workbook);
	session.setCellValue(0, 0, 0, 2);
	session.setCellValue(0, 0, 1, 4);
	session.setCellInput(0, 1, 0, '=B2+$J$1');
	session.setCellValue(0, 0, 9, 3);
	for (const ref of ['D4', 'E4', 'D5', 'E5']) {
		const at = range(ref).start;
		session.setCellValue(0, at.row, at.col, 9);
	}
	const expression = (
		formula: string,
		priority: number,
		rgb: string,
		bold = false,
	): ConditionalRule => ({
		type: 'expression',
		formula,
		priority,
		stopIfTrue: false,
		style: { fill: { type: 'pattern', pattern: 'solid', fgColor: { rgb } }, font: { bold } },
	});
	sheet.conditionalFormats = [
		{ ranges: [range('A1:B2')], rules: [expression('A1>0', 1, 'FF0000', true)] },
		{
			ranges: [range('A1:B1')],
			rules: [
				{
					type: 'cellIs',
					operator: 'greaterThan',
					formulas: ['1'],
					priority: 2,
					stopIfTrue: true,
					style: {
						fill: { type: 'pattern', pattern: 'solid', fgColor: { rgb: 'FFFF00' } },
						font: { bold: false },
					},
				},
			],
		},
		{ ranges: [range('D3:F6')], rules: [expression('D3<0', 3, '0000FF')] },
		{ ranges: [range('D4:E5')], rules: [expression('D4>5', 4, '00FF00')] },
	];
	return { workbook, sheet, session };
};

describe('Conditional format clipboard semantics recorded in Microsoft Excel', () => {
	it('uses the entire tiled destination for native color-scale statistics', () => {
		const workbook = createWorkbook();
		const sheet = workbook.sheets[0]!;
		const session = createEditSession(workbook);
		session.setCellValue(0, 0, 0, 0);
		session.setCellValue(0, 0, 1, 10);
		sheet.conditionalFormats = [
			{
				ranges: [range('A1:B1')],
				rules: [
					{
						type: 'colorScale',
						priority: 1,
						thresholds: [{ type: 'min' }, { type: 'max' }],
						colors: [{ rgb: 'FFFF0000' }, { rgb: 'FF00FF00' }],
					},
				],
			},
		];
		for (const cell of fixture.colorScale.cells) {
			const at = range(cell.ref).start;
			session.setCellValue(0, at.row, at.col, cell.value);
		}
		session.paste(0, range('D4:G5'), session.copy(0, range('A1:B1')), 'formats');
		const cf = createConditionalFormatEvaluator(workbook, 0, () => null);
		for (const cell of fixture.colorScale.cells) {
			const at = range(cell.ref).start;
			expect(color(cf.at(at.row, at.col)?.colorScale ?? 'FFFFFF')).toBe(cell.fill);
		}
	});
	it.each(fixture.cases)(
		'$source $mode transpose=$transpose skip=$skipBlanks op=$operation',
		(native) => {
			const { workbook, sheet, session } = setup();
			const before = structuredClone(sheet.conditionalFormats);
			session.paste(0, range('D4'), session.copy(0, range(native.source)), native);
			const normalized = sheet.conditionalFormats
				.flatMap((format) =>
					format.rules.map((rule) => ({
						type: rule.type === 'cellIs' ? 1 : 2,
						formula1: `=${rule.type === 'expression' ? rule.formula : rule.type === 'cellIs' ? rule.formulas[0] : ''}`,
						operator: rule.type === 'cellIs' ? 5 : null,
						priority: rule.priority,
						ranges: format.ranges
							.map((r) => formatRange(r).replace(/([A-Z]+)(\d+)/g, '$$$1$$$2'))
							.join(','),
						fill: 'style' in rule ? fillColor(rule.style.fill) : 0,
						bold: 'style' in rule ? (rule.style.font?.bold ?? false) : false,
						stopIfTrue: rule.stopIfTrue ?? false,
					})),
				)
				.sort((a, b) => a.priority - b.priority);
			expect(normalized).toEqual(native.rules);
			const engine = createCalcEngine(workbook);
			const evaluator = createConditionalFormatEvaluator(workbook, 0, (f, at) =>
				engine.evaluate(f, at),
			);
			for (const cell of native.cells) {
				expect(getCell(sheet, cell.row, cell.col)?.value ?? null).toBe(cell.value);
				expect(fillColor(evaluator.at(cell.row, cell.col)?.style?.fill)).toBe(cell.fill);
				expect(
					sheet.conditionalFormats
						.filter((f) => f.ranges.some((r) => rangeContains(r, cell)))
						.flatMap((f) => f.rules.map((r) => r.priority))
						.sort((a, b) => a - b),
				).toEqual(cell.priorities);
			}
			const after = structuredClone(sheet.conditionalFormats);
			session.undo();
			expect(sheet.conditionalFormats).toEqual(before);
			session.redo();
			expect(sheet.conditionalFormats).toEqual(after);
		},
	);
	it('groups repeated blocks into one rule, as Excel does, and undoes the whole selection', async () => {
		const { workbook, sheet, session } = setup();
		sheet.conditionalFormats = sheet.conditionalFormats.slice(0, 2);
		sheet.conditionalFormats[1]!.ranges = [range('B1')];
		const rule = sheet.conditionalFormats[1]!.rules[0]!;
		if (rule.type === 'cellIs') rule.formulas = ['A1'];
		const before = structuredClone(sheet.conditionalFormats);
		session.paste(0, range('D4:G7'), session.copy(0, range('A1:B2')));
		const added = sheet.conditionalFormats.filter((f) => f.rules[0]!.priority <= 2);
		expect(added.map((f) => f.ranges.map(formatRange))).toEqual([
			['D4:G7'],
			['E4', 'E6', 'G4', 'G6'],
		]);
		expect(added[1]!.rules[0]).toMatchObject({ formulas: ['D4'], priority: 2 });
		const after = structuredClone(sheet.conditionalFormats);
		const loaded = await loadXlsx(await saveXlsx(workbook));
		expect(loaded.sheets[0]!.conditionalFormats.map((f) => f.ranges)).toEqual(
			after.map((f) => f.ranges),
		);
		session.undo();
		expect(sheet.conditionalFormats).toEqual(before);
		session.redo();
		expect(sheet.conditionalFormats).toEqual(after);
	});
	it('rebases surviving relative rules when clearing their original anchor', () => {
		const { sheet, session } = setup();
		sheet.conditionalFormats = sheet.conditionalFormats.slice(0, 1);
		session.clearConditionalFormats(0, range('A1:B1'));
		expect(sheet.conditionalFormats[0]).toMatchObject({
			ranges: [range('A2:B2')],
			rules: [{ formula: 'A2>0', priority: 1 }],
		});
		session.undo();
		expect(sheet.conditionalFormats[0]!.rules[0]).toMatchObject({ formula: 'A1>0' });
	});
	it('moves cut rules across sheets while retaining source rules outside the move', () => {
		const { workbook, sheet, session } = setup();
		sheet.conditionalFormats = sheet.conditionalFormats.slice(0, 1);
		const before = structuredClone(sheet.conditionalFormats);
		const target = session.addSheet('Target');
		session.paste(target, range('D4'), session.cut(0, range('A1:B1')));
		expect(sheet.conditionalFormats[0]).toMatchObject({
			ranges: [range('A2:B2')],
			rules: [{ formula: 'A2>0' }],
		});
		expect(workbook.sheets[target]!.conditionalFormats[0]).toMatchObject({
			ranges: [range('D4:E4')],
			rules: [{ formula: 'D4>0', priority: 1 }],
		});
		session.undo();
		expect(sheet.conditionalFormats).toEqual(before);
		expect(workbook.sheets[target]!.conditionalFormats).toEqual([]);
	});
	it('translates formula thresholds in visual rules through cross-workbook snapshots', async () => {
		const { sheet, session } = setup();
		sheet.conditionalFormats = [
			{
				ranges: [range('A1:B2')],
				rules: [
					{
						type: 'colorScale',
						priority: 1,
						thresholds: [
							{ type: 'formula', value: 'A1' },
							{ type: 'formula', value: '$J$1' },
						],
						colors: [{ rgb: 'FFFF0000' }, { rgb: 'FF00FF00' }],
					},
					{
						type: 'dataBar',
						priority: 2,
						extensionId: '{00000000-0000-0000-0000-000000000001}',
						min: { type: 'formula', value: 'A1' },
						max: { type: 'formula', value: '$J$1' },
						color: { rgb: 'FF0000FF' },
					},
					{
						type: 'iconSet',
						priority: 3,
						iconSet: '3TrafficLights1',
						thresholds: [
							{ type: 'formula', value: 'A1' },
							{ type: 'percent', value: '33' },
							{ type: 'percent', value: '67' },
						],
					},
				],
			},
		];
		const payload = session.copy(0, range('B2'));
		sheet.conditionalFormats = [];
		const target = createWorkbook();
		createEditSession(target).paste(0, range('D4'), payload, 'formats');
		const rules = target.sheets[0]!.conditionalFormats[0]!.rules;
		expect(rules[0]).toMatchObject({ thresholds: [{ value: 'D4' }, { value: '$J$1' }] });
		expect(rules[1]).toMatchObject({ min: { value: 'D4' }, max: { value: '$J$1' } });
		expect(rules[1]).not.toHaveProperty('extensionId');
		expect(rules[2]).toMatchObject({
			thresholds: [{ value: 'D4' }, { value: '33' }, { value: '67' }],
		});
		const loaded = await loadXlsx(await saveXlsx(target));
		expect(loaded.sheets[0]!.conditionalFormats[0]!.rules).toMatchObject(rules);
	});
});
