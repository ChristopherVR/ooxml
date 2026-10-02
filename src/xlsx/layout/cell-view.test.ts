import { describe, expect, it } from 'vitest';
import { putCell } from '../cells.js';
import type { CellStyle, Workbook } from '../model.js';
import { internStyle } from '../styles.js';
import { createWorkbook, defaultCellStyle } from '../workbook.js';
import { cellView, effectiveStyleId, generalAlignment, rotationDegrees } from './cell-view.js';
import { BORDER_STYLES, fillView } from './style-view.js';
import type { ConditionalFormatEvaluator } from './types.js';

function book(): Workbook {
	return createWorkbook();
}

function style(workbook: Workbook, patch: Partial<CellStyle>): number {
	return internStyle(workbook, { ...defaultCellStyle(), ...patch });
}

describe('cellView text and alignment', () => {
	it('formats numbers, booleans, errors and text with General alignment', () => {
		const wb = book();
		const sheet = wb.sheets[0]!;
		putCell(sheet, 0, 0, { value: 1234.5 });
		putCell(sheet, 0, 1, { value: true });
		putCell(sheet, 0, 2, { value: { error: '#DIV/0!' } });
		putCell(sheet, 0, 3, { value: 'hello' });
		const n = cellView(wb, 0, 0, 0);
		expect(n.text).toBe('1234.5');
		expect(n.hAlign).toBe('right');
		expect(n.isNumber).toBe(true);
		expect(n.overflow).toBe(false);
		const b = cellView(wb, 0, 0, 1);
		expect(b.text).toBe('TRUE');
		expect(b.hAlign).toBe('center');
		const e = cellView(wb, 0, 0, 2);
		expect(e.text).toBe('#DIV/0!');
		expect(e.isError).toBe(true);
		expect(e.hAlign).toBe('center');
		const t = cellView(wb, 0, 0, 3);
		expect(t.text).toBe('hello');
		expect(t.hAlign).toBe('left');
		expect(t.overflow).toBe(true);
		expect(t.vAlign).toBe('bottom');
	});

	it('applies number formats and their colours', () => {
		const wb = book();
		const sheet = wb.sheets[0]!;
		putCell(sheet, 0, 0, { value: -5, styleId: style(wb, { numFmt: '0.00;[Red]-0.00' }) });
		putCell(sheet, 0, 1, { value: 45292, styleId: style(wb, { numFmt: 'yyyy-mm-dd' }) });
		const v = cellView(wb, 0, 0, 0);
		expect(v.text).toBe('-5.00');
		expect(v.font.color).toBe('#FF0000');
		expect(cellView(wb, 0, 0, 1).text).toBe('2024-01-01');
	});

	it('renders empty cells and hides zeros when the sheet says so', () => {
		const wb = book();
		const sheet = wb.sheets[0]!;
		putCell(sheet, 0, 0, { value: 0 });
		expect(cellView(wb, 0, 5, 5).text).toBe('');
		expect(cellView(wb, 0, 0, 0).text).toBe('0');
		sheet.view.showZeros = false;
		expect(cellView(wb, 0, 0, 0).text).toBe('');
	});

	it('maps explicit alignment, wrap, shrink, indent and rotation', () => {
		const wb = book();
		const id = style(wb, {
			alignment: {
				horizontal: 'right',
				vertical: 'top',
				indent: 2,
				textRotation: 45,
				shrinkToFit: true,
			},
		});
		putCell(wb.sheets[0]!, 0, 0, { value: 'x', styleId: id });
		const v = cellView(wb, 0, 0, 0);
		expect(v.hAlign).toBe('right');
		expect(v.vAlign).toBe('top');
		expect(v.indentPx).toBe(18);
		expect(v.rotation).toBe(45);
		expect(v.shrink).toBe(true);
		expect(v.overflow).toBe(false);
		const wrapped = style(wb, { alignment: { wrapText: true, horizontal: 'center' } });
		putCell(wb.sheets[0]!, 0, 1, { value: 'y', styleId: wrapped });
		const w = cellView(wb, 0, 0, 1);
		expect(w.wrap).toBe(true);
		expect(w.overflow).toBe(false);
		expect(w.indentPx).toBe(0);
	});

	it('converts text rotation like Excel', () => {
		expect(rotationDegrees(undefined)).toEqual({ rotation: 0, vertical: false });
		expect(rotationDegrees(90)).toEqual({ rotation: 90, vertical: false });
		expect(rotationDegrees(135)).toEqual({ rotation: -45, vertical: false });
		expect(rotationDegrees(180)).toEqual({ rotation: -90, vertical: false });
		expect(rotationDegrees(255)).toEqual({ rotation: 0, vertical: true });
	});

	it('decides General alignment by value type', () => {
		expect(generalAlignment(1)).toBe('right');
		expect(generalAlignment('a')).toBe('left');
		expect(generalAlignment(null)).toBe('left');
		expect(generalAlignment(false)).toBe('center');
	});
});

describe('cellView fonts, fills and borders', () => {
	it('resolves the default font through the theme', () => {
		const wb = book();
		const v = cellView(wb, 0, 0, 0);
		expect(v.font.family).toBe('Calibri');
		expect(v.font.sizePx).toBeCloseTo(14.667, 2);
		expect(v.font.color).toBe('#000000');
		expect(v.font.bold).toBe(false);
	});

	it('maps font properties', () => {
		const wb = book();
		const id = style(wb, {
			font: {
				name: 'Arial',
				size: 12,
				bold: true,
				italic: true,
				underline: 'doubleAccounting',
				strike: true,
				vertAlign: 'superscript',
				color: { theme: 4 },
			},
		});
		putCell(wb.sheets[0]!, 0, 0, { value: 'a', styleId: id });
		const f = cellView(wb, 0, 0, 0).font;
		expect(f).toEqual({
			family: 'Arial',
			sizePx: 16,
			bold: true,
			italic: true,
			strike: true,
			underline: 'double',
			vertAlign: 'super',
			color: '#4472C4',
		});
	});

	it('uses the major theme font for headings', () => {
		const wb = book();
		putCell(wb.sheets[0]!, 0, 0, {
			value: 'a',
			styleId: style(wb, { font: { scheme: 'major', size: 18 } }),
		});
		expect(cellView(wb, 0, 0, 0).font.family).toBe('Calibri Light');
	});

	it('maps solid, pattern and gradient fills', () => {
		const wb = book();
		const theme = wb.theme;
		expect(fillView({ type: 'pattern', pattern: 'none' }, theme)).toBeUndefined();
		expect(
			fillView({ type: 'pattern', pattern: 'solid', fgColor: { rgb: 'FFFFFF00' } }, theme),
		).toEqual({ background: '#FFFF00' });
		expect(fillView({ type: 'pattern', pattern: 'gray125' }, theme)).toEqual({
			pattern: 'gray125',
			fg: '#000000',
			bg: '#FFFFFF',
		});
		expect(
			fillView(
				{
					type: 'gradient',
					gradient: 'linear',
					degree: 90,
					stops: [
						{ position: 0, color: { rgb: 'FF0000' } },
						{ position: 1, color: { theme: 4 } },
					],
				},
				theme,
			),
		).toEqual({ gradient: 'linear-gradient(180deg, #FF0000 0%, #4472C4 100%)' });
		expect(
			fillView(
				{ type: 'gradient', gradient: 'path', stops: [{ position: 0, color: { rgb: 'FFFFFF' } }] },
				theme,
			),
		).toEqual({
			gradient: 'radial-gradient(circle, #FFFFFF 0%, #FFFFFF 0%)',
		});
		expect(
			fillView({ type: 'pattern', pattern: 'solid', bgColor: { rgb: 'C6EFCE' } }, theme, true),
		).toEqual({ background: '#C6EFCE' });
	});

	it('maps every border style to a width and line style', () => {
		const wb = book();
		const id = style(wb, {
			border: {
				top: { style: 'thin' },
				bottom: { style: 'double', color: { rgb: 'FF0000' } },
				left: { style: 'mediumDashed' },
				right: { style: 'thick' },
				diagonal: { style: 'hair' },
				diagonalUp: true,
			},
		});
		putCell(wb.sheets[0]!, 0, 0, { value: 1, styleId: id });
		const b = cellView(wb, 0, 0, 0).borders;
		expect(b.top).toEqual({ widthPx: 1, style: 'solid', color: '#000000' });
		expect(b.bottom).toEqual({ widthPx: 3, style: 'double', color: '#FF0000' });
		expect(b.left).toEqual({ widthPx: 2, style: 'dashed', color: '#000000' });
		expect(b.right).toEqual({ widthPx: 3, style: 'solid', color: '#000000' });
		expect(b.diagonalUp).toEqual({ widthPx: 1, style: 'dotted', color: '#000000' });
		expect(b.diagonalDown).toBeUndefined();
		expect(Object.keys(BORDER_STYLES)).toHaveLength(13);
		expect(BORDER_STYLES.dotted.style).toBe('dotted');
		expect(BORDER_STYLES.slantDashDot.widthPx).toBe(2);
	});

	it('inherits row and column styles for empty cells', () => {
		const wb = book();
		const sheet = wb.sheets[0]!;
		const rowStyle = style(wb, {
			fill: { type: 'pattern', pattern: 'solid', fgColor: { rgb: '00FF00' } },
		});
		const colStyle = style(wb, {
			fill: { type: 'pattern', pattern: 'solid', fgColor: { rgb: '0000FF' } },
		});
		sheet.rowInfo.set(2, { styleId: rowStyle });
		sheet.columns = [{ min: 3, max: 3, styleId: colStyle }];
		expect(effectiveStyleId(sheet, 2, 0)).toBe(rowStyle);
		expect(effectiveStyleId(sheet, 0, 3)).toBe(colStyle);
		expect(cellView(wb, 0, 2, 0).fill).toEqual({ background: '#00FF00' });
		expect(cellView(wb, 0, 5, 3).fill).toEqual({ background: '#0000FF' });
		expect(cellView(wb, 0, 5, 5).fill).toBeUndefined();
	});
});

describe('cellView decorations', () => {
	it('flags comments, hyperlinks, list validations and rich text', () => {
		const wb = book();
		const sheet = wb.sheets[0]!;
		putCell(sheet, 0, 0, {
			value: 'ab',
			richText: [{ text: 'a', font: { bold: true } }, { text: 'b' }],
		});
		sheet.comments.push({ address: { row: 0, col: 0 }, author: 'A', text: 'note' });
		sheet.hyperlinks.push({
			range: { start: { row: 0, col: 0 }, end: { row: 0, col: 0 } },
			target: 'https://example.com',
			tooltip: 'Go',
		});
		sheet.dataValidations.push({
			ranges: [{ start: { row: 0, col: 0 }, end: { row: 3, col: 0 } }],
			type: 'list',
			formula1: '"a,b"',
		});
		const v = cellView(wb, 0, 0, 0);
		expect(v.hasComment).toBe(true);
		expect(v.hasHyperlink).toBe(true);
		expect(v.title).toBe('Go');
		expect(v.validationList).toBe(true);
		expect(v.rich?.map((r) => [r.text, r.font.bold])).toEqual([
			['a', true],
			['b', false],
		]);
		const other = cellView(wb, 0, 5, 5);
		expect(other.hasComment || other.hasHyperlink || other.validationList).toBe(false);
	});

	it('applies conditional-format styles, colour scales, data bars and icons', () => {
		const wb = book();
		putCell(wb.sheets[0]!, 0, 0, { value: 5 });
		const cf: ConditionalFormatEvaluator = {
			at: () => ({
				style: {
					font: { bold: true, color: { rgb: '9C0006' } },
					fill: { type: 'pattern', pattern: 'solid', bgColor: { rgb: 'FFC7CE' } },
				},
				dataBar: { fraction: 0.5, color: '#638EC6' },
				icon: { set: '3Arrows', index: 2 },
			}),
		};
		const v = cellView(wb, 0, 0, 0, cf);
		expect(v.font.bold).toBe(true);
		expect(v.font.color).toBe('#9C0006');
		expect(v.fill).toEqual({ background: '#FFC7CE' });
		expect(v.dataBar).toEqual({ fraction: 0.5, color: '#638EC6' });
		expect(v.icon).toEqual({ set: '3Arrows', index: 2 });
		const scale: ConditionalFormatEvaluator = {
			at: () => ({ colorScale: '#63BE7B', hideValue: true }),
		};
		const s = cellView(wb, 0, 0, 0, scale);
		expect(s.fill).toEqual({ background: '#63BE7B' });
		expect(s.text).toBe('');
	});

	it('marks merged text as not overflowing and rejects bad sheet indices', () => {
		const wb = book();
		const sheet = wb.sheets[0]!;
		putCell(sheet, 0, 0, { value: 'merged text' });
		sheet.merges.push({ start: { row: 0, col: 0 }, end: { row: 0, col: 2 } });
		expect(cellView(wb, 0, 0, 0).overflow).toBe(false);
		expect(() => cellView(wb, 3, 0, 0)).toThrow(RangeError);
	});
});
