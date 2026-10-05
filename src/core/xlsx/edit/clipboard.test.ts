import { describe, expect, it } from 'vitest';
import { parseAddress, parseRange } from '../address.js';
import { getCell } from '../cells.js';
import type { CellValue, Workbook, Worksheet } from '../model.js';
import { defaultCellStyle } from '../workbook.js';
import { styleAt } from '../styles.js';
import { createWorkbook } from '../workbook.js';
import { parseHtmlTable } from './clipboard-html-parse.js';
import { parseTsv, toTsv } from './clipboard-text.js';
import { createEditSession } from './session.js';

const A = (ref: string) => {
	const a = parseAddress(ref);
	if (!a) throw new Error(ref);
	return a;
};
const R = (ref: string) => {
	const r = parseRange(ref);
	if (!r) throw new Error(ref);
	return r;
};
const ws = (wb: Workbook, i = 0): Worksheet => {
	const sheet = wb.sheets[i];
	if (!sheet) throw new Error('sheet');
	return sheet;
};
const cell = (wb: Workbook, ref: string, i = 0) => getCell(ws(wb, i), A(ref).row, A(ref).col);
const values = (wb: Workbook, ref: string, i = 0): CellValue[] => {
	const r = R(ref);
	const out: CellValue[] = [];
	for (let row = r.start.row; row <= r.end.row; row++)
		for (let col = r.start.col; col <= r.end.col; col++)
			out.push(getCell(ws(wb, i), row, col)?.value ?? null);
	return out;
};
const setup = () => {
	const wb = createWorkbook({ sheets: ['Sheet1', 'Sheet2'] });
	const s = createEditSession(wb, { recalc: false });
	return { wb, s };
};

describe('TSV', () => {
	it('writes Excel-style text with quoting', () => {
		expect(
			toTsv([
				['a', 'b'],
				['c', 'd'],
			]),
		).toBe('a\tb\r\nc\td\r\n');
		expect(toTsv([['x\ty', 'say "hi"', 'two\nlines']])).toBe(
			'"x\ty"\t"say ""hi"""\t"two\nlines"\r\n',
		);
	});
	it('parses rows, quoted fields and a trailing line break', () => {
		expect(parseTsv('a\tb\r\nc\td\r\n')).toEqual([
			['a', 'b'],
			['c', 'd'],
		]);
		expect(parseTsv('"x\ty"\t"say ""hi"""\r\n"two\nlines"\tz')).toEqual([
			['x\ty', 'say "hi"'],
			['two\nlines', 'z'],
		]);
	});
	it('pads ragged rows and keeps empty fields', () => {
		expect(parseTsv('a\t\tc\nd')).toEqual([
			['a', '', 'c'],
			['d', '', ''],
		]);
		expect(parseTsv('')).toEqual([]);
	});
	it('round-trips through toTsv', () => {
		const rows = [
			['1', 'a "q"'],
			['tab\there', ''],
		];
		expect(parseTsv(toTsv(rows))).toEqual(rows);
	});
});

describe('copy', () => {
	it('produces TSV, HTML and cells', () => {
		const { s } = setup();
		s.setRangeValues(0, A('A1'), [
			['Name', 'Qty'],
			['<pear>', 2],
		]);
		s.applyStyle(0, [R('A1:B1')], { font: { bold: true } });
		const payload = s.copy(0, R('A1:B2'));
		expect(payload.tsv).toBe('Name\tQty\r\n<pear>\t2\r\n');
		expect(payload.html).toContain('<table');
		expect(payload.html).toContain('&lt;pear&gt;');
		expect(payload.html).toContain('font-weight:700');
		expect(payload.html).toContain('x:num="2"');
		expect(payload.cells.rows).toBe(2);
		expect(payload.cells.data[0]?.[0]?.style?.font.bold).toBe(true);
	});
	it('uses display text for formatted values', () => {
		const { s } = setup();
		s.setCellInput(0, 0, 0, '15%');
		expect(s.copy(0, R('A1')).tsv).toBe('15%\r\n');
	});
	it('clips whole columns to the used area', () => {
		const { s } = setup();
		s.setRangeValues(0, A('A1'), [[1], [2]]);
		expect(s.copy(0, R('A:A')).cells.rows).toBe(2);
	});
	it('writes merges as spans', () => {
		const { s } = setup();
		s.setCellValue(0, 0, 0, 'm');
		s.merge(0, R('A1:B1'), 'merge');
		const payload = s.copy(0, R('A1:B2'));
		expect(payload.html).toContain('colspan="2"');
		expect(payload.cells.merges).toEqual([R('A1:B1')]);
	});
});

describe('paste', () => {
	it('pastes values, formulas and formats, translating formulas', () => {
		const { wb, s } = setup();
		s.setCellValue(0, 0, 0, 1);
		s.setCellInput(0, 0, 1, '=A1+1');
		s.applyStyle(0, [R('A1')], { font: { italic: true } });
		const payload = s.copy(0, R('A1:B1'));
		const area = s.paste(0, A('C5'), payload);
		expect(area).toEqual(R('C5:D5'));
		expect(cell(wb, 'C5')?.value).toBe(1);
		expect(cell(wb, 'D5')?.formula).toBe('C5+1');
		expect(styleAt(wb, cell(wb, 'C5')?.styleId).font.italic).toBe(true);
	});
	it('pastes values only, keeping destination formats', () => {
		const { wb, s } = setup();
		s.setCellValue(0, 0, 0, 5);
		s.setCellInput(0, 0, 1, '=A1*2');
		const b1 = cell(wb, 'B1');
		if (b1) b1.value = 10;
		s.applyStyle(0, [R('D1')], { font: { bold: true } });
		s.paste(0, A('C1'), s.copy(0, R('A1:B1')), 'values');
		expect(cell(wb, 'D1')?.value).toBe(10);
		expect(cell(wb, 'D1')?.formula).toBeUndefined();
		expect(styleAt(wb, cell(wb, 'D1')?.styleId).font.bold).toBe(true);
	});
	it('pastes formats only', () => {
		const { wb, s } = setup();
		s.setCellValue(0, 0, 0, 'src');
		s.applyStyle(0, [R('A1')], {
			fill: { type: 'pattern', pattern: 'solid', fgColor: { rgb: 'FF0000' } },
		});
		s.setCellValue(0, 2, 2, 'keep');
		s.paste(0, A('C3'), s.copy(0, R('A1')), 'formats');
		expect(cell(wb, 'C3')?.value).toBe('keep');
		expect(styleAt(wb, cell(wb, 'C3')?.styleId).fill).toMatchObject({ pattern: 'solid' });
	});
	it('pastes formulas without formats', () => {
		const { wb, s } = setup();
		s.setCellInput(0, 0, 0, '=B1');
		s.applyStyle(0, [R('A1')], { font: { bold: true } });
		s.paste(0, A('A3'), s.copy(0, R('A1')), 'formulas');
		expect(cell(wb, 'A3')?.formula).toBe('B3');
		expect(cell(wb, 'A3')?.styleId).toBeUndefined();
	});
	it('transposes rows and columns', () => {
		const { wb, s } = setup();
		s.setRangeValues(0, A('A1'), [
			[1, 2, 3],
			[4, 5, 6],
		]);
		const area = s.paste(0, A('E1'), s.copy(0, R('A1:C2')), 'transpose');
		expect(area).toEqual(R('E1:F3'));
		expect(values(wb, 'E1:F3')).toEqual([1, 4, 2, 5, 3, 6]);
	});
	it('transposes merges', () => {
		const { wb, s } = setup();
		s.setCellValue(0, 0, 0, 'm');
		s.merge(0, R('A1:B1'), 'merge');
		s.paste(0, A('D1'), s.copy(0, R('A1:B1')), 'transpose');
		expect(ws(wb).merges).toContainEqual(R('D1:D2'));
	});
	it('pastes blanks over existing cells', () => {
		const { wb, s } = setup();
		s.setRangeValues(0, A('A1'), [[1, null]]);
		s.setCellValue(0, 4, 1, 'old');
		s.paste(0, A('A5'), s.copy(0, R('A1:B1')));
		expect(cell(wb, 'B5')).toBeUndefined();
	});
	it('moves cells on a cut paste; references to the moved cells follow them', () => {
		const { wb, s } = setup();
		s.setCellValue(0, 0, 0, 1);
		s.setCellInput(0, 0, 1, '=A1');
		const payload = s.cut(0, R('A1:B1'));
		s.paste(0, A('A3'), payload);
		expect(cell(wb, 'A1')).toBeUndefined();
		expect(cell(wb, 'B3')?.formula).toBe('A3');
		expect(payload.cut).toBe(false);
		s.undo();
		expect(cell(wb, 'A1')?.value).toBe(1);
		expect(cell(wb, 'A3')).toBeUndefined();
	});
	it('cuts across sheets', () => {
		const { wb, s } = setup();
		s.setCellValue(0, 0, 0, 'x');
		s.paste(1, A('B2'), s.cut(0, R('A1')));
		expect(cell(wb, 'A1')).toBeUndefined();
		expect(cell(wb, 'B2', 1)?.value).toBe('x');
		s.undo();
		expect(cell(wb, 'A1')?.value).toBe('x');
		expect(cell(wb, 'B2', 1)).toBeUndefined();
	});
	it('pastes TSV text through input parsing', () => {
		const { wb, s } = setup();
		const area = s.paste(0, A('B2'), '1,234\t15%\r\nabc\t=1+1\r\n');
		expect(area).toEqual(R('B2:C3'));
		expect(values(wb, 'B2:B3')).toEqual([1234, 'abc']);
		expect(cell(wb, 'C2')?.value).toBeCloseTo(0.15);
		expect(styleAt(wb, cell(wb, 'C2')?.styleId).numFmt).toBe('0%');
		expect(cell(wb, 'C3')?.formula).toBe('1+1');
	});
	it('pastes text into the destination format', () => {
		const { wb, s } = setup();
		s.applyStyle(0, [R('A1')], { font: { bold: true } });
		s.paste(0, A('A1'), 'hello');
		expect(styleAt(wb, cell(wb, 'A1')?.styleId).font.bold).toBe(true);
	});
	it('undoes a paste', () => {
		const { wb, s } = setup();
		s.paste(0, A('A1'), 'a\tb');
		s.undo();
		expect(ws(wb).rows.size).toBe(0);
	});
	it('rejects a paste beyond the sheet edge', () => {
		const { s } = setup();
		expect(() => s.paste(0, { row: 1_048_575, col: 0 }, 'a\nb')).toThrow(RangeError);
	});
});

const EXCEL_HTML = `<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">
<head><meta http-equiv=Content-Type content="text/html; charset=utf-8">
<style>
<!--table
	{mso-displayed-decimal-separator:"\\.";}
.xl65
	{font-weight:700;
	color:#C00000;}
.xl66
	{mso-number-format:"0\\.00";
	background:yellow;}
.xl67
	{mso-number-format:"\\@";}
-->
</style>
</head>
<body link="#0563C1" vlink="#954F72">
<table border=0 cellpadding=0 cellspacing=0 width=128 style='border-collapse:collapse;width:96pt'>
<!--StartFragment-->
 <tr height=20 style='height:15.0pt'>
  <td height=20 class=xl65 width=64 style='height:15.0pt;width:48pt'>Header</td>
  <td class=xl65 width=64 style='width:48pt'>Value</td>
 </tr>
 <tr height=20 style='height:15.0pt'>
  <td height=20 style='height:15.0pt'>Tom &amp; Jerry</td>
  <td class=xl66 align=right x:num="3.5">3.50</td>
 </tr>
 <tr height=20 style='height:15.0pt'>
  <td height=20 class=xl67 style='height:15.0pt'>00123</td>
  <td align=right>1,234</td>
 </tr>
 <tr>
  <td colspan=2 style='text-align:center'>Merged<br>two lines</td>
 </tr>
<!--EndFragment-->
</table>
</body>
</html>`;

describe('HTML table paste', () => {
	it('parses an Excel clipboard table', () => {
		const cells = parseHtmlTable(EXCEL_HTML, defaultCellStyle());
		expect(cells?.rows).toBe(4);
		expect(cells?.cols).toBe(2);
		const at = (r: number, c: number) => cells?.data[r]?.[c];
		expect(at(0, 0)?.value).toBe('Header');
		expect(at(0, 0)?.style?.font.bold).toBe(true);
		expect(at(0, 0)?.style?.font.color).toEqual({ rgb: 'C00000' });
		expect(at(1, 0)?.value).toBe('Tom & Jerry');
		expect(at(1, 1)?.value).toBe(3.5);
		expect(at(1, 1)?.style?.numFmt).toBe('0.00');
		expect(at(1, 1)?.style?.fill).toEqual({
			type: 'pattern',
			pattern: 'solid',
			fgColor: { rgb: 'FFFF00' },
		});
		expect(at(2, 0)?.value).toBe('00123');
		expect(at(2, 1)?.value).toBe(1234);
		expect(at(3, 0)?.value).toBe('Merged\ntwo lines');
		expect(at(3, 0)?.style?.alignment?.horizontal).toBe('center');
		expect(cells?.merges).toEqual([R('A4:B4')]);
	});
	it('pastes HTML into the sheet with formats and merges', () => {
		const { wb, s } = setup();
		const area = s.paste(0, A('B2'), EXCEL_HTML);
		expect(area).toEqual(R('B2:C5'));
		expect(cell(wb, 'C3')?.value).toBe(3.5);
		expect(styleAt(wb, cell(wb, 'B2')?.styleId).font.bold).toBe(true);
		expect(ws(wb).merges).toEqual([R('B5:C5')]);
	});
	it('parses simple browser tables with th, rowspan and entities', () => {
		const cells = parseHtmlTable(
			'<table><tr><th>A</th><th>B</th></tr><tr><td rowspan="2">x&nbsp;y</td><td><b>1</b></td></tr><tr><td>&#50;</td></tr></table>',
			defaultCellStyle(),
		);
		expect(cells?.data[0]?.[0]?.style?.font.bold).toBe(true);
		expect(cells?.data[1]?.[0]?.value).toBe('x y');
		expect(cells?.data[1]?.[1]?.value).toBe(1);
		expect(cells?.data[2]?.[0]).toBeNull();
		expect(cells?.data[2]?.[1]?.value).toBe(2);
		expect(cells?.merges).toEqual([R('A2:A3')]);
	});
	it('returns undefined without a table', () => {
		expect(parseHtmlTable('<p>hi</p>', defaultCellStyle())).toBeUndefined();
	});
	it('round-trips copied HTML', () => {
		const { wb, s } = setup();
		s.setRangeValues(0, A('A1'), [['a', 1.5]]);
		s.applyStyle(0, [R('B1')], { numFmt: '0.00' });
		const html = s.copy(0, R('A1:B1')).html;
		s.paste(1, A('A1'), html);
		expect(values(wb, 'A1:B1', 1)).toEqual(['a', 1.5]);
		expect(styleAt(wb, cell(wb, 'B1', 1)?.styleId).numFmt).toBe('0.00');
	});
});
