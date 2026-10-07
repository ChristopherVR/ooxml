import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { parseRange } from '../address';
import { getCell } from '../cells';
import type { Border, CellValue } from '../model';
import { styleAt } from '../styles';
import { createWorkbook } from '../workbook';
import { loadXlsx } from '../read/index';
import { saveXlsx } from '../write/index';
import { createEditSession } from './session';
import type { PasteOptions } from './types';

interface NativeCell {
	row: number;
	col: number;
	value: CellValue;
	formula: string | null;
	bold: boolean;
	italic: boolean;
	numFmt: string;
	fill: number;
	borders: Record<string, { style: number; color: string }>;
}
interface NativeCase extends PasteOptions {
	cells: NativeCell[];
}
const fixture = JSON.parse(
	readFileSync(new URL('./__fixtures__/excel-paste-borders.json', import.meta.url), 'utf8'),
) as { cases: NativeCase[] };
const range = (ref: string) => parseRange(ref)!;
const edges = ['left', 'right', 'top', 'bottom', 'diagonalDown', 'diagonalUp'] as const;
const setup = () => {
	const workbook = createWorkbook();
	const session = createEditSession(workbook);
	const sheet = workbook.sheets[0]!;
	session.setCellValue(0, 0, 0, 2);
	session.setCellValue(0, 0, 1, 'text');
	session.setCellInput(0, 1, 0, '=B2+$J$1');
	session.setCellValue(0, 0, 9, 3);
	const border = (color: string, style: 'thin' | 'double'): Border => {
		const edge = { style, color: { rgb: `FF${color}` } };
		return {
			left: edge,
			right: edge,
			top: edge,
			bottom: edge,
			diagonal: edge,
			diagonalUp: true,
			diagonalDown: true,
		};
	};
	session.applyStyle(0, [range('A1:B2')], {
		font: { bold: true },
		numFmt: '0.00',
		fill: { type: 'pattern', pattern: 'solid', fgColor: { rgb: 'FF0000' } },
		border: border('FF0000', 'thin'),
	});
	const payload = session.copy(0, range('A1:B2'));
	session.setRangeValues(0, { row: 3, col: 3 }, [
		[9, 9],
		[9, 9],
	]);
	session.applyStyle(0, [range('D4:E5')], {
		font: { italic: true },
		fill: { type: 'pattern', pattern: 'solid', fgColor: { rgb: '00FF00' } },
		border: border('0000FF', 'double'),
	});
	return { workbook, session, sheet, payload };
};

describe('All Except Borders recorded in Microsoft Excel', () => {
	it('rejects cut cells before editing or consuming the move', () => {
		const { session, sheet } = setup();
		const payload = session.cut(0, range('A1:B2'));
		const label = session.undoLabel();
		expect(() => session.paste(0, range('D4:E5'), payload, 'noBorders')).toThrow(RangeError);
		expect(payload.cut).toBe(true);
		expect(session.undoLabel()).toBe(label);
		expect(getCell(sheet, 3, 3)?.value).toBe(9);
		expect(getCell(sheet, 0, 0)?.value).toBe(2);
	});
	it.each(['all', 'formulas', 'noBorders'] as const)(
		'%s uses the block offset for transposed formulas',
		(mode) => {
			const { session, sheet, payload } = setup();
			session.paste(0, { row: 3, col: 3 }, payload, { mode, transpose: true });
			expect(getCell(sheet, 3, 4)?.formula).toBe('E5+$J$1');
		},
	);
	it('tiles source merges while retaining destination borders', () => {
		const { workbook, session, sheet } = setup();
		session.merge(0, range('A1:B1'), 'merge');
		const payload = session.copy(0, range('A1:B1'));
		const border = structuredClone(styleAt(workbook, getCell(sheet, 3, 3)?.styleId).border);
		session.paste(0, range('D4:E5'), payload, 'noBorders');
		expect(sheet.merges).toContainEqual(range('D4:E4'));
		expect(sheet.merges).toContainEqual(range('D5:E5'));
		expect(styleAt(workbook, getCell(sheet, 3, 3)?.styleId).border).toEqual(border);
		session.undo();
		expect(sheet.merges).not.toContainEqual(range('D4:E4'));
	});
	it.each(fixture.cases)('transpose=$transpose skip=$skipBlanks operation=$operation', (native) => {
		const { workbook, session, sheet, payload } = setup();
		const read = () =>
			native.cells.map(({ row, col }) => {
				const cell = getCell(sheet, row + 3, col + 3);
				const style = styleAt(workbook, cell?.styleId);
				const borders: NativeCell['borders'] = {};
				for (const key of edges) {
					const diagonal = key === 'diagonalDown' || key === 'diagonalUp';
					const edge = diagonal ? style.border.diagonal : style.border[key];
					const enabled = !diagonal || style.border[key];
					borders[key] = {
						style: !enabled || !edge ? -4142 : edge.style === 'double' ? -4119 : 1,
						color: edge?.color?.rgb?.slice(-6) ?? '000000',
					};
				}
				const rgb = style.fill.type === 'pattern' ? style.fill.fgColor?.rgb?.slice(-6) : undefined;
				const fill = rgb
					? parseInt(rgb.slice(0, 2), 16) +
						(parseInt(rgb.slice(2, 4), 16) << 8) +
						(parseInt(rgb.slice(4), 16) << 16)
					: 0;
				return {
					row,
					col,
					value: cell?.value ?? null,
					formula: cell?.formula ?? null,
					bold: style.font.bold ?? false,
					italic: style.font.italic ?? false,
					numFmt: style.numFmt,
					fill,
					borders,
				};
			});
		const before = read();
		session.paste(0, range('D4:E5'), payload, native);
		expect(read()).toEqual(native.cells);
		session.undo();
		expect(read()).toEqual(before);
		session.redo();
		expect(read()).toEqual(native.cells);
	});
	it('preserves destination borders when a truly empty source clears content and other formatting', async () => {
		const { workbook, session, sheet } = setup();
		const oldBorder = structuredClone(styleAt(workbook, getCell(sheet, 3, 3)?.styleId).border);
		session.paste(0, range('D4:E5'), session.copy(0, range('L1')), 'noBorders');
		const cell = getCell(sheet, 3, 3);
		expect(cell?.value).toBeNull();
		expect(styleAt(workbook, cell?.styleId).border).toEqual(oldBorder);
		expect(styleAt(workbook, cell?.styleId).font.italic).toBeUndefined();
		const loaded = await loadXlsx(await saveXlsx(workbook));
		expect(styleAt(loaded, getCell(loaded.sheets[0]!, 3, 3)?.styleId).border).toEqual(oldBorder);
	});
});
