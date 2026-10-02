import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { parseRange } from '../address.js';
import { getCell } from '../cells.js';
import { BUILTIN_CELL_STYLES, builtinCellStyle } from '../cell-styles.js';
import { loadXlsx } from '../read/index.js';
import { patchStyle, styleAt } from '../styles.js';
import { createWorkbook } from '../workbook.js';
import { saveXlsx } from '../write/index.js';
import { createEditSession } from './session.js';

const R = (ref: string) => {
	const r = parseRange(ref);
	if (!r) throw new Error(ref);
	return r;
};

describe('cell styles', () => {
	it('lists the built-in styles with unique names and Excel builtin ids', () => {
		const names = BUILTIN_CELL_STYLES.map((s) => s.name);
		expect(new Set(names).size).toBe(names.length);
		expect(builtinCellStyle('good')?.builtinId).toBe(26);
		expect(builtinCellStyle('Heading 1')?.builtinId).toBe(16);
		expect(builtinCellStyle('40% - Accent3')?.builtinId).toBe(39);
		expect(builtinCellStyle('Explanatory Text')?.builtinId).toBe(53);
		expect(names).toContain('Currency [0]');
	});

	it('patches cellStyleName through StylePatch', () => {
		const base = styleAt(createWorkbook(), 0);
		expect(patchStyle(base, { cellStyleName: 'Good' }).cellStyleName).toBe('Good');
		const named = { ...base, cellStyleName: 'Good' };
		expect(patchStyle(named, { cellStyleName: undefined }).cellStyleName).toBeUndefined();
		expect(patchStyle(named, { numFmt: '0' }).cellStyleName).toBe('Good');
	});

	it('applies a named style, keeps other aspects and writes it to the gallery', async () => {
		const wb = createWorkbook();
		const s = createEditSession(wb, { recalc: false });
		s.setCellValue(0, 0, 0, 5);
		s.applyStyle(0, [R('A1')], { numFmt: '0.00' });
		s.applyCellStyle(0, [R('A1:B2')], 'Good');
		const sheet = wb.sheets[0]!;
		const style = styleAt(wb, getCell(sheet, 0, 0)?.styleId);
		expect(style.cellStyleName).toBe('Good');
		expect(style.font.color).toEqual({ rgb: 'FF006100' });
		expect(style.numFmt).toBe('0.00');
		expect(styleAt(wb, getCell(sheet, 1, 1)?.styleId).cellStyleName).toBe('Good');
		expect(s.undoLabel()).toBe('Cell style Good');
		s.applyCellStyle(0, [R('C1')], 'Heading 1');
		expect(() => s.applyCellStyle(0, [R('C1')], 'No Such Style')).toThrow();
		const bytes = await saveXlsx(wb);
		const xml = (await (await JSZip.loadAsync(bytes)).file('xl/styles.xml')?.async('text')) ?? '';
		expect(xml).toMatch(/<cellStyle name="Good" xfId="\d+" builtinId="26"\/>/);
		expect(xml).toMatch(/<cellStyle name="Heading 1" xfId="\d+" builtinId="16"\/>/);
		const back = await loadXlsx(bytes);
		expect(back.namedStyles.map((n) => n.name)).toEqual(
			expect.arrayContaining(['Good', 'Heading 1']),
		);
		const cell = getCell(back.sheets[0]!, 0, 0);
		expect(styleAt(back, cell?.styleId).cellStyleName).toBe('Good');
		expect(styleAt(back, cell?.styleId).fill).toMatchObject({ pattern: 'solid' });
		s.undo();
		s.undo();
		expect(styleAt(wb, getCell(sheet, 0, 0)?.styleId).cellStyleName).toBeUndefined();
	});

	it('resets every aspect with Normal', () => {
		const wb = createWorkbook();
		const s = createEditSession(wb, { recalc: false });
		s.applyCellStyle(0, [R('A1')], 'Bad');
		s.applyStyle(0, [R('A1')], { numFmt: '0%' });
		s.applyCellStyle(0, [R('A1')], 'Normal');
		const style = styleAt(wb, getCell(wb.sheets[0]!, 0, 0)?.styleId);
		expect(style.numFmt).toBe('General');
		expect(style.fill).toEqual(styleAt(wb, 0).fill);
	});
});
