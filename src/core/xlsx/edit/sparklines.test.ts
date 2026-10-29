import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { parseRange } from '../address';
import { getCell } from '../cells';
import { sheetSparklineGroups } from '../read/sparklines';
import { loadXlsx } from '../read/index';
import { saveXlsx } from '../write/index';
import { createWorkbook } from '../workbook';
import { createEditSession } from './session';

const range = (ref: string) => parseRange(ref)!;
const setup = () => {
	const wb = createWorkbook({ sheets: ['Sales', "O'Brien"] });
	const edit = createEditSession(wb);
	return { wb, edit, sheet: wb.sheets[0]!, groups: () => sheetSparklineGroups(wb.sheets[0]!) };
};

describe('sparkline authoring', () => {
	it('authors, clears and restores date vectors through saved XML and history', async () => {
		const { wb, edit, groups } = setup();
		edit.createSparklineGroup(0, 'A1:C2', range('D1:D2'));
		edit.setSparklineDateRange(0, [range('D1')], "'O''Brien'!A1:C1");
		expect(groups()[0]).toMatchObject({ dateAxis: true, dateFormula: "'O''Brien'!A1:C1" });
		edit.insertColumns(1, 0, 1);
		edit.renameSheet(1, 'Dates');
		expect(groups()[0]?.dateFormula).toBe('Dates!B1:D1');
		const reloaded = await loadXlsx(await saveXlsx(wb));
		expect(sheetSparklineGroups(reloaded.sheets[0]!)[0]?.dateFormula).toBe('Dates!B1:D1');
		edit.setSparklineDateRange(0, [range('D2')]);
		expect(groups()[0]).toMatchObject({ dateAxis: false });
		expect(groups()[0]?.dateFormula).toBeUndefined();
		edit.undo();
		expect(groups()[0]?.dateAxis).toBe(true);
		edit.redo();
		expect(groups()[0]?.dateAxis).toBe(false);
	});
	it('rejects invalid date vectors atomically across selected groups', () => {
		const { edit, groups } = setup();
		edit.createSparklineGroup(0, 'A1:C1', range('F1'));
		edit.createSparklineGroup(0, 'A2:D2', range('F2'));
		for (const ref of ['A3:C3', 'A1:C2', '[Book.xlsx]Sheet1!A1:C1', 'A1:A10001'])
			expect(() => edit.setSparklineDateRange(0, [range('F1:F2')], ref)).toThrow();
		expect(groups().every((group) => !group.dateAxis)).toBe(true);
	});
	it('normalizes reversed location and selection ranges before overlap checks and edits', () => {
		const { edit, groups } = setup();
		const reversed = { start: { row: 1, col: 3 }, end: { row: 0, col: 3 } };
		edit.createSparklineGroup(0, 'A1:C2', reversed);
		expect(groups()[0]?.sparklines.map((line) => line.host)).toEqual([range('D1'), range('D2')]);
		expect(() => edit.createSparklineGroup(0, 'A1:C2', reversed)).toThrow('already contains');
		edit.setSparklineGroups(0, [reversed], { type: 'column' });
		expect(groups()[0]?.type).toBe('column');
		edit.deleteSparklines(0, [reversed]);
		expect(groups()).toEqual([]);
		edit.undo();
		expect(groups()[0]?.sparklines).toHaveLength(2);
	});
	it('validates color selectors before writing XML and leaves the workbook unchanged on failure', () => {
		const { edit, sheet, groups } = setup();
		for (const color of [{}, { indexed: -1 }, { indexed: 1.5 }, { indexed: 66 }])
			expect(() =>
				edit.createSparklineGroup(0, 'A1:C1', range('D1'), { colorSeries: color }),
			).toThrow('settings');
		expect(groups()).toEqual([]);
		expect(edit.canUndo()).toBe(false);
		edit.createSparklineGroup(0, 'A1:C1', range('D1'), { colorSeries: { indexed: 3 } });
		const before = sheet.preserved.get('extLst')![0];
		expect(() => edit.setSparklineGroups(0, [range('D1')], { colorSeries: {} })).toThrow(
			'settings',
		);
		expect(sheet.preserved.get('extLst')![0]).toBe(before);
	});
	it('ignores runtime unknown and undefined patch keys and rejects invalid scalar types', () => {
		const { edit, sheet, groups } = setup();
		const malformed = {
			type: 'line',
			lineWeight: undefined,
			unknown: 'bad',
			dateAxis: true,
			sparklines: [],
			high: true,
		} as unknown as Parameters<typeof edit.createSparklineGroup>[3];
		edit.createSparklineGroup(0, 'A1:C1', range('D1'), malformed);
		expect(groups()[0]).toMatchObject({
			type: 'line',
			lineWeight: 0.75,
			dateAxis: false,
			high: true,
		});
		const xml = sheet.preserved.get('extLst')![0];
		expect(xml).not.toContain('unknown=');
		expect(xml).not.toContain('undefined');
		expect(xml).not.toContain('dateAxis=');
		expect(() => edit.setSparklineGroups(0, [range('D1')], { markers: 'false' } as never)).toThrow(
			'settings',
		);
		expect(sheet.preserved.get('extLst')![0]).toBe(xml);
		expect(() =>
			edit.deleteSparklines(0, [{ start: { row: -1, col: 0 }, end: { row: -1, col: 1 } }]),
		).toThrow();
	});
	it('maps data rows to vertical hosts, preserving cell content and undo', () => {
		const { edit, sheet, groups } = setup();
		edit.setCellInput(0, 0, 3, 'note');
		edit.createSparklineGroup(0, 'A1:C2', range('D1:D2'), { type: 'column', high: true });
		expect(groups()[0]).toMatchObject({
			type: 'column',
			high: true,
			displayEmptyCellsAs: 'gap',
			sparklines: [
				{ formula: "'Sales'!A1:C1", host: range('D1') },
				{ formula: "'Sales'!A2:C2", host: range('D2') },
			],
		});
		expect(getCell(sheet, 0, 3)?.value).toBe('note');
		expect(sheet.preserved.get('extLst')![0]).toContain('{05C60535-1F16-4fd2-B633-F4F36F0B64E0}');
		edit.undo();
		expect(groups()).toEqual([]);
		edit.redo();
		expect(groups()[0]?.sparklines).toHaveLength(2);
	});
	it('maps data columns and handles quoted source sheets', () => {
		const { edit, groups } = setup();
		edit.createSparklineGroup(0, "='O''Brien'!$A$1:$B$3", range('D1:E1'), { type: 'stacked' });
		expect(groups()[0]?.sparklines.map((line) => line.formula)).toEqual([
			"'O''Brien'!A1:A3",
			"'O''Brien'!B1:B3",
		]);
	});
	it('keeps unknown extensions and group attributes when settings change', () => {
		const { edit, sheet, groups } = setup();
		sheet.preserved.set('extLst', [
			'<extLst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><ext uri="custom"><custom xmlns="urn:test" value="keep"/></ext></extLst>',
		]);
		edit.createSparklineGroup(0, 'A1:C1', range('D1'));
		sheet.preserved.set(
			'extLst',
			sheet.preserved
				.get('extLst')!
				.map((xml) => xml.replace('<x14:sparklineGroup ', '<x14:sparklineGroup custom="keep" ')),
		);
		edit.setSparklineGroups(0, [range('D1')], {
			markers: true,
			colorSeries: { rgb: '123456' },
			minAxisType: 'custom',
			manualMin: -10,
		});
		expect(groups()[0]).toMatchObject({
			markers: true,
			colorSeries: { rgb: 'FF123456' },
			minAxisType: 'custom',
			manualMin: -10,
		});
		expect(sheet.preserved.get('extLst')![0]).toContain('custom="keep"');
		expect(sheet.preserved.get('extLst')![0]).toContain('value="keep"');
		edit.setSparklineGroups(0, [range('D1')], { minAxisType: 'group' });
		expect(groups()[0]?.manualMin).toBeUndefined();
	});
	it('clears one selected host or its whole group and undoes it', () => {
		const { edit, groups } = setup();
		edit.createSparklineGroup(0, 'A1:C2', range('D1:D2'));
		edit.deleteSparklines(0, [range('D1')]);
		expect(groups()[0]?.sparklines.map((line) => line.host)).toEqual([range('D2')]);
		edit.undo();
		expect(groups()[0]?.sparklines).toHaveLength(2);
		edit.deleteSparklines(0, [range('D1')], true);
		expect(groups()).toEqual([]);
		edit.undo();
		expect(groups()[0]?.sparklines).toHaveLength(2);
	});
	it('validates inputs atomically without adding an undo step', () => {
		const { edit, sheet, groups } = setup();
		for (const source of ['bad', '[1]Sales!A1:C1', 'Sales:Other!A1:C1'])
			expect(() => edit.createSparklineGroup(0, source, range('D1'))).toThrow();
		expect(() => edit.createSparklineGroup(0, 'A1:C2', range('D1:D3'))).toThrow('do not match');
		expect(() => edit.createSparklineGroup(0, 'A1:C1', range('D1'), { lineWeight: -1 })).toThrow();
		expect(() =>
			edit.createSparklineGroup(0, 'A1:C1', range('D1'), { minAxisType: 'custom' }),
		).toThrow();
		expect(groups()).toEqual([]);
		expect(sheet.preserved.size).toBe(0);
		expect(edit.canUndo()).toBe(false);
	});
	it('saves, reloads and follows row inserts and sheet renames', async () => {
		const { wb, edit, groups } = setup();
		edit.createSparklineGroup(0, "'O''Brien'!A1:C2", range('D1:D2'), {
			type: 'column',
			negative: true,
		});
		edit.insertRows(0, 0, 1);
		edit.renameSheet(1, 'Renamed');
		expect(groups()[0]?.sparklines[0]).toEqual({ formula: 'Renamed!A1:C1', host: range('D2') });
		const loaded = await loadXlsx(await saveXlsx(wb));
		expect(sheetSparklineGroups(loaded.sheets[0]!)).toEqual(groups());
	});
	it('edits native imported groups and keeps untouched groups', async () => {
		const wb = await loadXlsx(
			new Uint8Array(
				readFileSync(new URL('../__fixtures__/excel-sparklines.xlsx', import.meta.url)),
			),
		);
		const edit = createEditSession(wb);
		const sheet = wb.sheets[0]!;
		const before = sheetSparklineGroups(sheet);
		const host = before[0]!.sparklines[0]!.host;
		edit.setSparklineGroups(0, [host], { type: 'stacked', high: true });
		const after = sheetSparklineGroups(sheet);
		expect(after[0]?.type).toBe('stacked');
		expect(after.slice(1)).toEqual(before.slice(1));
		expect(sheetSparklineGroups((await loadXlsx(await saveXlsx(wb))).sheets[0]!)).toEqual(after);
	});

	it('adds a distinct revision uid when extending a native workbook', async () => {
		const wb = await loadXlsx(
			new Uint8Array(
				readFileSync(new URL('../__fixtures__/excel-sparklines.xlsx', import.meta.url)),
			),
		);
		const edit = createEditSession(wb);
		edit.createSparklineGroup(0, 'A1:E1', range('G1'));
		const loaded = await loadXlsx(await saveXlsx(wb));
		expect(sheetSparklineGroups(loaded.sheets[0]!)).toHaveLength(5);
		const xml = loaded.sheets[0]!.preserved.get('extLst')!.join('');
		const uids = [...xml.matchAll(/xr2:uid="([^"]+)"/g)].map((match) => match[1]);
		expect(uids).toHaveLength(5);
		expect(new Set(uids).size).toBe(5);
	});

	it('groups sparklines of several groups under the first group and ungroups them again', async () => {
		const { wb, edit, groups } = setup();
		edit.createSparklineGroup(0, 'A1:C1', range('F1'), { type: 'column', markers: true });
		edit.createSparklineGroup(0, 'A2:C2', range('F2'));
		edit.createSparklineGroup(0, 'A3:C3', range('F3'));
		expect(groups()).toHaveLength(3);
		// One group touched: nothing to merge.
		expect(edit.groupSparklines(0, [range('F1')])).toBe(1);
		expect(groups()).toHaveLength(3);

		expect(edit.groupSparklines(0, [range('F1:F3')])).toBe(3);
		expect(groups()).toHaveLength(1);
		expect(groups()[0]).toMatchObject({ type: 'column', markers: true });
		expect(groups()[0]?.sparklines.map((line) => line.host)).toEqual([
			range('F1'),
			range('F2'),
			range('F3'),
		]);
		const reloaded = await loadXlsx(await saveXlsx(wb));
		expect(sheetSparklineGroups(reloaded.sheets[0]!)).toHaveLength(1);

		// A group setting now applies to all three.
		edit.setSparklineGroups(0, [range('F2')], { type: 'line' });
		expect(groups()[0]?.type).toBe('line');

		expect(edit.ungroupSparklines(0, [range('F1')])).toBe(2);
		expect(groups()).toHaveLength(3);
		expect(groups().map((g) => g.sparklines.map((l) => l.formula))).toEqual([
			["'Sales'!A1:C1"],
			["'Sales'!A2:C2"],
			["'Sales'!A3:C3"],
		]);
		expect(groups().every((g) => g.type === 'line' && g.markers)).toBe(true);
		edit.undo();
		expect(groups()).toHaveLength(1);
		edit.undo();
		edit.undo();
		expect(groups()).toHaveLength(3);
		expect(edit.ungroupSparklines(0, [range('F1')])).toBe(0);
	});
});
