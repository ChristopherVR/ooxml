import { describe, expect, it } from 'vitest';
import { createWorkbook } from '../workbook';
import { createEditSession } from './session';
import { loadXlsx } from '../read/index';
import { saveXlsx } from '../write/index';
import { viewNamePrefix } from './custom-views';

function build() {
	const wb = createWorkbook({ sheets: ['Data', 'Notes'] });
	const s = createEditSession(wb);
	for (let r = 0; r < 6; r++)
		for (let c = 0; c < 4; c++) s.setCellInput(0, r, c, `R${r + 1}C${c + 1}`);
	return { wb, s, sheet: wb.sheets[0]! };
}

const hiddenRows = (sheet: ReturnType<typeof build>['sheet']) =>
	[...sheet.rowInfo].filter(([, i]) => i.hidden).map(([r]) => r);

describe('custom views', () => {
	it('saves and shows hidden rows and columns, view settings and print settings', () => {
		const { wb, s, sheet } = build();
		s.setHidden(0, 'row', [2], true);
		s.setHidden(0, 'col', [2], true);
		s.setSheetView(0, { zoom: 85, showGridLines: false });
		s.setPageSetup(0, { orientation: 'landscape', header: '&CCompact' });
		const compact = s.addCustomView({ name: 'Compact' });
		const names = wb.definedNames.filter((n) => n.name.startsWith(viewNamePrefix(compact)));
		expect(names.map((n) => [n.name.slice(viewNamePrefix(compact).length), n.formula])).toEqual([
			['Rows', 'Data!$3:$3'],
			['Cols', 'Data!$C:$C'],
		]);
		expect(names.every((n) => n.localSheet === 0 && n.hidden)).toBe(true);

		// Back to a plain look, saved without print and hidden settings.
		s.setHidden(0, 'row', [2], false);
		s.setHidden(0, 'col', [2], false);
		s.setSheetView(0, { zoom: 100, showGridLines: true });
		s.setPageSetup(0, { orientation: 'portrait', header: '' });
		const full = s.addCustomView({
			name: 'Full',
			includePrintSettings: false,
			includeHiddenRowCol: false,
		});
		expect(wb.customViews?.map((v) => v.name)).toEqual(['Compact', 'Full']);
		expect(wb.sheets[1]!.customViews).toHaveLength(2);

		s.showCustomView(compact);
		expect(hiddenRows(sheet)).toEqual([2]);
		expect(sheet.columns.some((c) => c.hidden && c.min <= 2 && c.max >= 2)).toBe(true);
		expect(sheet.view).toMatchObject({ zoom: 85, showGridLines: false });
		expect(sheet.pageSetup).toMatchObject({ orientation: 'landscape', header: '&CCompact' });

		// "Full" records no hidden rows or print settings: only the view settings come back.
		s.showCustomView(full);
		expect(sheet.view).toMatchObject({ zoom: 100, showGridLines: true });
		expect(hiddenRows(sheet)).toEqual([2]);
		expect(sheet.pageSetup).toMatchObject({ orientation: 'landscape' });

		s.undo();
		// Undo restores sheet copies, so read the sheet again.
		expect(wb.sheets[0]!.view.zoom).toBe(85);
		s.undo();
		expect(wb.customViews?.map((v) => v.name)).toEqual(['Compact', 'Full']);
	});

	it('records and reapplies the filter', () => {
		const { wb, s, sheet } = build();
		s.setAutoFilter(0, { start: { row: 0, col: 0 }, end: { row: 5, col: 3 } });
		s.filterColumn(0, 1, ['R2C2']);
		const filtered = s.addCustomView({ name: 'Filtered' });
		expect(hiddenRows(sheet)).toEqual([2, 3, 4, 5]);
		expect(
			wb.definedNames.find((n) => n.name === `${viewNamePrefix(filtered)}FilterData`)?.formula,
		).toBe('Data!$A$1:$D$6');

		s.setAutoFilter(0, undefined);
		expect(hiddenRows(sheet)).toEqual([]);
		const plain = s.addCustomView({ name: 'Plain' });
		s.showCustomView(filtered);
		expect(sheet.autoFilter?.columns?.[0]).toMatchObject({ offset: 1, values: ['R2C2'] });
		expect(hiddenRows(sheet)).toEqual([2, 3, 4, 5]);
		s.showCustomView(plain);
		expect(sheet.autoFilter).toBeUndefined();
		expect(hiddenRows(sheet)).toEqual([]);
	});

	it('replaces a view of the same name, deletes views with their names and refuses tables', () => {
		const { wb, s } = build();
		const first = s.addCustomView({ name: 'View1' });
		const second = s.addCustomView({ name: 'view1' });
		expect(second).not.toBe(first);
		expect(wb.customViews).toHaveLength(1);
		s.setHidden(0, 'row', [1], true);
		const third = s.addCustomView({ name: 'Other' });
		expect(wb.definedNames.some((n) => n.name.startsWith(viewNamePrefix(third)))).toBe(true);
		s.deleteCustomView(third);
		expect(wb.definedNames.some((n) => n.name.startsWith(viewNamePrefix(third)))).toBe(false);
		expect(wb.sheets[0]!.customViews?.map((v) => v.guid)).toEqual([second]);
		s.deleteCustomView(second);
		expect(wb.customViews).toEqual([]);
		expect(wb.sheets[0]!.customViews).toBeUndefined();
		expect(() => s.addCustomView({ name: ' ' })).toThrow('Enter a name');
		expect(() => s.showCustomView('{missing}')).toThrow('does not exist');
		s.createTable(0, { start: { row: 0, col: 0 }, end: { row: 5, col: 3 } }, true);
		expect(() => s.addCustomView({ name: 'T' })).toThrow('contains tables');
	});

	it('survives a save and reload and still shows', async () => {
		const { wb, s, sheet } = build();
		s.setHidden(0, 'row', [2], true);
		s.setSheetView(0, { zoom: 120 });
		const guid = s.addCustomView({ name: 'Saved' });
		s.setHidden(0, 'row', [2], false);
		s.setSheetView(0, { zoom: 100 });
		const reopened = await loadXlsx(await saveXlsx(wb));
		expect(reopened.customViews?.map((v) => [v.guid, v.name])).toEqual([[guid, 'Saved']]);
		const again = createEditSession(reopened);
		again.showCustomView(guid);
		expect(hiddenRows(reopened.sheets[0]!)).toEqual([2]);
		expect(reopened.sheets[0]!.view.zoom).toBe(120);
		expect(sheet.view.zoom).toBe(100);
		again.deleteCustomView(guid);
		const last = await loadXlsx(await saveXlsx(reopened));
		expect(last.customViews ?? []).toEqual([]);
		expect(last.sheets[0]!.customViews).toBeUndefined();
	});
});
