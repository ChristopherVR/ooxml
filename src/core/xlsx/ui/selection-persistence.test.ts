import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { createWorkbook, loadXlsx, parseRange, saveXlsx } from '../index';
import {
	createSelectionModel,
	initialSelection,
	setSelectionPane,
	SheetSelections,
} from './selection';

describe('saved worksheet selections', () => {
	it('writes disjoint ranges and the active range index, without sharing mutable selection objects', async () => {
		const workbook = createWorkbook();
		const sheet = workbook.sheets[0]!;
		const current = createSelectionModel();
		current.set({
			active: { row: 7, col: 5 },
			ranges: [parseRange('B2:C3')!, parseRange('F8:G9')!],
		});
		new SheetSelections().remember(sheet, current.get());
		const bytes = await saveXlsx(workbook);
		const zip = await JSZip.loadAsync(bytes);
		expect(await zip.file('xl/worksheets/sheet1.xml')!.async('string')).toContain(
			'activeCell="F8" activeCellId="1" sqref="B2:C3 F8:G9"',
		);
		const reloaded = await loadXlsx(bytes);
		expect(initialSelection(0, reloaded)).toMatchObject(current.get());
		current.get().active.row = 99;
		current.get().ranges[0]!.start.row = 99;
		expect(sheet.view.selection!.active.row).toBe(7);
		expect(sheet.view.selection!.ranges[0]!.start.row).toBe(1);
	});

	it('keeps each split pane selection across save and reload', async () => {
		const workbook = createWorkbook();
		const sheet = workbook.sheets[0]!;
		sheet.view.split = { x: 2000, y: 900 };
		const current = createSelectionModel();
		const memory = new SheetSelections();
		current.set({ active: { row: 8, col: 6 } });
		memory.remember(sheet, current.get());
		setSelectionPane(sheet, 'topLeft');
		current.set({ active: { row: 1, col: 1 } });
		memory.remember(sheet, current.get());
		const reloaded = await loadXlsx(await saveXlsx(workbook));
		const view = reloaded.sheets[0]!.view;
		expect(view.split?.activePane).toBe('topLeft');
		expect(view.split?.paneSelections?.bottomRight?.active).toEqual({ row: 8, col: 6 });
		expect(view.selection?.active).toEqual({ row: 1, col: 1 });
		setSelectionPane(sheet, 'bottomRight');
		current.set({ active: { row: 10, col: 7 } });
		memory.remember(sheet, current.get());
		const again = (await loadXlsx(await saveXlsx(workbook))).sheets[0]!.view;
		expect(again.split?.paneSelections?.topLeft?.active).toEqual({ row: 1, col: 1 });
		expect(again.selection?.active).toEqual({ row: 10, col: 7 });
	});

	it('saves a cell in the frozen area in the scrollable pane, as Excel does', async () => {
		const workbook = createWorkbook();
		const sheet = workbook.sheets[0]!;
		sheet.view.freeze = { rows: 3, cols: 2 };
		setSelectionPane(sheet, 'topLeft');
		const current = createSelectionModel();
		current.set({ active: { row: 1, col: 1 } });
		new SheetSelections().remember(sheet, current.get());
		const bytes = await saveXlsx(workbook);
		const zip = await JSZip.loadAsync(bytes);
		const xml = await zip.file('xl/worksheets/sheet1.xml')!.async('string');
		expect(xml).toContain('activePane="bottomRight"');
		expect(xml).toContain('<selection pane="bottomRight" activeCell="B2" sqref="B2"/>');
		expect((await loadXlsx(bytes)).sheets[0]!.view.selection).toEqual(sheet.view.selection);
	});
});
