import { describe, expect, it } from 'vitest';
import {
	createEditSession,
	createWorkbook,
	getCell,
	getValue,
	loadXlsx,
	saveXlsx,
	styleAt,
} from '../../index';
import { createScenarioSummary, ScenarioSummaryError } from './scenario-summary';

// The layout below is what Excel 16 wrote for Scenarios.CreateSummary(1, B3:B4) on this sheet:
// Price=10 (named), Qty=5, Total==Price*Qty, Tax==Total*0.1 and the scenarios Base (10, 5, with a
// comment), High (20, 8) and Low (5, 3).
function build() {
	const wb = createWorkbook({ sheets: ['Budget'] });
	const session = createEditSession(wb);
	session.setCellInput(0, 0, 0, 'Price');
	session.setCellInput(0, 0, 1, '10');
	session.setCellInput(0, 1, 0, 'Qty');
	session.setCellInput(0, 1, 1, '5');
	session.setCellInput(0, 2, 1, '=B1*B2');
	session.setCellInput(0, 3, 1, '=B3*0.1');
	wb.definedNames.push({ name: 'Price', formula: 'Budget!$B$1' });
	session.addScenario(0, { name: 'Base', cells: 'B1:B2', comment: 'note base' });
	session.addScenario(0, { name: 'High', cells: 'B1:B2', values: { B1: '20', B2: '8' } });
	session.addScenario(0, { name: 'Low', cells: 'B1:B2', values: { B1: '5', B2: '3' } });
	return { wb, session };
}

const text = (wb: ReturnType<typeof build>['wb'], sheet: number, ref: string) => {
	const m = /^([A-Z]+)(\d+)$/.exec(ref)!;
	const col = [...m[1]!].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
	return getValue(wb.sheets[sheet]!, Number(m[2]) - 1, col);
};

describe('scenario summary', () => {
	it('builds the report Excel builds', () => {
		const { wb, session } = build();
		const index = createScenarioSummary(session, { sheet: 0, resultCells: 'B3:B4' });
		expect(index).toBe(0);
		expect(wb.sheets.map((s) => s.name)).toEqual(['Scenario Summary', 'Budget']);
		const t = (ref: string) => text(wb, 0, ref);
		expect(t('B2')).toBe('Scenario Summary');
		expect([t('D3'), t('E3'), t('F3'), t('G3')]).toEqual([
			'Current Values:',
			'Base',
			'High',
			'Low',
		]);
		expect(t('E4')).toBe('note base');
		expect(t('B5')).toBe('Changing Cells:');
		expect([t('C6'), t('D6'), t('E6'), t('F6'), t('G6')]).toEqual(['Price', 10, 10, 20, 5]);
		expect([t('C7'), t('D7'), t('E7'), t('F7'), t('G7')]).toEqual(['$B$2', 5, 5, 8, 3]);
		expect(t('B8')).toBe('Result Cells:');
		expect([t('C9'), t('D9'), t('E9'), t('F9'), t('G9')]).toEqual(['$B$3', 50, 50, 160, 15]);
		expect([t('C10'), t('D10'), t('E10'), t('F10'), t('G10')]).toEqual(['$B$4', 5, 5, 16, 1.5]);
		expect(String(t('B11'))).toBe(
			'Notes:  Current Values column represents values of changing cells at',
		);
		expect(t('B13')).toBe('scenario are highlighted in gray.');
		// The source sheet is left as it was.
		expect(getValue(wb.sheets[1]!, 0, 1)).toBe(10);
		expect(getValue(wb.sheets[1]!, 2, 1)).toBe(50);
	});

	it('styles, outlines and hides rows like Excel', () => {
		const { wb, session } = build();
		createScenarioSummary(session, { sheet: 0, resultCells: 'B3:B4' });
		const sheet = wb.sheets[0]!;
		const style = (row: number, col: number) => styleAt(wb, getCell(sheet, row, col)?.styleId);
		expect(style(1, 1)).toMatchObject({
			font: { bold: true, size: 12, color: { indexed: 9 } },
			fill: { type: 'pattern', pattern: 'solid', fgColor: { indexed: 20 } },
			border: { top: { style: 'medium' } },
			alignment: { horizontal: 'left' },
		});
		expect(style(1, 4)).toMatchObject({ alignment: { horizontal: 'right' }, font: { size: 10 } });
		expect(style(4, 1)).toMatchObject({ font: { bold: true, color: { indexed: 18 } } });
		expect(style(5, 4).fill).toMatchObject({ fgColor: { indexed: 22 } });
		expect(style(5, 3).fill).not.toMatchObject({ fgColor: { indexed: 22 } });
		expect(style(9, 3).border).toMatchObject({ bottom: { style: 'medium' } });
		expect(sheet.view.showGridLines).toBe(false);
		expect(sheet.rowInfo.get(3)).toMatchObject({ hidden: true, outlineLevel: 1, height: 42 });
		expect([5, 6, 8, 9].map((r) => sheet.rowInfo.get(r)?.outlineLevel)).toEqual([1, 1, 1, 1]);
		expect(sheet.rowInfo.get(4)?.outlineLevel).toBeUndefined();
		expect(sheet.columns.find((c) => c.min <= 3 && c.max >= 3)?.outlineLevel).toBe(1);
	});

	it('is one undo step and survives a save', async () => {
		const { wb, session } = build();
		createScenarioSummary(session, { sheet: 0, resultCells: 'B3:B4' });
		const reopened = await loadXlsx(await saveXlsx(wb));
		expect(reopened.sheets[0]!.name).toBe('Scenario Summary');
		expect(getValue(reopened.sheets[0]!, 8, 5)).toBe(160);
		session.undo();
		expect(wb.sheets.map((s) => s.name)).toEqual(['Budget']);
	});

	it('names later summaries like Excel and refuses a sheet without scenarios', () => {
		const { session } = build();
		createScenarioSummary(session, { sheet: 0, resultCells: 'B3' });
		createScenarioSummary(session, { sheet: 1, resultCells: 'B3' });
		expect(session.workbook.sheets.map((s) => s.name)).toEqual([
			'Scenario Summary',
			'Scenario Summary 2',
			'Budget',
		]);
		expect(() => createScenarioSummary(session, { sheet: 0 })).toThrow(ScenarioSummaryError);
		expect(() => createScenarioSummary(session, { sheet: 2, resultCells: 'nonsense' })).toThrow(
			ScenarioSummaryError,
		);
	});

	it('copies the number format of the source cells', () => {
		const { wb, session } = build();
		session.applyStyle(0, [{ start: { row: 0, col: 1 }, end: { row: 0, col: 1 } }], {
			numFmt: '$#,##0.00',
		});
		createScenarioSummary(session, { sheet: 0, resultCells: 'B3' });
		const sheet = wb.sheets[0]!;
		const fmt = (row: number, col: number) => styleAt(wb, getCell(sheet, row, col)?.styleId).numFmt;
		expect(fmt(5, 3)).toBe('$#,##0.00');
		expect(fmt(5, 4)).toBe('$#,##0.00');
		expect(fmt(6, 3)).toBe('General');
	});
});

describe('merge scenarios', () => {
	it('appends scenarios of another sheet and renames taken names like Excel', () => {
		const { wb, session } = build();
		session.addSheet('Other');
		session.addScenario(1, { name: 'Base', cells: 'B1:B2', values: { B1: '111', B2: '222' } });
		session.addScenario(1, { name: 'Extra', cells: 'B1:B2,B5', comment: 'extra comment' });
		const now = new Date();
		const today = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')}`;
		const added = session.mergeScenarios(0, 1);
		expect(added).toBe(2);
		const names = wb.sheets[0]!.scenarios!.list.map((s) => s.name);
		expect(names).toEqual(['Base', 'High', 'Low', `Base ${today}`, 'Extra']);
		expect(wb.sheets[0]!.scenarios!.list[3]!.cells.map((c) => c.value)).toEqual(['111', '222']);
		expect(wb.sheets[0]!.scenarios!.list[4]!.comment).toBe('extra comment');
		session.undo();
		expect(wb.sheets[0]!.scenarios!.list).toHaveLength(3);
		expect(() => session.mergeScenarios(0, 0)).toThrow();
	});
});
