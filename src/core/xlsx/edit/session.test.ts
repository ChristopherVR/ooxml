import { describe, expect, it, vi } from 'vitest';
import { parseAddress, parseRange } from '../address.js';
import { getCell } from '../cells.js';
import type { Workbook } from '../model.js';
import { createWorkbook } from '../workbook.js';
import { createEditSession } from './session.js';
import type { WorkbookChange } from './types.js';

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
const value = (wb: Workbook, ref: string, sheet = 0) => {
	const ws = wb.sheets[sheet];
	if (!ws) throw new Error('sheet');
	return getCell(ws, A(ref).row, A(ref).col)?.value ?? null;
};
const setup = (sheets = ['Sheet1']) => {
	const wb = createWorkbook({ sheets });
	return { wb, s: createEditSession(wb, { recalc: false }) };
};

describe('history', () => {
	it('undoes and redoes a value edit', () => {
		const { wb, s } = setup();
		s.setCellValue(0, 0, 0, 5);
		expect(value(wb, 'A1')).toBe(5);
		expect(s.canUndo()).toBe(true);
		expect(s.undo()).toBe(true);
		expect(value(wb, 'A1')).toBeNull();
		expect(wb.sheets[0]?.rows.size).toBe(0);
		expect(s.canRedo()).toBe(true);
		expect(s.redo()).toBe(true);
		expect(value(wb, 'A1')).toBe(5);
	});
	it('returns false when there is nothing to undo or redo', () => {
		const { s } = setup();
		expect(s.undo()).toBe(false);
		expect(s.redo()).toBe(false);
		expect(s.undoLabel()).toBeUndefined();
	});
	it('labels steps', () => {
		const { s } = setup();
		s.setCellInput(0, 1, 1, 'x');
		expect(s.undoLabel()).toBe('Typing in B2');
		s.undo();
		expect(s.redoLabel()).toBe('Typing in B2');
	});
	it('clears the redo stack on a new edit', () => {
		const { s } = setup();
		s.setCellValue(0, 0, 0, 1);
		s.undo();
		s.setCellValue(0, 0, 0, 2);
		expect(s.canRedo()).toBe(false);
	});
	it('undoes several steps in order', () => {
		const { wb, s } = setup();
		s.setCellValue(0, 0, 0, 1);
		s.setCellValue(0, 0, 0, 2);
		s.setCellValue(0, 0, 0, 3);
		s.undo();
		expect(value(wb, 'A1')).toBe(2);
		s.undo();
		expect(value(wb, 'A1')).toBe(1);
		s.redo();
		s.redo();
		expect(value(wb, 'A1')).toBe(3);
	});
	it('caps the history', () => {
		const wb = createWorkbook();
		const s = createEditSession(wb, { recalc: false, historyLimit: 3 });
		for (let i = 0; i < 5; i++) s.setCellValue(0, 0, 0, i);
		let undone = 0;
		while (s.undo()) undone++;
		expect(undone).toBe(3);
		expect(value(wb, 'A1')).toBe(1);
	});
	it('groups a batch into one step', () => {
		const { wb, s } = setup();
		s.batch('Two cells', () => {
			s.setCellValue(0, 0, 0, 1);
			s.setCellValue(0, 0, 1, 2);
			s.insertRows(0, 0, 1);
		});
		expect(value(wb, 'A2')).toBe(1);
		expect(s.undoLabel()).toBe('Two cells');
		s.undo();
		expect(value(wb, 'A1')).toBeNull();
		expect(value(wb, 'B1')).toBeNull();
		expect(s.canUndo()).toBe(false);
		s.redo();
		expect(value(wb, 'A2')).toBe(1);
		expect(value(wb, 'B2')).toBe(2);
	});
	it('flattens nested batches', () => {
		const { s } = setup();
		s.batch('outer', () => {
			s.batch('inner', () => s.setCellValue(0, 0, 0, 1));
			s.setCellValue(0, 1, 0, 1);
		});
		expect(s.undoLabel()).toBe('outer');
		s.undo();
		expect(s.canUndo()).toBe(false);
	});
	it('rolls a failed batch back', () => {
		const { wb, s } = setup();
		expect(() =>
			s.batch('fails', () => {
				s.setCellValue(0, 0, 0, 1);
				throw new Error('boom');
			}),
		).toThrow('boom');
		expect(value(wb, 'A1')).toBeNull();
		expect(s.canUndo()).toBe(false);
	});
	it('rolls a failing command back and records nothing', () => {
		const { wb, s } = setup(['A', 'B']);
		expect(() => s.renameSheet(0, 'B')).toThrow();
		expect(wb.sheets[0]?.name).toBe('A');
		expect(s.canUndo()).toBe(false);
	});
	it('restores sheets in place so views keep their references', () => {
		const { wb, s } = setup();
		const sheet = wb.sheets[0];
		s.setColumnWidth(0, [0], 20);
		s.undo();
		expect(wb.sheets[0]).toBe(sheet);
		expect(sheet?.columns).toEqual([]);
	});
	it('does not allow undo inside a batch', () => {
		const { s } = setup();
		expect(() => s.batch('x', () => s.undo())).toThrow();
	});
});

describe('change events', () => {
	it('emits one change per step with the sheet and ranges', () => {
		const { s } = setup();
		const changes: WorkbookChange[] = [];
		const off = s.onChange((c) => changes.push(c));
		s.setCellValue(0, 2, 3, 'x');
		expect(changes).toEqual([
			{ kind: 'cells', label: 'Edit D3', sheet: 0, ranges: [R('D3')], structural: false },
		]);
		s.insertRows(0, 0, 1);
		expect(changes[1]?.structural).toBe(true);
		off();
		s.setCellValue(0, 0, 0, 1);
		expect(changes).toHaveLength(2);
	});
	it('emits undo and redo changes and one change per batch', () => {
		const { s } = setup();
		const listener = vi.fn();
		s.onChange(listener);
		s.batch('b', () => {
			s.setCellValue(0, 0, 0, 1);
			s.setCellValue(0, 0, 1, 1);
		});
		expect(listener).toHaveBeenCalledTimes(1);
		s.undo();
		s.redo();
		expect(listener.mock.calls.map((c) => (c[0] as WorkbookChange).kind)).toEqual([
			'batch',
			'undo',
			'redo',
		]);
	});
});

describe('recalculation', () => {
	it('recalculates formulas after edits and undo', () => {
		const wb = createWorkbook();
		const s = createEditSession(wb);
		s.setCellValue(0, 0, 0, 2);
		s.setCellInput(0, 0, 1, '=A1*10');
		expect(value(wb, 'B1')).toBe(20);
		s.setCellValue(0, 0, 0, 3);
		expect(value(wb, 'B1')).toBe(30);
		s.undo();
		expect(value(wb, 'B1')).toBe(20);
	});
	it('keeps formulas right after inserting rows', () => {
		const wb = createWorkbook();
		const s = createEditSession(wb);
		s.setCellValue(0, 0, 0, 4);
		s.setCellInput(0, 1, 0, '=A1+1');
		s.insertRows(0, 0, 2);
		expect(wb.sheets[0]?.rows.get(3)?.get(0)?.formula).toBe('A3+1');
		expect(value(wb, 'A4')).toBe(5);
	});
});
