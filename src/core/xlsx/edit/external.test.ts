import { describe, expect, it } from 'vitest';
import { getCell, putCell } from '../cells';
import { createWorkbook } from '../workbook';
import { createEditSession } from './session';
import type { WorkbookChange } from './types';

describe('applyExternal', () => {
	it('recalculates and announces a change without recording history', () => {
		const session = createEditSession(createWorkbook());
		session.setCellInput(0, 0, 1, '=A1*2');
		const changes: WorkbookChange[] = [];
		session.onChange((change) => changes.push(change));
		const sheet = session.workbook.sheets[0]!;
		const applied = session.applyExternal(() => {
			putCell(sheet, 0, 0, { value: 21 });
			return { label: 'Remote', sheet: 0, cells: [{ sheet: 0, row: 0, col: 0 }] };
		});
		expect(applied).toBe(true);
		expect(getCell(sheet, 0, 1)?.value).toBe(42);
		expect(changes).toEqual([
			{ kind: 'remote', label: 'Remote', structural: false, external: true, sheet: 0 },
		]);
		expect(session.undoLabel()).toBe('Typing in B1');
		expect(session.applyExternal(() => undefined)).toBe(false);
	});

	it('refuses to run inside a batch', () => {
		const session = createEditSession(createWorkbook());
		expect(() => session.batch('outer', () => session.applyExternal(() => undefined))).toThrow(
			/inside an edit/,
		);
	});
});
