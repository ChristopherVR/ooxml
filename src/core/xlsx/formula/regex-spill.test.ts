import { describe, expect, it } from 'vitest';
import { getCell } from '../cells';
import { createEditSession } from '../edit/session';
import { cellView } from '../layout/cell-view';
import { loadXlsx } from '../read/index';
import { saveXlsx } from '../write/index';
import { FUNCTION_CATALOG } from './functions/registry';
import { book, E, engine, get, grid } from './test-helpers';

describe('REGEXEXTRACT editing and grid integration', () => {
	it('shrinks a spill in a workbook whose inputs were entered through the session', () => {
		const wb = book();
		const edit = createEditSession(wb);
		edit.setCellInput(0, 0, 0, 'order 001, 22, 333');
		edit.setCellInput(0, 0, 1, '=REGEXEXTRACT(A1,"[0-9]+",1)');
		expect(grid(wb, 'B1:B3')).toEqual([['001'], ['22'], ['333']]);
		edit.setCellInput(0, 0, 0, 'order 99');
		expect(grid(wb, 'B1:B3')).toEqual([['99'], [null], [null]]);
	});
	it('recalculates, shrinks and restores spills through undo and redo', () => {
		const wb = book({ A1: 'order 001, 22, 333' });
		const edit = createEditSession(wb);
		edit.setCellInput(0, 1, 0, '=REGEXEXTRACT(A1,"[0-9]+",1)');
		expect(grid(wb, 'A2:A4')).toEqual([['001'], ['22'], ['333']]);
		expect(cellView(wb, 0, 1, 0).text).toBe('001');
		expect(cellView(wb, 0, 2, 0).text).toBe('22');
		edit.setCellInput(0, 0, 0, 'order 99');
		expect(grid(wb, 'A2:A4')).toEqual([['99'], [null], [null]]);
		edit.undo();
		expect(grid(wb, 'A2:A4')).toEqual([['001'], ['22'], ['333']]);
		edit.redo();
		expect(grid(wb, 'A2:A4')).toEqual([['99'], [null], [null]]);
	});

	it('reports a blocked spill without replacing existing cell content', () => {
		const wb = book({ A1: '=REGEXEXTRACT("a1b2","[0-9]",1)', A2: 'keep' });
		engine(wb).recalculateAll();
		expect(get(wb, 'A1')).toEqual(E.SPILL);
		expect(get(wb, 'A2')).toBe('keep');
	});

	it('preserves the formula and text results through save, load and recalculation', async () => {
		const formula = 'REGEXEXTRACT("John Smith","(\\w+) (\\w+)",2)';
		const wb = book({ A1: `=${formula}` });
		engine(wb).recalculateAll();
		const loaded = await loadXlsx(await saveXlsx(wb));
		expect(getCell(loaded.sheets[0]!, 0, 0)).toMatchObject({ formula, dynamicArray: true });
		expect(grid(loaded, 'A1:B1')).toEqual([['John', 'Smith']]);
		engine(loaded).recalculateAll();
		expect(grid(loaded, 'A1:B1')).toEqual([['John', 'Smith']]);
	});

	it('explains return modes in the catalog used by Insert Function and autocomplete', () => {
		expect(FUNCTION_CATALOG.find((fn) => fn.name === 'REGEXEXTRACT')?.description).toContain(
			'all matches (1)',
		);
	});
});
