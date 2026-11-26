import { describe, expect, it, vi } from 'vitest';
import { loadedXPath, pendingXPath } from '#xpath-runtime';
import { createEditSession } from '../../edit/session';
import { loadXlsx } from '../../read/load';
import { saveXlsx } from '../../write/save';
import { prepareFormula, prepareWorkbookFormulas } from '../prepare-filterxml';
import { book, engine, grid } from '../test-helpers';

const library = vi.hoisted(() => {
	let unblock: () => void = () => {};
	const blocked = new Promise<void>((resolve) => {
		unblock = resolve;
	});
	return { blocked, unblock, loaded: vi.fn() };
});

vi.mock('xpath', async (importOriginal) => {
	library.loaded();
	await library.blocked;
	return importOriginal();
});

describe('FILTERXML browser loading', () => {
	it('keeps the dependency idle for ordinary workbooks and finishes a first edit before save', async () => {
		const node = vi.spyOn(process, 'getBuiltinModule').mockReturnValue(undefined);
		try {
			const workbook = await loadXlsx(
				await saveXlsx(book({ A1: '<r><a>1</a><a>2</a></r>', B1: '="FILTERXML"' })),
			);
			expect(prepareFormula('="FILTERXML(A1,B1)"')).toBeUndefined();
			expect(library.loaded).not.toHaveBeenCalled();
			expect(loadedXPath()).toBeUndefined();

			const session = createEditSession(workbook);
			session.setCellInput(0, 0, 2, '=FILTERXML(A1,"//a")');
			expect(grid(workbook, 'C1:C2')).toEqual([[{ error: '#BUSY!' }], [null]]);
			const ready = pendingXPath(workbook);
			expect(ready).toBeDefined();
			expect(prepareFormula('=_xlfn.FILTERXML(A1,"//a")')).toBe(ready);
			expect(prepareWorkbookFormulas(workbook)).toBe(ready);
			const saved = saveXlsx(workbook);
			library.unblock();
			const reopened = await loadXlsx(await saved);
			expect(library.loaded).toHaveBeenCalledTimes(1);
			expect(grid(reopened, 'C1:C2')).toEqual([[1], [2]]);
			engine(reopened).recalculateAll();
			expect(grid(reopened, 'C1:C2')).toEqual([[1], [2]]);
			session.undo();
			expect(grid(workbook, 'C1:C2')).toEqual([[null], [null]]);
			session.redo();
			expect(grid(workbook, 'C1:C2')).toEqual([[1], [2]]);
		} finally {
			library.unblock();
			node.mockRestore();
		}
	});
});
