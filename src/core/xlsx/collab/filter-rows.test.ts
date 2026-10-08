import { describe, expect, it } from 'vitest';
import * as Y from 'yjs';
import { parseRange } from '../address';
import { getCell } from '../cells';
import { createEditSession } from '../edit/session';
import { createWorkbook } from '../workbook';
import { readSharedWorkbook, writeSharedWorkbook } from './adapter';

describe('shared filter visibility', () => {
	it('keeps filter and manual hiding distinct on a second client', () => {
		const workbook = createWorkbook();
		const session = createEditSession(workbook, { autoRowHeight: false });
		session.setRangeValues(0, { row: 0, col: 0 }, [
			['Kind', 'Amount'],
			['keep', 10],
			['drop', 20],
		]);
		session.setCellInput(0, 4, 0, '=SUBTOTAL(9,B2:B3)');
		session.setCellInput(0, 5, 0, '=SUBTOTAL(109,B2:B3)');
		session.setAutoFilter(0, parseRange('A1:B3')!);
		session.filterColumn(0, 0, ['keep']);
		session.setHidden(0, 'row', [1, 2], true);
		const shared = new Y.Doc();
		writeSharedWorkbook(shared, workbook);
		const received = readSharedWorkbook(shared);
		createEditSession(received, { autoRowHeight: false });
		expect(received.sheets[0]!.rowInfo).toEqual(workbook.sheets[0]!.rowInfo);
		expect(getCell(received.sheets[0]!, 4, 0)?.value).toBe(10);
		expect(getCell(received.sheets[0]!, 5, 0)?.value).toBe(0);
	});
});
