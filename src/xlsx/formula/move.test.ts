import { describe, expect, it } from 'vitest';
import { parseRange } from '../address.js';
import { deleteSheetInFormula, moveReferencesInFormula, type MoveSpec } from './move.js';

const range = (ref: string) => {
	const r = parseRange(ref);
	if (!r) throw new Error(ref);
	return r;
};

const move = (from: string, ref: string, to: string, dRow: number, dCol: number): MoveSpec => ({
	fromSheet: from,
	range: range(ref),
	toSheet: to,
	dRow,
	dCol,
});

describe('moveReferencesInFormula', () => {
	it('moves references wholly inside the cut range, absolute ones too', () => {
		const m = move('Sheet1', 'A1:B2', 'Sheet1', 4, 1);
		expect(moveReferencesInFormula('A1+$B$2+SUM(A1:B2)+C3', 'Sheet1', m)).toBe(
			'B5+$C$6+SUM(B5:C6)+C3',
		);
	});
	it('leaves references that only overlap the cut range', () => {
		const m = move('Sheet1', 'A1:A2', 'Sheet1', 0, 3);
		expect(moveReferencesInFormula('SUM(A1:A3)', 'Sheet1', m)).toBe('SUM(A1:A3)');
	});
	it('qualifies references moved to another sheet and handles quoted names', () => {
		const m = move('Sheet1', 'A1', 'My Data', 1, 0);
		expect(moveReferencesInFormula('A1*2', 'Sheet1', m)).toBe("'My Data'!A2*2");
		expect(moveReferencesInFormula('Sheet1!A1', 'Other', m)).toBe("'My Data'!A2");
		expect(moveReferencesInFormula('A1', 'Other', m)).toBe('A1');
	});
	it('turns references to overwritten destination cells into #REF!', () => {
		const m = move('Sheet1', 'A1', 'Sheet1', 0, 2);
		expect(moveReferencesInFormula('C1+A1', 'Sheet1', m)).toBe('#REF!+C1');
	});
	it('qualifies untouched references when the formula itself changes sheet', () => {
		const m = move('Sheet1', 'B1', 'Sheet2', 0, 0);
		expect(moveReferencesInFormula('A1+B1', 'Sheet1', m, 'Sheet2')).toBe('Sheet1!A1+B1');
	});
});

describe('deleteSheetInFormula', () => {
	it('turns references to the deleted sheet and its tables into #REF!', () => {
		expect(deleteSheetInFormula("Data!A1+'data'!B2:C3+A1", 'Data')).toBe('#REF!+#REF!+A1');
		expect(deleteSheetInFormula('SUM(Sales[Amount])+ROWS(Sales)', 'Data', ['Sales'])).toBe(
			'SUM(#REF!)+ROWS(#REF!)',
		);
		expect(deleteSheetInFormula('Data!Rate*2', 'Data')).toBe('#REF!*2');
	});
	it('keeps other sheets, 3D references and unparsable text', () => {
		expect(deleteSheetInFormula('Other!A1+SUM(Data:Other!A1)', 'Data')).toBe(
			'Other!A1+SUM(Data:Other!A1)',
		);
		expect(deleteSheetInFormula('SUM((', 'Data')).toBe('SUM((');
	});
});
