import { describe, expect, it } from 'vitest';
import { createWorkbook, getCell } from '../index';
import { createEditSession } from '../edit';
import {
	replaceAllValidated,
	replaceOneValidated,
	validateReplacements,
} from './replace-validation';

function setup() {
	const workbook = createWorkbook();
	const session = createEditSession(workbook);
	session.setCellInput(0, 0, 0, 'a1');
	session.setCellInput(0, 1, 0, 'a2');
	session.setCellInput(0, 2, 0, 'a=3');
	workbook.sheets[0]!.dataValidations.push({
		ranges: [{ start: { row: 0, col: 0 }, end: { row: 0, col: 0 } }],
		type: 'whole',
		operator: 'greaterThan',
		formula1: '10',
		showErrorMessage: true,
		error: 'Too small',
		errorTitle: 'Value',
	});
	return { workbook, session };
}
const query = { text: 'a', lookIn: 'formulas' as const };

describe('replace validation', () => {
	it('splits matches into allowed and refused with the rule message', () => {
		const { session } = setup();
		const result = validateReplacements(session, query, '', session.findAll(query));
		expect(result.allowed.map((m) => m.row)).toEqual([1, 2]);
		expect(result.refused).toEqual([
			{ sheet: 0, row: 0, col: 0, ref: 'A1', value: '1', message: 'Too small', title: 'Value' },
		]);
	});

	it('lets a value that passes through', () => {
		const { session } = setup();
		expect(validateReplacements(session, query, '9', session.findAll(query)).refused).toEqual([]);
	});

	it('replaces what passes and reports what is refused', () => {
		const { workbook, session } = setup();
		const result = replaceAllValidated(session, query, '');
		expect(result.replaced).toBe(2);
		expect(result.refused).toHaveLength(1);
		const ws = workbook.sheets[0]!;
		expect(getCell(ws, 0, 0)?.value).toBe('a1');
		expect(getCell(ws, 1, 0)?.value).toBe(2);
	});

	it('replaceOneValidated refuses a single failing cell', () => {
		const { session } = setup();
		const at = session.findAll(query)[0]!;
		expect(replaceOneValidated(session, query, '', at)).toMatchObject({ replaced: false });
		expect(replaceOneValidated(session, query, '50', at)).toMatchObject({ replaced: true });
	});
});
