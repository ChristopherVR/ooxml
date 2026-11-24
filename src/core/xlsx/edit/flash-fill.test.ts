import { describe, expect, it } from 'vitest';
import { parseRange } from '../address';
import { getCell } from '../cells';
import { createWorkbook } from '../workbook';
import { createEditSession } from './session';

function makeSession(rows: (string | number | null)[][]) {
	const workbook = createWorkbook();
	const session = createEditSession(workbook, { autoRowHeight: false });
	session.setRangeValues(0, { row: 0, col: 0 }, rows);
	return { workbook, session, sheet: workbook.sheets[0]! };
}

describe('flashFill', () => {
	it('extracts tokens by whitespace delimiter and auto-skips header row', () => {
		const { session, sheet } = makeSession([
			['Full Name', 'First'],
			['Alice Smith', 'Alice'],
			['Bob Jones', null],
			['Cara Lee', null],
			['Dan Brown', null],
		]);

		const res = session.flashFill(0, { row: 1, col: 1 });
		expect(res.filled).toBe(3);
		expect(res.range).toEqual(parseRange('B3:B5'));
		expect(getCell(sheet, 0, 1)?.value).toBe('First');
		expect(getCell(sheet, 2, 1)?.value).toBe('Bob');
		expect(getCell(sheet, 3, 1)?.value).toBe('Cara');
		expect(getCell(sheet, 4, 1)?.value).toBe('Dan');
	});

	it('extracts tokens by punctuation delimiters and between delimiters', () => {
		const { session, sheet } = makeSession([
			['alice@example.com', 'example'],
			['bob@acme.org', null],
			['cara@widgets.io', null],
		]);

		const res = session.flashFill(0, { row: 0, col: 1 });
		expect(res.filled).toBe(2);
		expect(getCell(sheet, 1, 1)?.value).toBe('acme');
		expect(getCell(sheet, 2, 1)?.value).toBe('widgets');
	});

	it('supports prefix and suffix extraction', () => {
		const { session, sheet } = makeSession([
			['INV-2026-0042', '0042'],
			['INV-2026-0199', null],
			['INV-2026-1234', null],
		]);

		const res = session.flashFill(0, { row: 0, col: 1 });
		expect(res.filled).toBe(2);
		expect(getCell(sheet, 1, 1)?.value).toBe('0199');
		expect(getCell(sheet, 2, 1)?.value).toBe('1234');
	});

	it('concatenates multiple source columns with literal separators', () => {
		const { session, sheet } = makeSession([
			['Alice', 'Smith', 'Smith, Alice'],
			['Bob', 'Jones', null],
			['Cara', 'Lee', null],
		]);

		const res = session.flashFill(0, { row: 0, col: 2 });
		expect(res.filled).toBe(2);
		expect(getCell(sheet, 1, 2)?.value).toBe('Jones, Bob');
		expect(getCell(sheet, 2, 2)?.value).toBe('Lee, Cara');
	});

	it('builds email addresses combining initial, last name, lower case, and literal domain', () => {
		const { session, sheet } = makeSession([
			['Alice', 'Smith', 'asmith@acme.com'],
			['Bob', 'Jones', null],
			['Cara', 'Lee', null],
		]);

		const res = session.flashFill(0, { row: 0, col: 2 });
		expect(res.filled).toBe(2);
		expect(getCell(sheet, 1, 2)?.value).toBe('bjones@acme.com');
		expect(getCell(sheet, 2, 2)?.value).toBe('clee@acme.com');
	});

	it('supports UPPER, LOWER, and PROPER case transforms', () => {
		const upper = makeSession([
			['alice smith', 'ALICE SMITH'],
			['bob jones', null],
		]);
		expect(upper.session.flashFill(0, { row: 0, col: 1 }).filled).toBe(1);
		expect(getCell(upper.sheet, 1, 1)?.value).toBe('BOB JONES');

		const lower = makeSession([
			['ALICE SMITH', 'alice smith'],
			['BOB JONES', null],
		]);
		expect(lower.session.flashFill(0, { row: 0, col: 1 }).filled).toBe(1);
		expect(getCell(lower.sheet, 1, 1)?.value).toBe('bob jones');

		const proper = makeSession([
			['alice smith', 'Alice Smith'],
			['BOB JONES', null],
			['cARA lEE', null],
		]);
		expect(proper.session.flashFill(0, { row: 0, col: 1 }).filled).toBe(2);
		expect(getCell(proper.sheet, 1, 1)?.value).toBe('Bob Jones');
		expect(getCell(proper.sheet, 2, 1)?.value).toBe('Cara Lee');
	});

	it('extracts numeric tokens from mixed text', () => {
		const { session, sheet } = makeSession([
			['Order #1042 (USD)', '1042'],
			['Order #99 (EUR)', null],
			['Order #50812 (GBP)', null],
		]);

		const res = session.flashFill(0, { row: 0, col: 1 });
		expect(res.filled).toBe(2);
		expect(getCell(sheet, 1, 1)?.value).toBe('99');
		expect(getCell(sheet, 2, 1)?.value).toBe('50812');
	});

	it('extracts and reformats date tokens', () => {
		const { session, sheet } = makeSession([
			['2026-03-15', 'Mar 2026'],
			['2025-11-04', null],
			['2024-01-22', null],
		]);

		const res = session.flashFill(0, { row: 0, col: 1 });
		expect(res.filled).toBe(2);
		expect(getCell(sheet, 1, 1)?.value).toBe('Nov 2025');
		expect(getCell(sheet, 2, 1)?.value).toBe('Jan 2024');
	});

	it('disambiguates with multiple examples (last token vs second token)', () => {
		const { session, sheet } = makeSession([
			['John Doe', 'Doe'],
			['Mary Jane Watson', 'Watson'],
			['Ada Lovelace', null],
			['Grace Brewster Murray Hopper', null],
		]);

		const res = session.flashFill(0, { row: 0, col: 1 });
		expect(res.filled).toBe(2);
		expect(getCell(sheet, 2, 1)?.value).toBe('Lovelace');
		expect(getCell(sheet, 3, 1)?.value).toBe('Hopper');
	});

	it('respects an explicit single-column target range', () => {
		const { session, sheet } = makeSession([
			['Alice Smith', 'Alice'],
			['Bob Jones', null],
			['Cara Lee', null],
			['Dan Brown', null],
		]);

		const res = session.flashFill(0, parseRange('B1:B3')!);
		expect(res.filled).toBe(2);
		expect(getCell(sheet, 1, 1)?.value).toBe('Bob');
		expect(getCell(sheet, 2, 1)?.value).toBe('Cara');
		expect(getCell(sheet, 3, 1)).toBeUndefined();
	});

	it('returns a no-op without recording history when no consistent pattern exists', () => {
		const { session, sheet } = makeSession([
			['Apple', 'Red'],
			['Banana', 'Yellow'],
			['Grape', null],
		]);

		const beforeUndo = session.undoLabel();
		const res = session.flashFill(0, { row: 0, col: 1 });
		expect(res).toEqual({ filled: 0, range: undefined });
		expect(getCell(sheet, 2, 1)).toBeUndefined();
		expect(session.undoLabel()).toBe(beforeUndo);
	});

	it('supports full undo and redo', () => {
		const { session, sheet } = makeSession([
			['Alice Smith', 'A. Smith'],
			['Bob Jones', null],
			['Cara Lee', null],
		]);

		const res = session.flashFill(0, { row: 0, col: 1 });
		expect(res.filled).toBe(2);
		expect(session.undoLabel()).toBe('Flash Fill');
		expect(getCell(sheet, 1, 1)?.value).toBe('B. Jones');
		expect(getCell(sheet, 2, 1)?.value).toBe('C. Lee');

		expect(session.undo()).toBe(true);
		expect(getCell(sheet, 1, 1)).toBeUndefined();
		expect(getCell(sheet, 2, 1)).toBeUndefined();

		expect(session.redo()).toBe(true);
		expect(getCell(sheet, 1, 1)?.value).toBe('B. Jones');
		expect(getCell(sheet, 2, 1)?.value).toBe('C. Lee');
	});

	it('rejects multi-column target ranges', () => {
		const { session } = makeSession([['Alice', 'A', null]]);
		expect(() => session.flashFill(0, parseRange('B1:C1')!)).toThrow(RangeError);
	});
});
