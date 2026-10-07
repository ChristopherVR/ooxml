import { describe, expect, it } from 'vitest';
import { parseAddress, parseRange } from '../address';
import { getCell } from '../cells';
import { styleAt } from '../styles';
import { loadXlsx } from '../read/index';
import { saveXlsx } from '../write/index';
import { createWorkbook } from '../workbook';
import { createEditSession } from './session';
import { parseTsv } from './clipboard-text';
import { parseHtmlTable } from './clipboard-html-parse';

const range = (ref: string) => {
	const result = parseRange(ref);
	if (!result) throw new Error(ref);
	return result;
};
const address = (ref: string) => {
	const result = parseAddress(ref);
	if (!result) throw new Error(ref);
	return result;
};
const setup = () => {
	const workbook = createWorkbook();
	const session = createEditSession(workbook);
	return { workbook, session };
};

describe('clipboard selection dimensions', () => {
	it('keeps trailing empty rows and columns in cells, TSV and HTML', () => {
		const { workbook, session } = setup();
		session.setCellValue(0, 0, 0, 'source');
		const payload = session.copy(0, range('A1:C3'));
		expect(payload.cells).toMatchObject({ rows: 3, cols: 3, source: { range: range('A1:C3') } });
		const expected = [
			['source', '', ''],
			['', '', ''],
			['', '', ''],
		];
		expect(parseTsv(payload.tsv)).toEqual(expected);
		const style = styleAt(workbook, 0);
		expect(
			parseHtmlTable(payload.html, style)?.data.map((row) => row.map((c) => c?.text ?? '')),
		).toEqual(expected);
	});

	it('keeps an entirely empty finite selection, even beyond the used area', () => {
		const { session } = setup();
		expect(session.copy(0, range('E5:G7')).cells).toMatchObject({
			rows: 3,
			cols: 3,
			source: { range: range('E5:G7') },
		});
	});

	it.each(['all', 'values', 'formulas', 'transpose'] as const)(
		'pastes trailing blanks with %s and restores them through undo/redo',
		(mode) => {
			const { workbook, session } = setup();
			const sheet = workbook.sheets[0]!;
			session.setCellValue(0, 0, 0, 7);
			const payload = session.copy(0, range('A1:C2'));
			const transposed = mode === 'transpose';
			const rows = transposed ? 3 : 2;
			const cols = transposed ? 2 : 3;
			const old = Array.from({ length: rows }, () => Array.from({ length: cols }, () => 'old'));
			session.setRangeValues(0, address('E5'), old);
			expect(session.paste(0, address('E5'), payload, mode)).toEqual(
				range(transposed ? 'E5:F7' : 'E5:G6'),
			);
			const values = () =>
				Array.from({ length: rows }, (_, r) =>
					Array.from({ length: cols }, (_, c) => getCell(sheet, 4 + r, 4 + c)?.value ?? null),
				);
			const expected = Array.from({ length: rows }, (_, r) =>
				Array.from({ length: cols }, (_, c) => (r === 0 && c === 0 ? 7 : null)),
			);
			expect(values()).toEqual(expected);
			session.undo();
			expect(values()).toEqual(old);
			session.redo();
			expect(values()).toEqual(expected);
		},
	);

	it('clips only the unbounded axis of whole rows and columns', () => {
		const { session } = setup();
		session.setRangeValues(0, address('A1'), [[1], [2]]);
		expect(session.copy(0, range('A:C')).cells).toMatchObject({ rows: 2, cols: 3 });
		expect(session.copy(0, range('1:4')).cells).toMatchObject({ rows: 4, cols: 1 });
	});

	it('moves references to blank cells inside a cut selection', () => {
		const { workbook, session } = setup();
		const sheet = workbook.sheets[0]!;
		session.setCellValue(0, 0, 0, 7);
		const payload = session.cut(0, range('A1:C2'));
		session.setCellInput(0, 0, 3, '=C2');
		session.paste(0, address('E5'), payload);
		expect(getCell(sheet, 0, 3)?.formula).toBe('G6');
		session.undo();
		expect(getCell(sheet, 0, 3)?.formula).toBe('C2');
	});

	it('saves and reopens the cleared destination cells', async () => {
		const { workbook, session } = setup();
		session.setCellValue(0, 0, 0, 7);
		const payload = session.copy(0, range('A1:B2'));
		session.setRangeValues(0, address('E5'), [
			['old', 'old'],
			['old', 'old'],
		]);
		session.paste(0, address('E5'), payload);
		const reopened = await loadXlsx(await saveXlsx(workbook));
		const sheet = reopened.sheets[0]!;
		expect(getCell(sheet, 4, 4)?.value).toBe(7);
		for (const [r, c] of [
			[4, 5],
			[5, 4],
			[5, 5],
		])
			expect(getCell(sheet, r!, c!)).toBeUndefined();
	});
});
