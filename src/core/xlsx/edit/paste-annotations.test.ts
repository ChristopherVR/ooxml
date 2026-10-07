import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { parseRange } from '../address.js';
import { getCell } from '../cells.js';
import type { DataValidation } from '../model.js';
import { createWorkbook } from '../workbook.js';
import { saveXlsx } from '../write/index.js';
import { loadXlsx } from '../read/index.js';
import { createEditSession } from './session.js';
import { validationAt } from './validation.js';
import type { PasteOptions } from './types.js';

interface NativeValidation {
	type: number;
	operator: number;
	formula1: string;
	formula2: string;
	allowBlank: boolean;
	showInputMessage: boolean;
	showErrorMessage: boolean;
	errorStyle: number;
	errorTitle: string;
	error: string;
	promptTitle: string;
	prompt: string;
	showDropDown?: boolean;
}
interface NativeCase extends Required<PasteOptions> {
	cells: {
		row: number;
		col: number;
		value: number | null;
		comment: { author: string; text: string } | null;
		validation: NativeValidation | null;
	}[];
}
const fixture = JSON.parse(
	readFileSync(new URL('./__fixtures__/excel-paste-annotations.json', import.meta.url), 'utf8'),
) as { cases: NativeCase[] };
const range = (ref: string) => parseRange(ref)!;
const setup = () => {
	const workbook = createWorkbook();
	const sheet = workbook.sheets[0]!;
	const session = createEditSession(workbook);
	session.setCellValue(0, 0, 0, 2);
	session.setCellInput(0, 1, 0, '=B2+$J$1');
	session.setCellValue(0, 0, 9, 3);
	const author = fixture.cases[0]!.cells[0]!.comment!.author;
	for (const [ref, text] of [
		['A1', 'source A1'],
		['B1', 'source blank B1'],
		['A2', 'source formula A2'],
	])
		session.setComment(0, range(ref!).start, text!, author);
	const base = {
		allowBlank: true,
		showInputMessage: true,
		showErrorMessage: true,
		errorStyle: 'stop' as const,
	};
	session.setDataValidation(
		0,
		{
			...base,
			ranges: [],
			type: 'whole',
			operator: 'between',
			formula1: '1',
			formula2: '5',
			promptTitle: 'source input',
			prompt: 'enter a number',
			errorTitle: 'source error',
			error: 'invalid number',
		},
		range('A1'),
	);
	session.setDataValidation(
		0,
		{ ...base, ranges: [], type: 'custom', operator: 'between', formula1: 'A1>0' },
		range('B1'),
	);
	session.setDataValidation(
		0,
		{
			...base,
			ranges: [],
			type: 'list',
			operator: 'between',
			formula1: '"x,y"',
			showDropDown: true,
		},
		range('B2'),
	);
	for (const ref of ['D4', 'E4', 'D5', 'E5']) {
		const area = range(ref);
		session.setCellValue(0, area.start.row, area.start.col, 9);
		session.setComment(0, area.start, `dest ${ref}`, author);
		session.setDataValidation(
			0,
			{ ...base, ranges: [], type: 'whole', operator: 'between', formula1: '0', formula2: '99' },
			area,
		);
	}
	return { workbook, sheet, session };
};
const normalized = (dv: DataValidation | undefined): NativeValidation | null => {
	if (!dv) return null;
	return {
		type: { whole: 1, decimal: 2, list: 3, date: 4, time: 5, textLength: 6, custom: 7, none: 0 }[
			dv.type
		],
		operator: 1,
		formula1: dv.formula1?.replace(/^=/, '').replace(/^"(.*)"$/, '$1') ?? '',
		formula2: dv.formula2 ?? '',
		allowBlank: dv.allowBlank ?? false,
		showInputMessage: dv.showInputMessage ?? true,
		showErrorMessage: dv.showErrorMessage ?? true,
		errorStyle: 1,
		errorTitle: dv.errorTitle ?? '',
		error: dv.error ?? '',
		promptTitle: dv.promptTitle ?? '',
		prompt: dv.prompt ?? '',
		...(dv.type === 'list' ? { showDropDown: dv.showDropDown ?? true } : {}),
	};
};

describe('Clipboard annotations recorded in Microsoft Excel', () => {
	it('moves notes and validation with cut cells and restores both sheets on undo', () => {
		const { workbook, sheet, session } = setup();
		const target = session.addSheet('Target');
		const payload = session.cut(0, range('B1'));
		session.paste(target, range('D4'), payload);
		expect(sheet.comments.some((c) => c.address.row === 0 && c.address.col === 1)).toBe(false);
		expect(validationAt(workbook, 0, 0, 1)).toBeUndefined();
		expect(workbook.sheets[target]!.comments[0]?.text).toBe('source blank B1');
		expect(validationAt(workbook, target, 3, 3)?.formula1).toBe('Sheet1!A1>0');
		session.undo();
		expect(sheet.comments.some((c) => c.text === 'source blank B1')).toBe(true);
		expect(validationAt(workbook, 0, 0, 1)?.formula1).toBe('A1>0');
		expect(workbook.sheets[target]!.comments).toEqual([]);
	});
	it('rebases surviving validation when paste removes its original anchor', () => {
		const { workbook, session } = setup();
		session.setDataValidation(0, { ranges: [], type: 'custom', formula1: 'D4>0' }, range('D4:D8'));
		const payload = session.copy(0, range('A2'));
		session.paste(0, range('D4'), payload, 'validation');
		expect(validationAt(workbook, 0, 3, 3)).toBeUndefined();
		expect(validationAt(workbook, 0, 4, 3)?.formula1).toBe('D5>0');
	});
	it('clips a multi-area validation to the copied range and retains its relative formula', () => {
		const { workbook, sheet, session } = setup();
		sheet.dataValidations = [
			{
				ranges: [range('C3:C5'), range('A1:A2')],
				type: 'custom',
				formula1: 'A1>0',
				showErrorMessage: true,
			},
		];
		const payload = session.copy(0, range('C4:C5'));
		session.paste(0, range('D4'), payload, 'validation');
		expect(validationAt(workbook, 0, 3, 3)?.formula1).toBe('D4>0');
		expect(session.validate(0, 4, 3, -1).ok).toBe(false);
	});
	it('rejects annotation-only paste without metadata before adding history', () => {
		const { sheet, session } = setup();
		const before = structuredClone(sheet.comments);
		const label = session.undoLabel();
		for (const mode of ['comments', 'validation'] as const)
			expect(() => session.paste(0, range('D4'), '2', mode)).toThrow(RangeError);
		expect(sheet.comments).toEqual(before);
		expect(session.undoLabel()).toBe(label);
	});
	it.each(fixture.cases)('$mode transpose=$transpose skip=$skipBlanks op=$operation', (native) => {
		const { workbook, sheet, session } = setup();
		const payload = session.copy(0, range('A1:B2'));
		const before = structuredClone({
			comments: sheet.comments,
			dataValidations: sheet.dataValidations,
		});
		session.paste(0, range('D4:E5'), payload, native);
		for (const cell of native.cells) {
			expect(getCell(sheet, cell.row, cell.col)?.value ?? null).toBe(cell.value);
			const comment = sheet.comments.find(
				(c) => c.address.row === cell.row && c.address.col === cell.col,
			);
			expect(comment ? { author: comment.author, text: comment.text } : null).toEqual(cell.comment);
			const expected = cell.validation && {
				...cell.validation,
				formula1: cell.validation.formula1.replace(/^=/, ''),
			};
			expect(normalized(validationAt(workbook, 0, cell.row, cell.col))).toEqual(expected);
		}
		const after = structuredClone({
			comments: sheet.comments,
			dataValidations: sheet.dataValidations,
		});
		session.undo();
		expect({ comments: sheet.comments, dataValidations: sheet.dataValidations }).toEqual(before);
		session.redo();
		expect({ comments: sheet.comments, dataValidations: sheet.dataValidations }).toEqual(after);
	});
	it('retains note replies and validation through cross-workbook copy, tiling and save/reload', async () => {
		const { sheet, session } = setup();
		sheet.comments[0]!.replies = [
			{ author: 'Reviewer', text: 'reply', date: '2026-10-07T00:00:00Z' },
		];
		const payload = session.copy(0, range('A1:B2'));
		sheet.comments = [];
		const target = createWorkbook();
		const edits = createEditSession(target);
		edits.paste(0, range('D4:G7'), payload, { mode: 'all', transpose: true });
		expect(target.sheets[0]!.comments).toHaveLength(12);
		expect(target.sheets[0]!.comments[0]!.replies).toEqual([
			{ author: 'Reviewer', text: 'reply', date: '2026-10-07T00:00:00Z' },
		]);
		const reopened = await loadXlsx(await saveXlsx(target));
		expect(reopened.sheets[0]!.comments).toHaveLength(12);
		expect(
			reopened.sheets[0]!.comments.find((c) => c.address.row === 3 && c.address.col === 3)?.replies,
		).toEqual([{ author: 'Reviewer', text: 'reply', date: '2026-10-07T00:00:00Z' }]);
		expect(normalized(validationAt(reopened, 0, 6, 5))?.formula1).toBe('E7>0');
		edits.undo();
		expect(target.sheets[0]!.comments).toEqual([]);
		expect(target.sheets[0]!.dataValidations).toEqual([]);
	});
});
