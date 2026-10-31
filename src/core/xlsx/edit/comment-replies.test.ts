import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { loadXlsx } from '../read/index';
import { saveXlsx } from '../write/index';
import { createEditSession } from './session';

const at = { row: 1, col: 1 };
const load = () =>
	loadXlsx(
		new Uint8Array(
			readFileSync(new URL('../__fixtures__/excel-threaded-comments.xlsx', import.meta.url)),
		),
	);

describe('threaded comment reply edits', () => {
	it('edits a native reply, preserves authors and dates, and round-trips undo and redo', async () => {
		const workbook = await load();
		const original = structuredClone(workbook.sheets[0]!.comments);
		expect(original[0]!.date).toBeTruthy();
		expect(original[0]!.replies).toHaveLength(2);
		const session = createEditSession(workbook);
		session.setCommentReply(0, at, 1, 'Edited second reply < & >');
		const expected = structuredClone(original);
		expected[0]!.replies![1]!.text = 'Edited second reply < & >';
		expect(workbook.sheets[0]!.comments).toEqual(expected);
		expect((await loadXlsx(await saveXlsx(workbook))).sheets[0]!.comments).toEqual(expected);
		session.undo();
		expect(workbook.sheets[0]!.comments).toEqual(original);
		expect((await loadXlsx(await saveXlsx(workbook))).sheets[0]!.comments).toEqual(original);
		session.redo();
		expect((await loadXlsx(await saveXlsx(workbook))).sheets[0]!.comments).toEqual(expected);
	});

	it.each([-1, 2, 0.5, NaN])(
		'rejects an invalid reply index %s without recording an edit',
		async (index) => {
			const workbook = await load();
			const before = structuredClone(workbook.sheets[0]!.comments);
			const session = createEditSession(workbook);
			expect(() => session.setCommentReply(0, at, index, 'No')).toThrow('Comment reply not found.');
			expect(workbook.sheets[0]!.comments).toEqual(before);
			expect(session.canUndo()).toBe(false);
		},
	);

	it('ignores unchanged text and rejects cells without a thread', async () => {
		const workbook = await load();
		const session = createEditSession(workbook);
		session.setCommentReply(0, at, 0, 'First reply');
		expect(session.canUndo()).toBe(false);
		expect(() => session.setCommentReply(0, { row: 0, col: 0 }, 0, 'No')).toThrow(RangeError);
		expect(session.canUndo()).toBe(false);
	});
});
