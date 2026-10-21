import { expect, it } from 'vitest';
import { createEditSession, createWorkbook, parseRange } from '../index';
import { selectionFormatting } from './selection-format';

it('reports mixed fields separately while retaining shared formatting', () => {
	const book = createWorkbook();
	const session = createEditSession(book);
	session.applyStyle(0, [parseRange('A1')!], { font: { bold: true } });
	const result = selectionFormatting(book, 0, [parseRange('A1:B1')!]);
	expect(result.bold).toBeNull();
	expect(result.italic).toBe(false);
	expect(result.formattingComplete).toBe(true);
});
it('bounds work for whole-sheet selections and reports unknown fields honestly', () => {
	const book = createWorkbook();
	const result = selectionFormatting(book, 0, [parseRange('A1:XFD1048576')!]);
	expect(result.formattingComplete).toBe(false);
	expect(result.fontFamily).toBeNull();
});
