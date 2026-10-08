import type { TablePptxElement } from '../../index';
import { describe, expect, it } from 'vitest';

import { setCellText, withCellStyle, withCellText } from './table-cell-edit';

function makeTable(): TablePptxElement {
	return {
		id: 't1',
		type: 'table',
		x: 0,
		y: 0,
		width: 100,
		height: 100,
		tableData: {
			columnWidths: [0.5, 0.5],
			rows: [{ cells: [{ text: 'a' }, { text: 'b' }] }, { cells: [{ text: 'c' }, { text: 'd' }] }],
		},
	} as unknown as TablePptxElement;
}

describe('setCellText', () => {
	it('replaces the targeted cell text and leaves siblings untouched', () => {
		const el = makeTable();
		const next = setCellText(el, 0, 1, 'B!');
		expect(next.tableData?.rows[0].cells[1].text).toBe('B!');
		expect(next.tableData?.rows[0].cells[0].text).toBe('a');
		expect(next.tableData?.rows[1].cells[1].text).toBe('d');
	});

	it('does not mutate the source element', () => {
		const el = makeTable();
		const next = setCellText(el, 1, 0, 'C!');
		expect(el.tableData?.rows[1].cells[0].text).toBe('c');
		expect(next).not.toBe(el);
		expect(next.tableData).not.toBe(el.tableData);
		// Unchanged rows are reused by reference.
		expect(next.tableData?.rows[0]).toBe(el.tableData?.rows[0]);
	});

	it('returns the element unchanged when it carries no tableData', () => {
		const el = {
			id: 't',
			type: 'table',
			x: 0,
			y: 0,
			width: 1,
			height: 1,
		} as unknown as TablePptxElement;
		expect(setCellText(el, 0, 0, 'x')).toBe(el);
	});

	// Regression: every binding's table renderer paints `cell.textRuns` when it
	// is present and only falls back to `cell.text` when it is not, so a cell
	// re-texted with its old run model still attached keeps painting the OLD
	// wording. `desktop-manipulation` and `mobile-table` caught this in four of
	// the five bindings; Vanilla passed only because its editor spelled out
	// `textRuns: undefined` locally instead of the shared helper doing it.
	it('drops the stale per-run model of the text it replaced', () => {
		const el = makeTable();
		el.tableData!.rows[0].cells[1].textRuns = [{ text: 'b', bold: true }];
		const next = setCellText(el, 0, 1, 'B!');
		expect(next.tableData?.rows[0].cells[1].text).toBe('B!');
		expect(next.tableData?.rows[0].cells[1].textRuns).toBeUndefined();
		expect('textRuns' in next.tableData!.rows[0].cells[1]).toBeFalsy();
		// And the source cell keeps its runs: the update is immutable.
		expect(el.tableData?.rows[0].cells[1].textRuns).toHaveLength(1);
	});

	it('leaves an untouched cell its per-run model', () => {
		const el = makeTable();
		el.tableData!.rows[1].cells[0].textRuns = [{ text: 'c', italic: true }];
		const next = setCellText(el, 0, 1, 'B!');
		expect(next.tableData?.rows[1].cells[0].textRuns).toHaveLength(1);
	});
});

describe('withCellText', () => {
	it('re-texts a cell and removes the run model describing the old text', () => {
		const cell = { text: 'old', textRuns: [{ text: 'old' }], style: { bold: true } };
		const next = withCellText(cell, 'new');
		expect(next.text).toBe('new');
		expect('textRuns' in next).toBeFalsy();
		// Everything else about the cell survives.
		expect(next.style).toStrictEqual({ bold: true });
		// The source is untouched.
		expect(cell.textRuns).toHaveLength(1);
	});

	it("keeps only the first paragraph's layout, as the rewritten markup does", () => {
		const next = withCellText(
			{
				text: 'one\ntwo',
				textRuns: [{ text: 'one' }, { text: '', isParagraphBreak: true }, { text: 'two' }],
				paragraphs: [{ align: 'right' }, { align: 'center' }],
			},
			'new',
		);
		expect(next).toStrictEqual({ text: 'new', paragraphs: [{ align: 'right' }] });
	});

	it('is a no-op on the run model for a cell that never had one', () => {
		expect(withCellText({ text: 'a' }, 'b')).toStrictEqual({ text: 'b' });
	});
});

describe('withCellStyle', () => {
	it('merges the patch into the cell style', () => {
		const cell = { text: 'a', style: { bold: true } };
		const next = withCellStyle(cell, { italic: true });
		expect(next).toStrictEqual({ text: 'a', style: { bold: true, italic: true } });
		expect(cell.style).toStrictEqual({ bold: true });
	});

	it('aligns every paragraph with the cell and keeps their other layout', () => {
		const cell = {
			text: 'one\ntwo\nthree',
			style: { align: 'right' as const },
			paragraphs: [{ align: 'right' as const }, { align: 'center' as const, rtl: true }, {}],
		};
		expect(withCellStyle(cell, { align: 'left' })).toStrictEqual({
			text: 'one\ntwo\nthree',
			style: { align: 'left' },
			paragraphs: [{ align: 'left' }, { align: 'left', rtl: true }, { align: 'left' }],
		});
		expect(cell.paragraphs[1]).toStrictEqual({ align: 'center', rtl: true });
	});

	it('gives each line its alignment when the cell had no paragraph layout', () => {
		expect(withCellStyle({ text: 'one\ntwo' }, { align: 'center' }).paragraphs).toStrictEqual([
			{ align: 'center' },
			{ align: 'center' },
		]);
	});

	it('gives every line of an edited cell its alignment', () => {
		const edited = withCellText({ text: 'a', paragraphs: [{ paragraphIndent: 4 }] }, 'one\ntwo');
		expect(withCellStyle(edited, { align: 'right' }).paragraphs).toStrictEqual([
			{ paragraphIndent: 4, align: 'right' },
			{ align: 'right' },
		]);
	});
});
