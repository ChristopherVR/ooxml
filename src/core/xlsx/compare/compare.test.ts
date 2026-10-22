import { describe, expect, it } from 'vitest';
import { formatAddress, formatRange, parseAddress } from '../address';
import { putCell } from '../cells';
import type { SpilledCell } from '../formula/spill';
import { cellError, type Cell, type CellStyle, type Workbook } from '../model';
import { internStyle, patchStyle, type StylePatch } from '../styles';
import { createWorkbook, defaultCellStyle } from '../workbook';
import { compareWorkbooks } from './compare';

function put(workbook: Workbook, ref: string, cell: Partial<Cell>, sheet = 0): void {
	const at = parseAddress(ref);
	if (!at) throw new Error(ref);
	putCell(workbook.sheets[sheet]!, at.row, at.col, { value: null, ...cell });
}

const styled = (workbook: Workbook, patch: StylePatch): number =>
	internStyle(workbook, patchStyle(defaultCellStyle(), patch));

/** `kind A1` for every cell difference, in result order. */
function cells(before: Workbook, after: Workbook, options = {}): string[] {
	return compareWorkbooks(before, after, options).sheets.flatMap((sheet) =>
		sheet.cells.map((d) => `${d.kind} ${formatAddress(d.address)}`),
	);
}

describe('compareWorkbooks cells', () => {
	it('finds nothing between a workbook and a copy of itself', () => {
		const book = createWorkbook({ sheets: ['Data', 'Notes'] });
		put(book, 'A1', { value: 1, styleId: styled(book, { font: { bold: true } }) });
		put(book, 'B1', { value: 2, formula: 'A1*2' });
		put(book, 'C3', { value: 'text' }, 1);
		const result = compareWorkbooks(book, structuredClone(book));
		expect(result.count).toBe(0);
		expect(result.truncated).toBe(false);
		expect(result.workbook).toEqual([]);
		expect(result.sheets.map((s) => [s.name, s.matchedBy, s.cells.length])).toEqual([
			['Data', 'name', 0],
			['Notes', 'name', 0],
		]);
	});

	it('sorts value changes into entered values, formulas and calculated values', () => {
		const before = createWorkbook();
		put(before, 'A1', { value: 1 });
		put(before, 'B1', { value: 2, formula: 'A1*2' });
		put(before, 'C1', { value: 2, formula: 'A1+1' });
		put(before, 'E1', { value: 'gone' });
		put(before, 'F1', { value: '1' });
		put(before, 'G1', { value: 4 });
		put(before, 'H1', { value: cellError('#N/A') });
		put(before, 'I1', { value: 3, formula: 'SUM(A1:A3)' });
		const after = createWorkbook();
		put(after, 'A1', { value: 5 });
		put(after, 'B1', { value: 15, formula: 'A1*3' });
		put(after, 'C1', { value: 6, formula: 'A1+1' });
		put(after, 'D1', { value: 'new' });
		put(after, 'F1', { value: 1 });
		put(after, 'G1', { value: 4, formula: '2*2' });
		put(after, 'H1', { value: cellError('#N/A') });
		put(after, 'I1', { value: 3 });
		const result = compareWorkbooks(before, after);
		expect(cells(before, after)).toEqual([
			'enteredValue A1',
			'formula B1',
			'calculatedValue C1',
			'enteredValue D1',
			'enteredValue E1',
			'enteredValue F1',
			'formula G1',
			'formula I1',
		]);
		const [, b1, c1, d1, e1] = result.sheets[0]!.cells;
		expect(b1).toMatchObject({
			before: { value: 2, formula: 'A1*2' },
			after: { value: 15, formula: 'A1*3' },
		});
		expect(c1).toMatchObject({ before: { value: 2 }, after: { value: 6 } });
		expect(d1).toMatchObject({ before: { value: null }, after: { value: 'new' } });
		expect(e1).toMatchObject({ before: { value: 'gone' }, after: { value: null } });
	});

	it('counts spilled, array formula and data table results as calculated values', () => {
		const build = (first: number, second: number) => {
			const book = createWorkbook();
			const range = { start: { row: 0, col: 0 }, end: { row: 1, col: 0 } };
			put(book, 'A1', { value: first, formula: 'ROW(A1:A2)', arrayRange: range });
			put(book, 'A2', { value: second });
			const spilled: SpilledCell = { value: second, spillAnchor: { row: 0, col: 1 } };
			put(book, 'B1', { value: first, formula: 'SEQUENCE(2)' });
			put(book, 'B2', spilled);
			book.sheets[0]!.dataTables = [
				{ ref: { start: { row: 4, col: 0 }, end: { row: 4, col: 0 } }, r1: 'C1' },
			];
			put(book, 'A5', { value: second * 10 });
			return book;
		};
		expect(cells(build(1, 2), build(1, 3))).toEqual([
			'calculatedValue A2',
			'calculatedValue B2',
			'calculatedValue A5',
		]);
		// Typing over a spill result is an entered value.
		const typed = build(1, 2);
		put(typed, 'B2', { value: 9 });
		expect(cells(build(1, 2), typed)).toEqual(['enteredValue B2']);
		// A data table result shows the table's TABLE() formula, so a new input cell changes it.
		const moved = build(1, 2);
		moved.sheets[0]!.dataTables![0]!.r1 = 'D1';
		const [table] = compareWorkbooks(build(1, 2), moved).sheets[0]!.cells;
		expect(table).toMatchObject({
			kind: 'formula',
			before: { formula: 'TABLE(,C1)' },
			after: { formula: 'TABLE(,D1)' },
		});
	});

	it('tells array results from typed cells beside many array formulas', () => {
		const build = (second: number, below: number) => {
			const book = createWorkbook();
			for (let col = 0; col < 300; col++) {
				const range = { start: { row: 0, col }, end: { row: 1, col } };
				putCell(book.sheets[0]!, 0, col, { value: 1, formula: 'ROW(A1:A2)', arrayRange: range });
				putCell(book.sheets[0]!, 1, col, { value: second });
				putCell(book.sheets[0]!, 2, col, { value: below });
			}
			return book;
		};
		const kinds = cells(build(2, 0), build(3, 1)).map((d) => d.split(' ')[0]);
		expect(kinds.filter((k) => k === 'calculatedValue')).toHaveLength(300);
		expect(kinds.filter((k) => k === 'enteredValue')).toHaveLength(300);
		expect(kinds.slice(0, 1).concat(kinds.slice(-1))).toEqual(['calculatedValue', 'enteredValue']);
	});

	it('compares formats by their attributes, not by style index', () => {
		const before = createWorkbook();
		const bold = styled(before, { font: { bold: true } });
		const red = styled(before, {
			fill: { type: 'pattern', pattern: 'solid', fgColor: { rgb: 'FF0000' } },
		});
		put(before, 'A1', { value: 1, styleId: bold });
		put(before, 'B1', { value: 2, styleId: red });
		put(before, 'C1', { value: 3, styleId: bold });
		put(before, 'D1', { value: 4, styleId: bold });
		const after = createWorkbook();
		// The same formats, numbered the other way round and spelled differently.
		const base = defaultCellStyle();
		const redAgain: CellStyle = {
			...base,
			font: { ...base.font, bold: false },
			fill: {
				type: 'pattern',
				pattern: 'solid',
				fgColor: { rgb: 'ffff0000' },
				bgColor: { indexed: 64 },
			},
			alignment: { horizontal: 'general', vertical: 'bottom' },
			protection: { locked: true },
		};
		after.styles.push(redAgain, { ...base, font: { ...base.font, family: 3, bold: true } });
		put(after, 'A1', { value: 1, styleId: 2 });
		put(after, 'B1', { value: 2, styleId: 1 });
		put(after, 'C1', { value: 3, styleId: styled(after, { font: { italic: true } }) });
		put(after, 'D1', {
			value: 4,
			styleId: styled(after, { font: { bold: true }, numFmt: '0.00' }),
		});
		const result = compareWorkbooks(before, after);
		const diffs = result.sheets[0]!.cells;
		expect(diffs.map((d) => [formatAddress(d.address), d.kind, d.attributes])).toEqual([
			['C1', 'format', ['font']],
			['D1', 'format', ['numberFormat']],
		]);
		expect(diffs[1]!.before.style?.numFmt).toBe('General');
		expect(diffs[1]!.after.style?.numFmt).toBe('0.00');
	});

	it('gives a cell missing on one side the format of its row, then of its column', () => {
		const before = createWorkbook();
		const bold = styled(before, { font: { bold: true } });
		before.sheets[0]!.rowInfo.set(1, { styleId: bold });
		before.sheets[0]!.columns.push({ min: 3, max: 4, styleId: bold });
		const after = createWorkbook();
		const boldAfter = styled(after, { font: { bold: true } });
		put(after, 'B2', { styleId: boldAfter });
		put(after, 'D1', { styleId: boldAfter });
		put(after, 'F1', { styleId: boldAfter });
		expect(cells(before, after)).toEqual(['format F1']);
		expect(cells(after, before)).toEqual(['format F1']);
	});

	it('reports a change to the runs of a rich-text string as a format difference', () => {
		const build = (bold: boolean) => {
			const book = createWorkbook();
			put(book, 'A1', {
				value: 'Hello world',
				richText: [{ text: 'Hello ', font: { bold } }, { text: 'world' }],
			});
			return book;
		};
		expect(compareWorkbooks(build(false), build(false)).count).toBe(0);
		const [difference] = compareWorkbooks(build(false), build(true)).sheets[0]!.cells;
		expect(difference).toMatchObject({ kind: 'format', attributes: ['richText'] });
		// Runs split differently, and one run without formatting, are the same text.
		const split = createWorkbook();
		put(split, 'A1', {
			value: 'Hello world',
			richText: [{ text: 'Hel', font: { bold: false } }, { text: 'lo ' }, { text: 'world' }],
		});
		const plain = createWorkbook();
		put(plain, 'A1', { value: 'Hello world', richText: [{ text: 'Hello world' }] });
		expect(compareWorkbooks(build(false), split).count).toBe(0);
		expect(compareWorkbooks(split, plain).count).toBe(0);
		// A new font shows even when the text was edited as well.
		const word = (text: string, font: { bold?: boolean; italic?: boolean }) => {
			const book = createWorkbook();
			put(book, 'A1', { value: text, richText: [{ text, font }, { text: '!' }] });
			return book;
		};
		const edited = compareWorkbooks(word('one', { bold: true }), word('two', { italic: true }));
		expect(edited.sheets[0]!.cells.map((d) => [d.kind, d.attributes])).toEqual([
			['enteredValue', undefined],
			['format', ['richText']],
		]);
		expect(compareWorkbooks(word('one', { bold: true }), word('two', { bold: true })).count).toBe(
			1,
		);
		// One run in the cell's own font is the same as plain text.
		const boldCell = (runs: boolean) => {
			const book = createWorkbook();
			const styleId = styled(book, { font: { bold: true } });
			const font = { ...defaultCellStyle().font, bold: true };
			const richText = runs ? { richText: [{ text: 'abc', font }] } : {};
			put(book, 'A1', { value: 'abc', styleId, ...richText });
			return book;
		};
		expect(compareWorkbooks(boldCell(false), boldCell(true)).count).toBe(0);
		// A run that switches the cell's bold off is formatted, as the grid paints it.
		const unbold = boldCell(false);
		putCell(unbold.sheets[0]!, 0, 0, {
			...unbold.sheets[0]!.rows.get(0)!.get(0)!,
			richText: [{ text: 'abc', font: { bold: false } }],
		});
		expect(compareWorkbooks(boldCell(false), unbold).sheets[0]!.cells[0]?.attributes).toEqual([
			'richText',
		]);
		// A new cell font under unchanged runs is a font change, not a rich-text one.
		const runs = (size: number) => {
			const book = createWorkbook();
			const styleId = styled(book, { font: { size } });
			put(book, 'A1', {
				value: 'ab',
				styleId,
				richText: [{ text: 'a', font: { bold: true } }, { text: 'b' }],
			});
			return book;
		};
		expect(compareWorkbooks(runs(11), runs(14)).sheets[0]!.cells[0]?.attributes).toEqual(['font']);
		// A bold run in a cell that turns bold still paints bold: a font change only.
		const boldRun = (bold: boolean) => {
			const book = createWorkbook();
			const styleId = styled(book, { font: { bold } });
			put(book, 'A1', {
				value: 'ab',
				styleId,
				richText: [{ text: 'a', font: { bold: true } }, { text: 'b' }],
			});
			return book;
		};
		expect(compareWorkbooks(boldRun(false), boldRun(true)).sheets[0]!.cells[0]?.attributes).toEqual(
			['font'],
		);
		// A run made redundant by the new cell font, next to an unchanged run: still a font change.
		const pair = (bold: boolean) => {
			const book = createWorkbook();
			const styleId = styled(book, { font: { bold } });
			const first = bold ? { text: 'a' } : { text: 'a', font: { bold: true } };
			const second = { text: 'b', font: { italic: true } };
			put(book, 'A1', { value: 'ab', styleId, richText: [first, second] });
			return book;
		};
		const [redundant] = compareWorkbooks(pair(false), pair(true)).sheets[0]!.cells;
		expect(redundant?.attributes).toEqual(['font']);
		// Clearing a formatted string is a value change only.
		const cleared = compareWorkbooks(build(true), createWorkbook()).sheets[0]!.cells;
		expect(cleared.map((d) => d.kind)).toEqual(['enteredValue']);
	});

	it('reports a formula switched to or from implicit intersection where that matters', () => {
		const build = (legacy: boolean) => {
			const book = createWorkbook();
			put(book, 'B1', { value: 1, formula: 'A2:A3', ...(legacy ? { legacyFormula: true } : {}) });
			put(book, 'C1', { value: 2, formula: 'A1+1', ...(legacy ? { legacyFormula: true } : {}) });
			return book;
		};
		const [range, ...rest] = compareWorkbooks(build(false), build(true)).sheets[0]!.cells;
		expect(rest).toEqual([]);
		expect(range).toMatchObject({
			kind: 'formula',
			address: { row: 0, col: 1 },
			before: { formula: 'A2:A3' },
			after: { formula: 'A2:A3', legacyFormula: true },
		});
	});

	it('resolves a missing font name or size from the default format, as the grid does', () => {
		const build = (size: number) => {
			const book = createWorkbook();
			book.styles[0]!.font.size = size;
			book.styles.push({ ...defaultCellStyle(), font: { bold: true } });
			put(book, 'A1', { value: 1, styleId: 1 });
			return book;
		};
		expect(cells(build(11), build(11))).toEqual([]);
		const [difference] = compareWorkbooks(build(11), build(24)).sheets[0]!.cells;
		expect(difference).toMatchObject({ kind: 'format', attributes: ['font'] });
		// The formats reported are the ones compared, with the inherited size filled in.
		expect([difference?.before.style?.font.size, difference?.after.style?.font.size]).toEqual([
			11, 24,
		]);
	});

	it('treats a border edge without a line as no edge', () => {
		const build = (border: object) => {
			const book = createWorkbook();
			book.styles.push({ ...defaultCellStyle(), border: border as CellStyle['border'] });
			put(book, 'A1', { value: 1, styleId: 1 });
			return book;
		};
		expect(compareWorkbooks(build({}), build({ left: { style: 'none' } })).count).toBe(0);
	});

	it('compares only the chosen categories and says what it left out', () => {
		const before = createWorkbook();
		put(before, 'A1', { value: 1 });
		put(before, 'B1', { value: 1, formula: 'A1' });
		const after = createWorkbook();
		put(after, 'A1', { value: 2, styleId: styled(after, { font: { bold: true } }) });
		put(after, 'B1', { value: 2, formula: 'A1*1' });
		const result = compareWorkbooks(before, after, { categories: ['formula', 'format'] });
		expect(result.categories).toEqual(['formula', 'format']);
		expect(result.sheets[0]!.cells.map((d) => d.kind)).toEqual(['format', 'formula']);
		expect(result.notCompared.slice(0, 3)).toEqual(['enteredValue', 'calculatedValue', 'sheet']);
		expect(result.notCompared).toContain('comments');
	});

	it('stops at maxDifferences in row-major order and flags the result as truncated', () => {
		const before = createWorkbook();
		const after = createWorkbook();
		// Inserted out of order: the cap still keeps the first cells in row-major order.
		for (const ref of ['C3', 'A1', 'B2', 'A2', 'C1']) put(after, ref, { value: ref });
		const capped = compareWorkbooks(before, after, { maxDifferences: 3 });
		expect(capped.truncated).toBe(true);
		expect(capped.count).toBe(3);
		expect(capped.sheets[0]!.cells.map((d) => formatAddress(d.address))).toEqual([
			'A1',
			'C1',
			'A2',
		]);
		const exact = compareWorkbooks(before, after, { maxDifferences: 5 });
		expect([exact.truncated, exact.count]).toEqual([false, 5]);
	});

	it('does not read the sheets after the one where the cap cut a difference', () => {
		const before = createWorkbook({ sheets: ['First', 'Second'] });
		const after = createWorkbook({ sheets: ['First', 'Second'] });
		put(after, 'A1', { value: 1 });
		put(after, 'A2', { value: 2 });
		let reads = 0;
		class CountingRows extends Map<number, Map<number, Cell>> {
			override get(row: number) {
				reads++;
				return super.get(row);
			}
		}
		after.sheets[1]!.rows = new CountingRows([[0, new Map([[0, { value: 3 }]])]]);
		const result = compareWorkbooks(before, after, { maxDifferences: 1 });
		expect([result.count, result.truncated, reads]).toEqual([1, true, 0]);
		expect(result.sheets.map((s) => [s.name, s.cells.length])).toEqual([
			['First', 1],
			['Second', 0],
		]);
	});

	it('groups neighbouring differences of one kind into rectangles', () => {
		const before = createWorkbook();
		const after = createWorkbook();
		for (const ref of ['B2', 'C2', 'B3', 'C3', 'B4', 'E2', 'E3']) put(after, ref, { value: 1 });
		const italic = styled(after, { font: { italic: true } });
		const filled = styled(after, {
			fill: { type: 'pattern', pattern: 'solid', fgColor: { rgb: 'FFFF00' } },
		});
		put(after, 'A6', { styleId: italic });
		put(after, 'B6', { styleId: italic });
		put(after, 'C6', { styleId: filled });
		const { ranges } = compareWorkbooks(before, after).sheets[0]!;
		expect(ranges.map((r) => [formatRange(r.range), r.kind, r.count, r.attributes])).toEqual([
			['B2:C3', 'enteredValue', 4, undefined],
			['E2:E3', 'enteredValue', 2, undefined],
			['B4', 'enteredValue', 1, undefined],
			['A6:B6', 'format', 2, ['font']],
			['C6', 'format', 1, ['fill']],
		]);
	});
});
