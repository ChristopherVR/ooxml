/**
 * Tables made in code can skip the default table style and can have the
 * slots a merge covers filled in for them; column widths that cannot be
 * written are rejected. Each saved case is loaded back, so the written
 * `a:tbl` and the reloaded table agree.
 */
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import { PptxHandler } from '../../PptxHandler';
import type { TablePptxElement } from '../../types';
import { createTableElement } from './ElementFactory';
import type { TableInput, TableOptions } from './types';

async function roundTrip(
	input: TableInput,
	options?: TableOptions,
): Promise<{ xml: string; table: TablePptxElement | undefined }> {
	const { handler, createSlide } = await PptxHandler.create();
	const bytes = await handler.save([createSlide('Blank').addTable(input, options).build()]);
	const zip = await JSZip.loadAsync(bytes);
	const slide = (await zip.file('ppt/slides/slide1.xml')?.async('string')) ?? '';
	const loaded = await new PptxHandler().load(bytes);
	return {
		xml: /<a:tbl>[\s\S]*<\/a:tbl>/u.exec(slide)?.[0] ?? '',
		table: loaded.slides[0]?.elements.find((e) => e.type === 'table') as
			| TablePptxElement
			| undefined,
	};
}

const layout = (table: TablePptxElement | undefined) =>
	(table?.tableData?.rows ?? []).map((row) =>
		row.cells.map((c) => (c.hMerge && c.vMerge ? 'hv' : c.hMerge ? 'h' : c.vMerge ? 'v' : c.text)),
	);

describe('table style', () => {
	const rows = [{ cells: [{ text: 'a' }] }];

	it('keeps the default table style when style is omitted', async () => {
		const { xml, table } = await roundTrip({ rows });
		expect(xml).toContain(
			'<a:tableStyleId>{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}</a:tableStyleId>',
		);
		expect(table?.tableData?.tableStyleId).toBe('{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}');
	});

	it("writes no table style for style: 'none' or ''", async () => {
		for (const style of ['none', '']) {
			const { xml, table } = await roundTrip({ rows, style });
			expect(xml).not.toContain('tableStyleId');
			expect(table?.tableData?.tableStyleId).toBeUndefined();
		}
		expect(createTableElement({ rows, style: 'none' }).tableData.tableStyleId).toBeUndefined();
	});

	it('keeps a style id that is given', async () => {
		const id = '{073A0DAA-6AF3-43AB-8588-CEC1D06C72B9}';
		const { table } = await roundTrip({ rows, style: id });
		expect(table?.tableData?.tableStyleId).toBe(id);
	});
});

describe('fillMergedCells', () => {
	it('fills the slots a merge covers and reads them back', async () => {
		const input: TableInput = {
			fillMergedCells: true,
			rows: [
				{ cells: [{ text: 'A', gridSpan: 2, rowSpan: 2 }, { text: 'B' }] },
				{ cells: [{ text: 'C' }] },
				{ cells: [{ text: 'D' }, { text: 'E', gridSpan: 2 }] },
			],
		};
		const expected = [
			['A', 'h', 'B'],
			['v', 'hv', 'C'],
			['D', 'E', 'h'],
		];
		const el = createTableElement(input);
		expect(el.tableData.columnWidths).toHaveLength(3);
		expect(layout(el)).toEqual(expected);
		const { xml, table } = await roundTrip(input);
		const trs = xml.match(/<a:tr\b[\s\S]*?<\/a:tr>/gu) ?? [];
		expect(trs.map((r) => (r.match(/<a:tc\b/gu) ?? []).length)).toEqual([3, 3, 3]);
		expect(trs[0]).toContain('<a:tc gridSpan="2" rowSpan="2">');
		expect(trs[1]).toContain('<a:tc hMerge="1" vMerge="1">');
		expect(layout(table)).toEqual(expected);
	});

	it('pads a short row and takes the column count from columnWidths', async () => {
		const { table } = await roundTrip({
			fillMergedCells: true,
			columnWidths: [1, 1, 2],
			rows: [{ cells: [{ text: 'A' }] }, { cells: [{ text: 'B', gridSpan: 2 }] }],
		});
		expect(layout(table)).toEqual([
			['A', '', ''],
			['B', 'h', ''],
		]);
		expect(table?.tableData?.columnWidths).toEqual([0.25, 0.25, 0.5]);
	});

	it('takes the cells as given without fillMergedCells', () => {
		const el = createTableElement({
			rows: [{ cells: [{ text: 'A', gridSpan: 2 }, { text: '' }] }],
		});
		expect(el.tableData.rows[0]?.cells.map((c) => c.hMerge ?? false)).toEqual([false, false]);
	});

	it('rejects spans it cannot place', () => {
		const make = (rows: TableInput['rows'], columnWidths?: number[]) => () =>
			createTableElement({
				fillMergedCells: true,
				rows,
				...(columnWidths ? { columnWidths } : {}),
			});
		expect(make([{ cells: [{ text: 'A', gridSpan: 0 }] }])).toThrow(RangeError);
		expect(make([{ cells: [{ text: 'A', gridSpan: 1.5 }] }])).toThrow(RangeError);
		expect(make([{ cells: [{ text: 'A', rowSpan: 2 }] }])).toThrow(/past the table/u);
		expect(make([{ cells: [{ text: 'A', gridSpan: 3 }] }], [1, 1])).toThrow(/columnWidths/u);
	});
});

describe('column widths', () => {
	it('rejects widths that cannot be written', () => {
		const rows = [{ cells: [{ text: 'a' }, { text: 'b' }] }];
		for (const columnWidths of [[1], [1, Number.NaN], [1, -1], [0, 0]]) {
			expect(() => createTableElement({ rows, columnWidths })).toThrow(RangeError);
		}
	});

	it('accepts a zero-width column beside a wider one', () => {
		const el = createTableElement({
			rows: [{ cells: [{ text: 'a' }, { text: 'b' }] }],
			columnWidths: [0, 3],
		});
		expect(el.tableData.columnWidths).toEqual([0, 1]);
	});
});
