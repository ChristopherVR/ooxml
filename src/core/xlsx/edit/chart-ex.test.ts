import { readFileSync } from 'node:fs';
import path from 'node:path';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { applyChartExSpaceDataSheetEdits, setChartExLegend, setChartExTitle } from '../../chart';
import { getCell } from '../cells';
import { loadXlsx } from '../read';
import { saveXlsx } from '../write';
import { createEditSession } from './session';
import { xlsxChartExDataSheet } from './chart-ex';
const fixture = () =>
	readFileSync(path.join(import.meta.dirname, '../__fixtures__/excel-chart-families.xlsx'));
function find(book: Awaited<ReturnType<typeof loadXlsx>>, layout: string) {
	for (const [s, ws] of book.sheets.entries())
		for (const [i, d] of ws.drawings.entries())
			if (d.kind === 'unsupported' && d.chartExSpace?.plotArea.series[0]?.layoutId === layout)
				return { s, i, d };
	throw Error('Chart not found.');
}

describe('Excel ChartEx operations', () => {
	it('hydrates hidden-name sources and edits values with one undo/redo transaction', async () => {
		const book = await loadXlsx(fixture());
		const { s, i } = find(book, 'waterfall');
		const initial = xlsxChartExDataSheet(book, s, i);
		expect(initial.categories).toHaveLength(6);
		const session = createEditSession(book);
		session.editChartEx(s, i, (space) =>
			applyChartExSpaceDataSheetEdits(space, [
				{ kind: 'value', series: 0, category: 1, value: 42 },
			]),
		);
		expect(xlsxChartExDataSheet(book, s, i).series[0]!.values[1]).toBe(42);
		const data = book.sheets.find((ws) => ws.name === 'Data')!;
		expect(getCell(data, 2, 5)?.value).toBe(42);
		session.undo();
		expect(xlsxChartExDataSheet(book, s, i)).toEqual(initial);
		session.redo();
		expect(xlsxChartExDataSheet(book, s, i).series[0]!.values[1]).toBe(42);
		const output = await saveXlsx(book);
		const back = await loadXlsx(output);
		const target = find(back, 'waterfall');
		expect(xlsxChartExDataSheet(back, target.s, target.i).series[0]!.values[1]).toBe(42);
	});
	it('saves title and legend to the kept ChartEx part, retaining its drawing frame', async () => {
		const bytes = fixture();
		const book = await loadXlsx(bytes);
		const { s, i, d } = find(book, 'treemap');
		const source = d.sourceXml;
		const session = createEditSession(book);
		session.editChartEx(s, i, (space) =>
			setChartExLegend(setChartExTitle(space, 'Edited hierarchy'), {
				visible: true,
				position: 'b',
			}),
		);
		const output = await saveXlsx(book);
		const back = await loadXlsx(output);
		const target = find(back, 'treemap');
		expect(target.d.sourceXml).toBe(source);
		expect(target.d.chartExSpace?.legend?.position).toBe('b');
		expect(target.d.chartExSpace?.title?.tx?.rich?.text).toBe('Edited hierarchy');
		const zip = await JSZip.loadAsync(output);
		expect(await zip.file(d.chartExPartName!)!.async('string')).toContain('Edited hierarchy');
	});
	it('preserves hierarchy parents and refuses formula-dependent source cells before history changes', async () => {
		const book = await loadXlsx(fixture());
		const { s, i } = find(book, 'treemap');
		const before = xlsxChartExDataSheet(book, s, i);
		const session = createEditSession(book);
		session.editChartEx(s, i, (space) =>
			applyChartExSpaceDataSheetEdits(space, [
				{ kind: 'category-label', category: 1, label: 'New leaf' },
			]),
		);
		expect(xlsxChartExDataSheet(book, s, i).categoryLevels.slice(1)).toEqual(
			before.categoryLevels.slice(1),
		);
		const data = book.sheets.find((ws) => ws.name === 'Data')!;
		const source = getCell(data, 1, 21)!;
		source.formula = '1+1';
		expect(() =>
			session.editChartEx(s, i, (space) =>
				applyChartExSpaceDataSheetEdits(space, [
					{ kind: 'value', series: 0, category: 0, value: 99 },
				]),
			),
		).toThrow('Formula-dependent');
		expect(source.formula).toBe('1+1');
		expect(source.value).not.toBe(99);
	});
});
