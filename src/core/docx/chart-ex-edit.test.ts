import { readFileSync } from 'node:fs';
import path from 'node:path';
import JSZip from 'jszip';
import { loadXlsx } from '../xlsx/read';
import { describe, expect, it } from 'vitest';
import {
	applyChartExSpaceDataSheetEdits,
	setChartExLegend,
	setChartExTitle,
	chartExSpaceDataSheet,
	createChartWorkbook,
} from '../chart';
import { chartsIn, documentDrawingBlocks, loadDocx, saveDocx } from './index';
import { editDocxChartEx, isDocxChartExEditable } from './chart-ex-edit';
const fixture = () =>
	readFileSync(path.join(import.meta.dirname, '__fixtures__/chart-families.docx'));

async function embeddedFixture() {
	const zip = await JSZip.loadAsync(fixture());
	zip.file(
		'[Content_Types].xml',
		(await zip.file('[Content_Types].xml')!.async('string')).replace(
			/<Types([^>]*)>/,
			'<Types$1><Default Extension="xlsx" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/>',
		),
	);
	const chart = 'word/charts/chartEx1.xml';
	zip.file(
		chart,
		(await zip.file(chart)!.async('string')).replaceAll('$E', '$A').replaceAll('$F', '$B'),
	);
	const rel = 'word/charts/_rels/chartEx1.xml.rels';
	zip.file(
		rel,
		(await zip.file(rel)!.async('string'))
			.replace('relationships/oleObject', 'relationships/package')
			.replace('Target="Book1" TargetMode="External"', 'Target="../embeddings/chart.xlsx"'),
	);
	const parsed = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
	const model = chartsIn(documentDrawingBlocks(parsed.model)).find(
		(c) => c.layoutId === 'waterfall',
	)!;
	zip.file(
		'word/embeddings/chart.xlsx',
		await createChartWorkbook(chartExSpaceDataSheet(model.chartExSpace!)),
	);
	return zip.generateAsync({ type: 'uint8array' });
}

describe('Word ChartEx edits', () => {
	it('saves title and legend edits through neutral writer, preserving all other package parts', async () => {
		const bytes = fixture();
		const loaded = await loadDocx(bytes);
		const chart = chartsIn(documentDrawingBlocks(loaded.model)).find(
			(c) => c.layoutId === 'waterfall',
		)!;
		expect(isDocxChartExEditable(chart)).toBe(true);
		Object.assign(
			chart,
			editDocxChartEx(chart, (space) =>
				setChartExLegend(setChartExTitle(space, 'New title'), { visible: true, position: 'b' }),
			),
		);
		const output = await saveDocx(loaded.model);
		const a = await JSZip.loadAsync(bytes),
			b = await JSZip.loadAsync(output);
		for (const name of Object.keys(a.files).filter((n) => !a.files[n]!.dir && n !== chart.partName))
			expect(await b.file(name)!.async('uint8array'), name).toEqual(
				await a.file(name)!.async('uint8array'),
			);
		const back = chartsIn(documentDrawingBlocks((await loadDocx(output)).model)).find(
			(c) => c.layoutId === 'waterfall',
		)!;
		expect(back.title).toBe('New title');
		expect(back.chartExSpace?.legend?.position).toBe('b');
	});
	it('keeps value edits and the embedded workbook synchronized', async () => {
		const loaded = await loadDocx(await embeddedFixture());
		const chart = chartsIn(documentDrawingBlocks(loaded.model)).find(
			(c) => c.layoutId === 'waterfall',
		)!;
		expect(chart.workbookPartName).toBeTruthy();
		Object.assign(
			chart,
			editDocxChartEx(chart, (s) =>
				applyChartExSpaceDataSheetEdits(s, [{ kind: 'value', series: 0, category: 1, value: 42 }]),
			),
		);
		const output = await saveDocx(loaded.model);
		const reloaded = await loadDocx(output);
		const back = chartsIn(documentDrawingBlocks(reloaded.model)).find(
			(c) => c.layoutId === 'waterfall',
		)!;
		expect(chartExSpaceDataSheet(back.chartExSpace!).series[0]!.values[1]).toBe(42);
		const zip = await JSZip.loadAsync(output);
		const workbook = await JSZip.loadAsync(
			await zip.file(back.workbookPartName!)!.async('uint8array'),
		);
		expect(await workbook.file('xl/worksheets/sheet1.xml')!.async('string')).toMatch(
			/r="B3"[^>]*><v>42<\/v>/,
		);
	});
	it('refuses external-source data edits atomically while title-only edits remain available', async () => {
		const loaded = await loadDocx(fixture());
		const chart = chartsIn(documentDrawingBlocks(loaded.model)).find(
			(c) => c.layoutId === 'sunburst',
		)!;
		expect(() =>
			editDocxChartEx(chart, (s) =>
				applyChartExSpaceDataSheetEdits(s, [{ kind: 'value', series: 0, category: 0, value: 42 }]),
			),
		).toThrow(/local|external/);
		Object.assign(
			chart,
			editDocxChartEx(chart, (s) => setChartExTitle(s, 'Allowed title')),
		);
		await expect(saveDocx(loaded.model)).resolves.toBeInstanceOf(Uint8Array);
	});
});

it('grows an embedded chart on a fresh worksheet while preserving its original source cells', async () => {
	const bytes = await embeddedFixture();
	const loaded = await loadDocx(bytes);
	const chart = chartsIn(documentDrawingBlocks(loaded.model)).find(
		(c) => c.layoutId === 'waterfall',
	)!;
	const originalZip = await JSZip.loadAsync(bytes);
	const originalBook = await JSZip.loadAsync(
		await originalZip.file(chart.workbookPartName!)!.async('uint8array'),
	);
	for (const edits of [
		[{ kind: 'add-category' }],
		[{ kind: 'add-category' }, { kind: 'value', series: 0, category: 7, value: 91 }],
	] as const)
		Object.assign(
			chart,
			editDocxChartEx(chart, (space) => applyChartExSpaceDataSheetEdits(space, edits)),
		);
	expect(chart.chartExRebaseData).toBe(true);
	const output = await saveDocx(loaded.model);
	const back = chartsIn(documentDrawingBlocks((await loadDocx(output)).model)).find(
		(c) => c.layoutId === 'waterfall',
	)!;
	const data = chartExSpaceDataSheet(back.chartExSpace!);
	expect(data.categories).toHaveLength(8);
	expect(data.series[0]!.values[7]).toBe(91);
	const zip = await JSZip.loadAsync(output);
	const book = await JSZip.loadAsync(await zip.file(back.workbookPartName!)!.async('uint8array'));
	expect(
		(await loadXlsx(await book.generateAsync({ type: 'uint8array' }))).sheets[0]!.rows,
	).toEqual(
		(await loadXlsx(await originalBook.generateAsync({ type: 'uint8array' }))).sheets[0]!.rows,
	);
	expect(await book.file('xl/workbook.xml')!.async('string')).toContain('ChartExData1');
	expect(await book.file('xl/worksheets/sheet2.xml')!.async('string')).toContain('<v>91</v>');
});
