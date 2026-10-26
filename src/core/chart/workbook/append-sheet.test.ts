import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { loadXlsx } from '../../xlsx/read';
import { getCell } from '../../xlsx/cells';
import { appendChartWorkbookSheet } from './append-sheet';
import { createChartWorkbook } from './create';
import { NS, children, first, parseXml } from '../../xml';
const source = () =>
	createChartWorkbook(
		{ categories: ['A', 'B'], series: [{ name: 'Original', values: [2, 3] }] },
		'ChartData1',
	);
describe('fresh chart workbook sheet allocation', () => {
	it('allocates a new worksheet, leaving original cells and unknown parts byte for byte', async () => {
		const input = await JSZip.loadAsync(await source());
		input.file('customXml/item1.xml', '<private preserve="true"/>');
		input.file('xl/media/image1.bin', new Uint8Array([0, 1, 2, 3]));
		const bytes = await input.generateAsync({ type: 'uint8array' });
		let allocated = '';
		const result = await appendChartWorkbookSheet(bytes, (name) => {
			allocated = name;
			return [
				['Label', 'New'],
				['C', 42],
				['D', null],
			];
		});
		expect(allocated).toBe('ChartData2');
		expect(result.sheetName).toBe(allocated);
		const zip = await JSZip.loadAsync(result.bytes);
		for (const name of Object.keys(input.files).filter(
			(n) =>
				!input.files[n]!.dir &&
				!['xl/workbook.xml', 'xl/_rels/workbook.xml.rels', '[Content_Types].xml'].includes(n),
		))
			expect(await zip.file(name)!.async('uint8array'), name).toEqual(
				await input.file(name)!.async('uint8array'),
			);
		const book = await loadXlsx(result.bytes);
		expect(book.sheets.map((sheet) => sheet.name)).toEqual(['ChartData1', 'ChartData2']);
		expect(getCell(book.sheets[1]!, 1, 1)?.value).toBe(42);
		expect(getCell(book.sheets[0]!, 1, 1)?.value).toBe(2);
		const workbook = parseXml(await zip.file('xl/workbook.xml')!.async('string')).documentElement;
		expect(
			first(first(workbook, 'bookViews', NS.x)!, 'workbookView', NS.x)?.getAttribute('activeTab'),
		).toBe('1');
	});
	it('selects each newly allocated data sheet in every existing workbook view', async () => {
		const zip = await JSZip.loadAsync(await source());
		const path = 'xl/workbook.xml';
		zip.file(
			path,
			(await zip.file(path)!.async('string'))
				.replace(/<bookViews>[\s\S]*?<\/bookViews>/u, '')
				.replace(
					'<sheets>',
					'<bookViews><workbookView activeTab="0" xWindow="42"/><workbookView activeTab="0"/></bookViews><sheets>',
				),
		);
		const firstAppend = await appendChartWorkbookSheet(
			await zip.generateAsync({ type: 'uint8array' }),
			() => [[1]],
		);
		const secondAppend = await appendChartWorkbookSheet(firstAppend.bytes, () => [[2]]);
		const result = await JSZip.loadAsync(secondAppend.bytes);
		const workbook = parseXml(await result.file(path)!.async('string')).documentElement;
		const views = children(first(workbook, 'bookViews', NS.x)!, 'workbookView', NS.x);
		expect(views.map((view) => view.getAttribute('activeTab'))).toEqual(['2', '2']);
		expect(views[0]!.getAttribute('xWindow')).toBe('42');
	});
	it('refuses protected structure and invalid numeric data without modifying the source', async () => {
		const bytes = await source();
		const zip = await JSZip.loadAsync(bytes);
		const part = 'xl/workbook.xml';
		zip.file(
			part,
			(await zip.file(part)!.async('string')).replace(
				'<sheets>',
				'<workbookProtection lockStructure="1"/><sheets>',
			),
		);
		await expect(
			appendChartWorkbookSheet(await zip.generateAsync({ type: 'uint8array' }), () => [[1]]),
		).rejects.toThrow('protected');
		await expect(appendChartWorkbookSheet(bytes, () => [[NaN]])).rejects.toThrow('finite');
		expect((await loadXlsx(bytes)).sheets).toHaveLength(1);
	});
});
