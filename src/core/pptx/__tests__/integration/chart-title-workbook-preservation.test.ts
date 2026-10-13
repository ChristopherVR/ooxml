import JSZip from 'jszip';
import { expect, it, vi } from 'vitest';
import * as workbookSave from '../../core/core/runtime/chart-external-workbook-save';

import { PptxHandler } from '../../index';
import { replaceInSlides } from '../../editor/render/find-replace';
import { createChartTitleDeck } from '../helpers/chart-title-deck';

it('leaves shared rich workbook strings byte-identical when only a chart title changes', async () => {
	const zip = await JSZip.loadAsync(await createChartTitleDeck());
	const path = 'ppt/charts/chart1.xml';
	zip.file(
		path,
		(await zip.file(path)!.async('string')).replace(
			'<c:tx><c:v>Dataset</c:v></c:tx>',
			'<c:tx><c:strRef><c:f>Sheet1!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Dataset</c:v></c:pt></c:strCache></c:strRef></c:tx>',
		),
	);
	const embedding = 'ppt/embeddings/Chart_Data.xlsx';
	const workbook = await JSZip.loadAsync(await zip.file(embedding)!.async('uint8array'));
	workbook.file(
		'xl/worksheets/sheet1.xml',
		(await workbook.file('xl/worksheets/sheet1.xml')!.async('string')).replace(
			'<c r="B1" t="inlineStr"><is><t>Dataset</t></is></c>',
			'<c r="B1" t="s"><v>0</v></c>',
		),
	);
	workbook.file(
		'xl/sharedStrings.xml',
		'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1"><si><r><rPr><b/></rPr><t>Dataset</t></r></si></sst>',
	);
	workbook.file(
		'xl/_rels/workbook.xml.rels',
		(await workbook.file('xl/_rels/workbook.xml.rels')!.async('string')).replace(
			'</Relationships>',
			'<Relationship Id="rIdStrings" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>',
		),
	);
	workbook.file(
		'[Content_Types].xml',
		(await workbook.file('[Content_Types].xml')!.async('string')).replace(
			'</Types>',
			'<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>',
		),
	);
	const originalWorkbook = await workbook.generateAsync({ type: 'uint8array' });
	zip.file(embedding, originalWorkbook);
	const handler = new PptxHandler();
	const data = await handler.load(await zip.generateAsync({ type: 'arraybuffer' }));
	const replaced = replaceInSlides(data.slides, 'report', 'summary');
	expect(replaced.replacements).toBe(1);
	const saved = await JSZip.loadAsync(await handler.save([...replaced.slides]));
	expect(await saved.file(embedding)!.async('uint8array')).toEqual(originalWorkbook);
	expect(await saved.file(path)!.async('string')).toContain('sum');
	handler.dispose();
});

it('restores workbook source cells after save, data undo and a title edit on the same handler', async () => {
	const zip = await JSZip.loadAsync(await createChartTitleDeck());
	const path = 'ppt/charts/chart1.xml';
	zip.file(
		path,
		(await zip.file(path)!.async('string')).replace(
			'<c:tx><c:v>Dataset</c:v></c:tx>',
			'<c:tx><c:strRef><c:f>Sheet1!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Dataset</c:v></c:pt></c:strCache></c:strRef></c:tx>',
		),
	);
	const handler = new PptxHandler();
	const data = await handler.load(await zip.generateAsync({ type: 'arraybuffer' }));
	const original = structuredClone(data.slides);
	const chart = data.slides[0]!.elements[0]!;
	if (chart.type !== 'chart' || !chart.chartData) throw new Error('Expected chart');
	chart.chartData.series[0]!.name = 'Edited series';
	await handler.save(data.slides);
	const undoneAndRetitled = replaceInSlides(original, 'report', 'summary');
	const saved = await JSZip.loadAsync(await handler.save([...undoneAndRetitled.slides]));
	const workbook = await JSZip.loadAsync(
		await saved.file('ppt/embeddings/Chart_Data.xlsx')!.async('uint8array'),
	);
	const sheet = await workbook.file('xl/worksheets/sheet1.xml')!.async('string');
	expect(sheet).toContain('Dataset');
	expect(sheet).not.toContain('Edited series');
	expect(await saved.file(path)!.async('string')).toContain('Dataset');
	const reopened = await new PptxHandler().load(await saved.generateAsync({ type: 'arraybuffer' }));
	const result = reopened.slides[0]!.elements[0]!;
	if (result.type !== 'chart') throw new Error('Expected chart');
	expect(result.chartData!.titleRuns!.map((run) => run.text).join('')).toBe('Annual summary');
	expect(result.chartData!.series[0]!.name).toBe('Dataset');
	handler.dispose();
});

it('retries a warned workbook write on a later save instead of treating the changed cache as reconciled', async () => {
	const zip = await JSZip.loadAsync(await createChartTitleDeck());
	const path = 'ppt/charts/chart1.xml';
	zip.file(
		path,
		(await zip.file(path)!.async('string')).replace(
			'<c:tx><c:v>Dataset</c:v></c:tx>',
			'<c:tx><c:strRef><c:f>Sheet1!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Dataset</c:v></c:pt></c:strCache></c:strRef></c:tx>',
		),
	);
	const handler = new PptxHandler();
	const data = await handler.load(await zip.generateAsync({ type: 'arraybuffer' }));
	const chart = data.slides[0]!.elements[0]!;
	if (chart.type !== 'chart' || !chart.chartData) throw new Error('Expected chart');
	chart.chartData.series[0]!.name = 'Retried series';
	const actual = workbookSave.saveChartExternalWorkbookUpdates;
	const spy = vi
		.spyOn(workbookSave, 'saveChartExternalWorkbookUpdates')
		.mockResolvedValueOnce(false)
		.mockImplementation(actual);
	try {
		await handler.save(data.slides);
		const saved = await JSZip.loadAsync(await handler.save(data.slides));
		expect(spy).toHaveBeenCalledTimes(2);
		const workbook = await JSZip.loadAsync(
			await saved.file('ppt/embeddings/Chart_Data.xlsx')!.async('uint8array'),
		);
		expect(await workbook.file('xl/worksheets/sheet1.xml')!.async('string')).toContain(
			'Retried series',
		);
		await handler.save(data.slides);
		expect(spy).toHaveBeenCalledTimes(2);
	} finally {
		spy.mockRestore();
		handler.dispose();
	}
});
