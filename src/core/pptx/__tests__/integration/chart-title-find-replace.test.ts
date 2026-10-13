import JSZip from 'jszip';
import { expect, it } from 'vitest';
import { PptxHandler } from '../../index';
import {
	findInSlides,
	findReplacementAvailability,
	replaceInSlides,
} from '../../editor/render/find-replace';
import { collectElementText, searchSlides } from '../../editor/render/slide-search';
import { createChartTitleDeck } from '../helpers/chart-title-deck';

it('finds across chart-title runs and saves text-only edits while preserving XML properties and workbook bytes', async () => {
	const source = await createChartTitleDeck();
	const handler = new PptxHandler(),
		data = await handler.load(source);
	const matches = findInSlides(data.slides, 'Annual report');
	expect(matches).toHaveLength(1);
	expect(matches[0]).toMatchObject({ chartTitle: true, readOnly: false });
	expect(collectElementText(data.slides[0]!.elements[0]!)).toBe('Annual report');
	expect(searchSlides(data.slides, 'Annual report')[0]?.matchCount).toBe(1);
	const replaced = replaceInSlides(data.slides, 'report', 'summary');
	expect(replaced.replacements).toBe(1);
	const saved = await JSZip.loadAsync(await handler.save([...replaced.slides]));
	const xml = await saved.file('ppt/charts/chart1.xml')!.async('string');
	expect(xml).toContain('Annual summary');
	for (const marker of [
		'lang="fr-FR"',
		'lang="en-GB"',
		'u="sng"',
		'kern="1200"',
		'accent2',
		'Arial',
		'rIdTitleLink',
		'chart-title-parity',
		'xml:space="preserve"',
	])
		expect(xml).toContain(marker);
	const original = await JSZip.loadAsync(source);
	expect(await saved.file('ppt/embeddings/Chart_Data.xlsx')!.async('uint8array')).toEqual(
		await original.file('ppt/embeddings/Chart_Data.xlsx')!.async('uint8array'),
	);
	const reopened = await new PptxHandler().load(await saved.generateAsync({ type: 'arraybuffer' }));
	expect(findInSlides(reopened.slides, 'Annual summary')).toHaveLength(1);
	expect(reopened.slides[0]!.elements[0]!.type).toBe('chart');
	handler.dispose();
});

it('keeps the remaining title runs when deleting the entire first run', async () => {
	const handler = new PptxHandler(),
		data = await handler.load(await createChartTitleDeck());
	const replaced = replaceInSlides(data.slides, 'Annual re', '');
	const saved = await handler.save([...replaced.slides]);
	const reopened = await new PptxHandler().load(
		saved.buffer.slice(saved.byteOffset, saved.byteOffset + saved.byteLength) as ArrayBuffer,
	);
	expect(findInSlides(reopened.slides, 'port')).toHaveLength(1);
	handler.dispose();
});

it('does not let a saved text-only edit mask a later run-formatting change', async () => {
	const handler = new PptxHandler(),
		data = await handler.load(await createChartTitleDeck());
	const replaced = replaceInSlides(data.slides, 'report', 'summary');
	await handler.save([...replaced.slides]);
	const chart = replaced.slides[0]!.elements[0]!;
	if (chart.type !== 'chart') throw new Error('Expected a chart');
	chart.chartData!.titleRuns = chart.chartData!.titleRuns!.map((run) => ({ ...run, bold: false }));
	const saved = await handler.save([...replaced.slides]);
	const reopened = await new PptxHandler().load(
		saved.buffer.slice(saved.byteOffset, saved.byteOffset + saved.byteLength) as ArrayBuffer,
	);
	const reloaded = reopened.slides[0]!.elements[0]!;
	if (reloaded.type !== 'chart') throw new Error('Expected a chart');
	expect(reloaded.chartData!.titleRuns![0]!.bold).toBe(false);
	handler.dispose();
});

it('finds linked chart titles but rejects replacements atomically instead of changing their caches', async () => {
	const handler = new PptxHandler(),
		data = await handler.load(await createChartTitleDeck(true));
	const matches = findInSlides(data.slides, 'Annual report');
	expect(matches).toHaveLength(1);
	expect(matches[0]?.readOnly).toBe(true);
	expect(findReplacementAvailability(matches, 0)).toEqual({
		current: false,
		all: false,
		hasReadOnly: true,
	});
	expect(() => replaceInSlides(data.slides, 'report', 'summary')).toThrow('source workbook');
	expect(findInSlides(data.slides, 'Annual report')).toHaveLength(1);
	handler.dispose();
});

it('excludes chart titles to replace ordinary text alongside a linked title without editing its workbook or cache', async () => {
	const source = await createChartTitleDeck(true),
		handler = new PptxHandler(),
		data = await handler.load(source);
	expect(findInSlides(data.slides, 'report')).toHaveLength(2);
	expect(() => replaceInSlides(data.slides, 'report', 'summary')).toThrow('source workbook');
	const replaced = replaceInSlides(data.slides, 'report', 'summary', { includeChartTitles: false });
	expect(replaced.replacements).toBe(1);
	expect(findInSlides(replaced.slides, 'Annual report')).toHaveLength(1);
	expect(findInSlides(replaced.slides, 'Slide-only summary')).toHaveLength(1);
	const saved = await JSZip.loadAsync(await handler.save([...replaced.slides])),
		original = await JSZip.loadAsync(source);
	expect(await saved.file('ppt/embeddings/Chart_Data.xlsx')!.async('uint8array')).toEqual(
		await original.file('ppt/embeddings/Chart_Data.xlsx')!.async('uint8array'),
	);
	expect(await saved.file('ppt/charts/chart1.xml')!.async('string')).toContain('Sheet1!$A$1');
	handler.dispose();
});

it.each(['break', 'field', 'paragraph'])(
	'reports a %s in a complex rich chart title as read-only',
	async (kind) => {
		const zip = await JSZip.loadAsync(await createChartTitleDeck());
		const path = 'ppt/charts/chart1.xml';
		let xml = await zip.file(path)!.async('string');
		if (kind === 'break') xml = xml.replace('</a:r><a:r>', '</a:r><a:br/><a:r>');
		if (kind === 'field')
			xml = xml.replace(
				'</a:r><a:r>',
				'</a:r><a:fld id="{00000000-0000-0000-0000-000000000001}" type="slidenum"><a:t>1</a:t></a:fld><a:r>',
			);
		if (kind === 'paragraph') xml = xml.replace('</a:r><a:r>', '</a:r></a:p><a:p><a:r>');
		zip.file(path, xml);
		const handler = new PptxHandler(),
			data = await handler.load(await zip.generateAsync({ type: 'arraybuffer' }));
		expect(findInSlides(data.slides, 'Annual')[0]?.readOnly).toBe(true);
		expect(() => replaceInSlides(data.slides, 'Annual', 'Updated')).toThrow('source workbook');
		handler.dispose();
	},
);
