import JSZip from 'jszip';
import { readFileSync } from 'node:fs';
import path from 'node:path';
import { expect, it } from 'vitest';
import { parseChartSpace } from '../../../chart';
import { PptxHandler } from '../../core/PptxHandler';
import type { ChartPptxElement } from '../../index';

it('saves repeated structural edits with chart caches and a fresh embedded worksheet', async () => {
	const bytes = readFileSync(path.resolve(__dirname, '../fixtures/e2e/chart-data-fidelity.pptx'));
	const handler = new PptxHandler();
	const data = await handler.load(new Uint8Array(bytes).buffer);
	const slide = data.slides.find((item) =>
		item.elements.some(
			(element) =>
				element.type === 'chart' && (element as ChartPptxElement).chartData?.chartType === 'bar',
		),
	)!;
	const element = slide.elements.find(
		(item) => item.type === 'chart' && (item as ChartPptxElement).chartData?.chartType === 'bar',
	) as ChartPptxElement;
	const chart = element.chartData!;
	const oldCount = chart.series.length;
	chart.series.push({ name: 'Added series', values: chart.categories.map((_, i) => i + 41) });
	slide.isDirty = true;
	const saved = await handler.save(data.slides);
	const zip = await JSZip.loadAsync(saved);
	const space = parseChartSpace(await zip.file(chart.chartPartPath!)!.async('string')).chartSpace;
	const added = space.plotArea.groups.flatMap((group) => group.series).at(-1)!;
	const newSheet = added.values!.formula!.split('!')[0]!;
	const source = await JSZip.loadAsync(bytes);
	let found = false;
	for (const name of Object.keys(zip.files).filter((name) => name.endsWith('.xlsx'))) {
		const embedded = await JSZip.loadAsync(await zip.file(name)!.async('uint8array'));
		const workbook = await embedded.file('xl/workbook.xml')!.async('string');
		if (!workbook.includes(`name="${newSheet}"`)) continue;
		found = true;
		const previous = await JSZip.loadAsync(await source.file(name)!.async('uint8array'));
		for (const part of Object.keys(previous.files).filter((part) =>
			/^xl\/worksheets\/sheet.*\.xml$/.test(part),
		))
			expect(await embedded.file(part)!.async('string')).toBe(
				await previous.file(part)!.async('string'),
			);
	}
	expect(found).toBe(true);
	expect(chart.series.at(-1)!.idx).toBe(added.index);
	chart.series.splice(0, 1);
	slide.isDirty = true;
	const second = await handler.save(data.slides);
	const reloaded = await new PptxHandler().load(new Uint8Array(second).buffer);
	const loaded = reloaded.slides
		.flatMap((slide) => slide.elements)
		.find(
			(item) =>
				item.type === 'chart' &&
				(item as ChartPptxElement).chartData?.chartPartPath === chart.chartPartPath,
		) as ChartPptxElement;
	expect(loaded.chartData!.series).toHaveLength(oldCount);
	expect(loaded.chartData!.series.at(-1)!.name).toBe('Added series');
	expect(loaded.chartData!.series.at(-1)!.values[0]).toBe(41);
});
