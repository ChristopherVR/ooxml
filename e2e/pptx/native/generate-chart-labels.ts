import { mkdir, writeFile } from 'node:fs/promises';
import { resolve } from 'node:path';

import JSZip from 'jszip';

import { replaceInSlides } from '../../../src/core/dist/pptx/editor/render/find-replace.mjs';
import { PptxHandler } from '../../../src/core/dist/pptx/index.mjs';
import { createChartLabelDeck } from '../../../src/core/pptx/__tests__/helpers/chart-label-deck';

const output = resolve(
	process.argv[2] ??
		(() => {
			throw new Error('Pass an output directory');
		})(),
);
await mkdir(output, { recursive: true });
const source = await createChartLabelDeck();
const handler = new PptxHandler();
try {
	const data = await handler.load(source);
	const chart = data.slides[0]!.elements.find((element) => element.type === 'chart');
	if (chart?.type !== 'chart' || !chart.chartData?.dataLabelTextEditable)
		throw new Error('Expected editable sources');
	const options = { includeChartData: true, includeChartTitles: false };
	const named = replaceInSlides(data.slides, 'Dataset', 'Renamed dataset', options);
	const categorized = replaceInSlides(named.slides, 'Q1', 'First quarter', options);
	if (named.replacements !== 1 || categorized.replacements !== 1)
		throw new Error('Expected both label replacements');
	const saved = await handler.save([...categorized.slides]);
	const reloaded = await handler.load(
		saved.buffer.slice(saved.byteOffset, saved.byteOffset + saved.byteLength) as ArrayBuffer,
	);
	const resaved = await handler.save(reloaded.slides);
	for (const [name, bytes] of [
		['source-labels', source],
		['labels-replaced', saved],
		['labels-resaved', resaved],
	] as const) {
		const zip = await JSZip.loadAsync(bytes);
		const workbook = await JSZip.loadAsync(
			await zip.file('ppt/embeddings/Chart_Data.xlsx')!.async('uint8array'),
		);
		const sheet = await workbook.file('xl/worksheets/sheet1.xml')!.async('string');
		if (!sheet.includes('001 untouched ') || !sheet.includes('s="7"'))
			throw new Error('Unrelated cell preservation failed');
		await writeFile(resolve(output, `${name}.pptx`), new Uint8Array(bytes));
	}
} finally {
	handler.dispose();
}
console.log(`Generated synthetic chart-label transaction decks: ${output}`);
