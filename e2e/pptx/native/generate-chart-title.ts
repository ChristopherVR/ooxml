import { mkdir, writeFile } from 'node:fs/promises';
import { resolve } from 'node:path';

import JSZip from 'jszip';

import { replaceInSlides } from '../../../src/core/dist/pptx/editor/render/find-replace.mjs';
import { PptxHandler } from '../../../src/core/dist/pptx/index.mjs';
import { createChartTitleDeck } from '../../../src/core/pptx/__tests__/helpers/chart-title-deck';

const directory = process.argv[2];
if (!directory) throw new Error('Pass an artifact output directory');
const output = resolve(directory);
await mkdir(output, { recursive: true });
const source = await createChartTitleDeck(),
	linked = await createChartTitleDeck(true);
const handler = new PptxHandler();
try {
	const data = await handler.load(source);
	const replaced = replaceInSlides(data.slides, 'report', 'summary');
	if (replaced.replacements !== 1) throw new Error('Expected one rich chart-title replacement');
	const saved = await handler.save([...replaced.slides]);
	const loaded = await handler.load(
		saved.buffer.slice(saved.byteOffset, saved.byteOffset + saved.byteLength) as ArrayBuffer,
	);
	const resaved = await handler.save(loaded.slides);
	const originalZip = await JSZip.loadAsync(source);
	for (const bytes of [saved, resaved]) {
		const zip = await JSZip.loadAsync(bytes);
		if (
			!Buffer.from(await zip.file('ppt/embeddings/Chart_Data.xlsx')!.async('uint8array')).equals(
				Buffer.from(await originalZip.file('ppt/embeddings/Chart_Data.xlsx')!.async('uint8array')),
			)
		)
			throw new Error('Title edit changed the embedded workbook');
	}
	const linkedData = await handler.load(linked);
	let rejected = false;
	try {
		replaceInSlides(linkedData.slides, 'report', 'summary');
	} catch {
		rejected = true;
	}
	if (!rejected) throw new Error('Linked chart-title replacement must require source editing');
	for (const [name, bytes] of [
		['source-rich', source],
		['rich-replaced', saved],
		['rich-resaved', resaved],
		['source-linked', linked],
	] as const)
		await writeFile(resolve(output, `${name}.pptx`), new Uint8Array(bytes));
} finally {
	handler.dispose();
}
console.log(`Generated synthetic chart-title acceptance artifacts: ${output}`);
