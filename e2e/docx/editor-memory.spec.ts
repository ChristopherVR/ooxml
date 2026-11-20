import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { expect, test } from '@playwright/test';
import { pictureGridPng } from '../../src/core/diagram/test-support/picture-grid';
import { heapReport, openAndCloseEditor } from '../../scripts/heap-snapshot.mjs';
import { demoReady, newDocument } from './helpers';

const fixture = (name: string) =>
	readFileSync(new URL(`../../src/core/docx/__fixtures__/${name}`, import.meta.url));

/** A picture tiled with mirroring: each tile is a group clipped with `clip-path: url(#id)`. */
async function tiledPicture(): Promise<Buffer> {
	const zip = await JSZip.loadAsync(fixture('review-inline/picture-insert-accepted.docx'));
	for (const name of Object.keys(zip.files))
		if (name.startsWith('word/media/') && name.endsWith('.png')) zip.file(name, pictureGridPng());
	const xml = await zip.file('word/document.xml')!.async('string');
	zip.file(
		'word/document.xml',
		xml.replace(
			/<a:stretch>[\s\S]*?<\/a:stretch>/,
			'<a:tile sx="200000" sy="200000" tx="95250" algn="ctr" flip="xy"/>',
		),
	);
	return Buffer.from(await zip.generateAsync({ type: 'uint8array' }));
}

// Pictures, charts and SmartArt paint with `url(#id)` clip paths, patterns and filters. Chromium's
// style cache kept a closed editor alive through them until some later stylesheet change.
const CYCLES = 2;
const PROBE = /^<docx-editor[^>]*data-leak-probe/;

test('editors opened and closed repeatedly are collected', async ({ page }) => {
	test.setTimeout(240_000);
	await page.goto('/?framework=vanilla');
	await demoReady(page);
	await newDocument(page);
	const session = await page.context().newCDPSession(page);
	const start = await heapReport(session, PROBE);
	const files: [string, Buffer][] = [
		['tiled-picture.docx', await tiledPicture()],
		['chart-families.docx', fixture('chart-families.docx')],
	];
	let references = 0;
	for (const [fileName, file] of files) {
		const bytes = [...file];
		for (let cycle = 1; cycle <= CYCLES; cycle++) {
			references += await page.evaluate(openAndCloseEditor, {
				tag: 'docx-editor',
				bytes,
				fileName,
			});
			const report = await heapReport(session, PROBE);
			test.info().annotations.push({
				type: 'heap',
				description: `${fileName} close ${cycle}: ${report.count} closed editors alive, ${report.detachedBytes} detached bytes, heap ${report.heapBytes - start.heapBytes} bytes over the start`,
			});
			expect.soft(report.count, `${fileName}, close ${cycle}`).toBe(0);
		}
	}
	expect(references).toBeGreaterThan(0);
});
