/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec */
import { writeFile } from 'node:fs/promises';

import { expect, test } from '@playwright/test';
import { XMLParser } from 'fast-xml-parser';
import JSZip from 'jszip';
import { PptxHandler } from 'pptx-viewer-core';
import type { ChartPptxElement } from 'pptx-viewer-core';

import { savePptxViaBackstage } from './save-pptx';
import { chooseCommand, openMenuOn } from './support/context-menu';
import { loadDeck, thumbnail } from './support/deck';
import { downloadBytes } from './support/exports';
import { validatePptxIntegrity } from './support/pptx-integrity';

test('duplicate a native chart and notes, then save with independent package parts', async ({
	page,
}, info) => {
	const { handler, data, createSlide } = await PptxHandler.createBlank();
	data.slides.push(
		createSlide('Blank')
			.addText('Original')
			.setNotes('Duplicate notes')
			.addChart(
				'line',
				{
					categories: ['A', 'B', 'C'],
					series: [{ name: 'Series A', values: [0, 2, 3] }],
				},
				{ x: 60, y: 80, width: 700, height: 400 },
			)
			.build(),
	);
	const sourcePath = info.outputPath('source.pptx');
	await writeFile(sourcePath, await handler.save(data.slides));
	await loadDeck(page, sourcePath);
	await openMenuOn(page, thumbnail(page, 1));
	await chooseCommand(page, 'duplicate');
	await expect(page.locator('[aria-label^="Go to slide"]')).toHaveCount(2);
	const bytes = await downloadBytes(await savePptxViaBackstage(page));
	await writeFile(info.outputPath('duplicated.pptx'), bytes);
	expect(await validatePptxIntegrity(bytes)).toEqual([]);
	const loaded = await new PptxHandler().load(
		bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer,
	);
	expect(loaded.slides).toHaveLength(2);
	expect(loaded.slides.map((slide) => slide.notes)).toEqual(['Duplicate notes', 'Duplicate notes']);
	const charts = loaded.slides.map(
		(slide) => slide.elements.find((element) => element.type === 'chart') as ChartPptxElement,
	);
	expect(charts[1].chartData!.chartPartPath).not.toBe(charts[0].chartData!.chartPartPath);
	const zip = await JSZip.loadAsync(bytes);
	expect(
		Object.keys(zip.files).filter((path) => /^ppt\/notesSlides\/notesSlide\d+\.xml$/u.test(path)),
	).toHaveLength(2);
	const parser = new XMLParser({ ignoreAttributes: false });
	for (const slide of loaded.slides) {
		const fileName = slide.id.slice(slide.id.lastIndexOf('/') + 1);
		const rels = parser.parse(await zip.file(`ppt/slides/_rels/${fileName}.rels`)!.async('string'))
			.Relationships.Relationship;
		const ids = new Set(
			(Array.isArray(rels) ? rels : [rels]).map((rel: Record<string, string>) => rel['@_Id']),
		);
		const xml = await zip.file(slide.id)!.async('string');
		for (const match of xml.matchAll(/r:(?:id|embed|link)="([^"]+)"/gu))
			expect(ids.has(match[1])).toBe(true);
	}
});
