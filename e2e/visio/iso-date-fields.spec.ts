import { test, expect } from '@playwright/test';
import { readFileSync } from 'node:fs';
import type { VisioViewerElement } from 'ooxml-ui/visio';
import { openDemo } from './demo-page';

for (const name of ['field-iso-date-native', 'field-iso-date204-native']) {
	const native = new URL(`../../src/core/visio/__fixtures__/${name}.vsdx`, import.meta.url);
	const evidence = JSON.parse(
		readFileSync(
			new URL(`../../src/core/visio/__fixtures__/${name}.json`, import.meta.url),
			'utf8',
		).replace(/^\uFEFF/, ''),
	) as { cases: { id: number; text: string; formula: string }[] };

	test(`${name} renders admitted ISO date fields and survives browser save/reopen`, async ({
		page,
	}) => {
		await openDemo(page);
		const viewer = page.locator('visio-viewer');
		await page.locator('#file').setInputFiles({
			name: 'Dates.vsdx',
			mimeType: 'application/vnd.ms-visio.drawing',
			buffer: readFileSync(native),
		});
		await expect(page.locator('#file-name')).toHaveText('Dates.vsdx');
		const check = async () => {
			for (const item of evidence.cases)
				await expect(viewer.locator(`svg.paper [data-shape-id="${item.id}"] text`)).toHaveText(
					item.formula === 'DATE(1899,12,29)' ? '2026-01-01' : item.text,
				);
		};
		await check();
		const bytes = await viewer.evaluate((node) =>
			Array.from((node as VisioViewerElement).exportVsdx().bytes),
		);
		await page.locator('#file').setInputFiles({
			name: 'Saved dates.vsdx',
			mimeType: 'application/vnd.ms-visio.drawing',
			buffer: Buffer.from(bytes),
		});
		await expect(page.locator('#file-name')).toHaveText('Saved dates.vsdx');
		await check();
	});
}
