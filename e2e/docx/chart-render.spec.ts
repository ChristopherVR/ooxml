import { expect, test, type Locator, type Page } from '@playwright/test';
import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import JSZip from 'jszip';
import { fileInput } from './helpers';

// Built by ooxml-core (src/core/scripts/docx/build-chart-fixture.mjs) from the Excel chart of
// src/core/xlsx/__fixtures__/excel-features.xlsx: one inline clustered column chart.
const FIXTURE = fileURLToPath(new URL('./support/chart.docx', import.meta.url));
const CHART_PARTS = [
	'word/charts/chart1.xml',
	'word/charts/_rels/chart1.xml.rels',
	'word/charts/style1.xml',
	'word/charts/colors1.xml',
];

async function openFixture(page: Page, framework: string) {
	await page.goto(`/?framework=${framework}`);
	await (
		await fileInput(page)
	).setInputFiles({
		name: 'chart.docx',
		mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
		buffer: await readFile(FIXTURE),
	});
	const editor = page.locator('docx-editor');
	await expect(editor.locator('.ProseMirror')).toContainText('End of document.');
	return editor;
}

/** Series marks and the size of the SVG painted inside `chart`. */
const summarize = (chart: Locator) =>
	chart.evaluate((element) => {
		const svg = element.querySelector('svg')!;
		const box = svg.getBoundingClientRect();
		return {
			role: svg.getAttribute('role'),
			series0: svg.querySelectorAll('g[data-chart-series="0"]').length,
			series1: svg.querySelectorAll('g[data-chart-series="1"]').length,
			fills: [
				...new Set(
					[...svg.querySelectorAll('g[data-chart-series] > :first-child')].map((mark) =>
						mark.getAttribute('fill'),
					),
				),
			],
			texts: [...svg.querySelectorAll('text')].map((text) => text.textContent),
			width: box.width,
			height: box.height,
		};
	});

for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid']) {
	test(`${framework}: a Word chart is drawn as SVG in the editor and in Print Layout`, async ({
		page,
	}) => {
		await page.setViewportSize({ width: 1400, height: 900 });
		const editor = await openFixture(page, framework);

		const chart = editor.locator('.ProseMirror [data-docx-chart="1"]');
		await expect(chart).toHaveCount(1);
		const drawn = await summarize(chart);
		expect(drawn.role).toBe('img');
		expect(drawn.series0).toBe(4);
		expect(drawn.series1).toBe(4);
		// Two series in two theme accents, never the unresolved grey of a placeholder.
		expect(drawn.fills).toHaveLength(2);
		expect(drawn.texts).toEqual(expect.arrayContaining(['Sales by region', 'North', 'Sales']));
		expect(drawn.width).toBeGreaterThan(400);
		expect(drawn.height).toBeGreaterThan(200);
		await expect(editor.locator('.ProseMirror .dve-image-placeholder')).toHaveCount(0);

		await editor.getByRole('tab', { name: 'View', exact: true }).click();
		await editor.getByRole('button', { name: 'Print Layout', exact: true }).click();
		const printed = editor.locator('.dve-print-pages [data-docx-chart="1"]');
		await expect(printed).toHaveCount(1);
		const page1 = await summarize(printed);
		expect(page1.series0).toBe(4);
		expect(page1.series1).toBe(4);
		expect(page1.fills).toEqual(drawn.fills);

		// The chart stays non-editable and its parts are saved unchanged.
		const bytes = await editor.evaluate(async (element) =>
			Array.from(await (element as unknown as { saveBytes(): Promise<Uint8Array> }).saveBytes()),
		);
		const [saved, original] = await Promise.all([
			JSZip.loadAsync(Uint8Array.from(bytes)),
			JSZip.loadAsync(await readFile(FIXTURE)),
		]);
		for (const part of CHART_PARTS)
			expect(await saved.file(part)!.async('string')).toBe(
				await original.file(part)!.async('string'),
			);
	});
}
