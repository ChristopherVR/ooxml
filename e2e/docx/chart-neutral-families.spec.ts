import { readFileSync } from 'node:fs';
import path from 'node:path';
import { expect, test } from '@playwright/test';
import JSZip from 'jszip';
import { reopenDocument } from './helpers';

const source = readFileSync(
	path.join(import.meta.dirname, '../../src/core/docx/__fixtures__/chart-families.docx'),
);
const points = (values: (string | number)[]) =>
	`<cx:lvl ptCount="${values.length}">${values.map((v, i) => `<cx:pt idx="${i}">${v}</cx:pt>`).join('')}</cx:lvl>`;
const chartXml = (kind: string) =>
	`<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><cx:chartData><cx:data id="0"><cx:strDim type="cat">${points(['A', 'A', 'A', 'B', 'B', 'B'])}</cx:strDim><cx:numDim type="val">${points([1, 2, 3, 4, 5, 6])}</cx:numDim></cx:data></cx:chartData><cx:chart><cx:title><cx:tx><cx:txData><cx:v>Shared ${kind}</cx:v></cx:txData></cx:tx></cx:title><cx:plotArea><cx:plotAreaRegion><cx:series layoutId="${kind}"><cx:dataId val="0"/></cx:series></cx:plotAreaRegion></cx:plotArea></cx:chart></cx:chartSpace>`;

for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid'])
	for (const kind of ['funnel', 'boxWhisker'])
		test(`${framework}: shared ${kind} paints in Word and Print Layout without changing its source`, async ({
			page,
		}, testInfo) => {
			const zip = await JSZip.loadAsync(source);
			const part = 'word/charts/chartEx1.xml';
			const xml = chartXml(kind);
			zip.file(part, xml);
			const bytes = await zip.generateAsync({ type: 'uint8array' });
			await page.setViewportSize({ width: 1500, height: 1000 });
			await page.goto(`/?framework=${framework}`);
			await reopenDocument(page, 'neutral-chart.docx', bytes);
			const editor = page.locator('docx-editor');
			const chart = editor
				.locator('.ProseMirror [data-docx-chart="1"]')
				.filter({ has: page.locator('svg text').filter({ hasText: `Shared ${kind}` }) });
			await expect(chart).toHaveCount(1);
			const marks = kind === 'boxWhisker' ? '[data-chart-part="box"]' : 'g[data-chart-point]';
			await expect(chart.locator(marks)).toHaveCount(kind === 'boxWhisker' ? 2 : 6);
			await chart.scrollIntoViewIfNeeded();
			await testInfo.attach(`${kind}-browser.png`, {
				body: await chart.screenshot(),
				contentType: 'image/png',
			});
			await editor.getByRole('tab', { name: 'View', exact: true }).click();
			await editor.getByRole('button', { name: 'Print Layout', exact: true }).click();
			const printed = editor
				.locator('.dve-print-pages [data-docx-chart="1"]')
				.filter({ has: page.locator('svg text').filter({ hasText: `Shared ${kind}` }) });
			await expect(printed.locator(marks)).toHaveCount(kind === 'boxWhisker' ? 2 : 6);
			const saved = Uint8Array.from(
				await editor.evaluate(async (element) =>
					Array.from(
						await (element as unknown as { saveBytes(): Promise<Uint8Array> }).saveBytes(),
					),
				),
			);
			const reopened = await JSZip.loadAsync(saved);
			expect(await reopened.file(part)!.async('string')).toBe(xml);
		});
