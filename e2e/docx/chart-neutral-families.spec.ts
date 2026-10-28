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
// Published Bing PCA test vector, not provider boundaries.
const mapXml = `<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"><cx:chartData><cx:data id="0"><cx:strDim type="cat"><cx:lvl ptCount="1"><cx:pt idx="0">Sample</cx:pt></cx:lvl></cx:strDim><cx:numDim type="colorVal"><cx:lvl ptCount="1"><cx:pt idx="0">42</cx:pt></cx:lvl></cx:numDim></cx:data></cx:chartData><cx:chart><cx:title><cx:tx><cx:txData><cx:v>Shared regionMap</cx:v></cx:txData></cx:tx></cx:title><cx:plotArea><cx:plotAreaRegion><cx:series layoutId="regionMap"><cx:valueColors><cx:minColor><a:srgbClr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" val="FF0000"/></cx:minColor><cx:midColor><a:srgbClr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" val="00FF00"/></cx:midColor><cx:maxColor><a:srgbClr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" val="0000FF"/></cx:maxColor></cx:valueColors><cx:valueColorPositions count="3"><cx:min><cx:extremeValue/></cx:min><cx:mid><cx:percent val="25"/></cx:mid><cx:max><cx:extremeValue/></cx:max></cx:valueColorPositions><cx:dataId val="0"/><cx:layoutPr><cx:geography projectionType="robinson"><cx:geoCache><cx:clear><cx:geoData entityId="sample" entityName="Sample"><cx:geoPolygons><cx:geoPolygon numPoints="4" pcaRings="1,vx1vilihnM6hR7mEl2Q"/></cx:geoPolygons></cx:geoData></cx:clear></cx:geoCache></cx:geography></cx:layoutPr></cx:series></cx:plotAreaRegion></cx:plotArea></cx:chart></cx:chartSpace>`;
const chartXml = (kind: string) =>
	kind === 'regionMap'
		? mapXml
		: `<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><cx:chartData><cx:data id="0"><cx:strDim type="cat">${points(['A', 'A', 'A', 'B', 'B', 'B'])}</cx:strDim><cx:numDim type="val">${points([1, 2, 3, 4, 5, 6])}</cx:numDim></cx:data></cx:chartData><cx:chart><cx:title><cx:tx><cx:txData><cx:v>Shared ${kind}</cx:v></cx:txData></cx:tx></cx:title><cx:plotArea><cx:plotAreaRegion><cx:series layoutId="${kind}"><cx:dataId val="0"/></cx:series></cx:plotAreaRegion></cx:plotArea></cx:chart></cx:chartSpace>`;

for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid'])
	for (const kind of ['funnel', 'boxWhisker', 'regionMap'])
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
			const marks =
				kind === 'regionMap'
					? '[data-map-geometry="cached"] path[data-chart-point]'
					: kind === 'boxWhisker'
						? '[data-chart-part="box"]'
						: 'g[data-chart-point]';
			await expect(chart.locator(marks)).toHaveCount(
				kind === 'regionMap' ? 1 : kind === 'boxWhisker' ? 2 : 6,
			);
			if (kind === 'regionMap')
				await expect(chart.locator(marks)).toHaveAttribute('fill', '#00aa55');
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
			await expect(printed.locator(marks)).toHaveCount(
				kind === 'regionMap' ? 1 : kind === 'boxWhisker' ? 2 : 6,
			);
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
