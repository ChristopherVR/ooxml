import { readFileSync } from 'node:fs';
import path from 'node:path';
import { expect, test } from '@playwright/test';
import JSZip from 'jszip';
import { loadXlsx, saveXlsx } from 'ooxml-core/xlsx';
import { FRAMEWORKS, editor, openLanding } from './helpers';

const points = (values: (number | string)[]) =>
	`<cx:lvl ptCount="${values.length}">${values.map((v, i) => `<cx:pt idx="${i}">${v}</cx:pt>`).join('')}</cx:lvl>`;
// Published Bing PCA test vector, not provider boundaries.
const mapXml = `<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"><cx:chartData><cx:data id="0"><cx:strDim type="cat"><cx:lvl ptCount="1"><cx:pt idx="0">Sample</cx:pt></cx:lvl></cx:strDim><cx:numDim type="colorVal"><cx:lvl ptCount="1"><cx:pt idx="0">42</cx:pt></cx:lvl></cx:numDim></cx:data></cx:chartData><cx:chart><cx:title><cx:tx><cx:txData><cx:v>Shared regionMap</cx:v></cx:txData></cx:tx></cx:title><cx:plotArea><cx:plotAreaRegion><cx:series layoutId="regionMap"><cx:valueColors><cx:minColor><a:srgbClr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" val="FF0000"/></cx:minColor><cx:midColor><a:srgbClr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" val="00FF00"/></cx:midColor><cx:maxColor><a:srgbClr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" val="0000FF"/></cx:maxColor></cx:valueColors><cx:valueColorPositions count="3"><cx:min><cx:extremeValue/></cx:min><cx:mid><cx:percent val="25"/></cx:mid><cx:max><cx:extremeValue/></cx:max></cx:valueColorPositions><cx:dataId val="0"/><cx:layoutPr><cx:geography projectionType="robinson"><cx:geoCache><cx:clear><cx:geoData entityId="sample" entityName="Sample"><cx:geoPolygons><cx:geoPolygon numPoints="4" pcaRings="1,vx1vilihnM6hR7mEl2Q"/></cx:geoPolygons></cx:geoData></cx:clear></cx:geoCache></cx:geography></cx:layoutPr></cx:series></cx:plotAreaRegion></cx:plotArea></cx:chart></cx:chartSpace>`;
const chartXml = (kind: string) =>
	kind === 'regionMap'
		? mapXml
		: `<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"><cx:chartData><cx:data id="0"><cx:strDim type="cat">${points(['A', 'A', 'A', 'B', 'B', 'B'])}</cx:strDim><cx:numDim type="val">${points([1, 2, 3, 4, 5, 6])}</cx:numDim></cx:data></cx:chartData><cx:chart><cx:plotArea><cx:plotAreaRegion><cx:series layoutId="${kind}"><cx:dataId val="0"/></cx:series></cx:plotAreaRegion></cx:plotArea></cx:chart></cx:chartSpace>`;
async function fixture(kind: string) {
	const book = await loadXlsx(
		readFileSync(
			path.join(import.meta.dirname, '../../src/core/xlsx/__fixtures__/excel-chart-families.xlsx'),
		),
	);
	const drawing = book.sheets[0]!.drawings.find(
		(d) => d.kind === 'unsupported' && d.chartExSpace?.plotArea.series[0]?.layoutId === 'waterfall',
	);
	if (drawing?.kind !== 'unsupported' || !drawing.chartExPartName)
		throw new Error('Missing ChartEx source host');
	drawing.anchor = {
		from: { row: 0, col: 0, rowOffset: 0, colOffset: 0 },
		ext: { cx: 600 * 9525, cy: 350 * 9525 },
	};
	book.sheets[0]!.drawings = [drawing];
	const zip = await JSZip.loadAsync(await saveXlsx(book));
	const xml = chartXml(kind);
	zip.file(drawing.chartExPartName, xml);
	return {
		part: drawing.chartExPartName,
		xml,
		bytes: await zip.generateAsync({ type: 'uint8array' }),
	};
}
for (const framework of FRAMEWORKS)
	for (const kind of ['funnel', 'boxWhisker', 'regionMap'])
		test(`${framework}: shared ${kind} paints in Excel and saves its source unchanged`, async ({
			page,
		}, testInfo) => {
			const source = await fixture(kind);
			await page.setViewportSize({ width: 1400, height: 900 });
			await openLanding(page, framework);
			await page.locator('#landing-file').setInputFiles({
				name: 'neutral-chart.xlsx',
				mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
				buffer: Buffer.from(source.bytes),
			});
			const host = editor(page);
			const chart = host.locator('.xg-obj[data-index="0"]');
			await expect(
				chart.locator(
					kind === 'regionMap'
						? '[data-map-geometry="cached"] path[data-chart-point]'
						: kind === 'boxWhisker'
							? '[data-chart-part="box"]'
							: 'g[data-chart-point]',
				),
			).toHaveCount(kind === 'regionMap' ? 1 : kind === 'boxWhisker' ? 2 : 6);
			if (kind === 'regionMap')
				await expect(
					chart.locator('[data-map-geometry="cached"] path[data-chart-point]'),
				).toHaveAttribute('fill', '#00aa55');
			await testInfo.attach(`${kind}-browser.png`, {
				body: await chart.screenshot(),
				contentType: 'image/png',
			});
			const bytes = Uint8Array.from(
				await host.evaluate(async (element) =>
					Array.from(
						await (element as unknown as { saveBytes(): Promise<Uint8Array> }).saveBytes(),
					),
				),
			);
			const saved = await JSZip.loadAsync(bytes);
			expect(await saved.file(source.part)!.async('string')).toBe(source.xml);
		});
