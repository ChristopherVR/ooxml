import { readFileSync } from 'node:fs';

import { expect, test } from '@playwright/test';
import JSZip from 'jszip';

import { resetTabSession } from './support/deck';
test('cached geography paints provider polygons through every PowerPoint binding', async ({
	page,
}, testInfo) => {
	const zip = await JSZip.loadAsync(
		readFileSync(new URL('./fixtures/chart-gallery.pptx', import.meta.url)),
	);
	let chartPart: string | undefined;
	for (const name of Object.keys(zip.files))
		if (
			/^ppt\/charts\/chart[^/]*\.xml$/u.test(name) &&
			(await zip.file(name)!.async('string')).includes('layoutId="regionMap"')
		) {
			chartPart = name;
			break;
		}
	if (!chartPart) throw new Error('Missing region map host');
	// A published Bing PCA vector, transformed into a synthetic polygon, never provider geodata.
	zip.file(
		chartPart,
		`<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"><cx:chartData><cx:data id="0"><cx:strDim type="cat"><cx:lvl ptCount="1"><cx:pt idx="0">Sample</cx:pt></cx:lvl></cx:strDim><cx:numDim type="colorVal"><cx:lvl ptCount="1"><cx:pt idx="0">42</cx:pt></cx:lvl></cx:numDim></cx:data></cx:chartData><cx:chart><cx:plotArea><cx:plotAreaRegion><cx:series layoutId="regionMap"><cx:valueColors><cx:minColor><a:srgbClr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" val="FF0000"/></cx:minColor><cx:midColor><a:srgbClr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" val="00FF00"/></cx:midColor><cx:maxColor><a:srgbClr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" val="0000FF"/></cx:maxColor></cx:valueColors><cx:valueColorPositions count="3"><cx:min><cx:extremeValue/></cx:min><cx:mid><cx:percent val="25"/></cx:mid><cx:max><cx:extremeValue/></cx:max></cx:valueColorPositions><cx:dataId val="0"/><cx:layoutPr><cx:geography projectionType="robinson"><cx:geoCache><cx:clear><cx:geoData entityId="sample" entityName="Sample"><cx:geoPolygons><cx:geoPolygon numPoints="4" pcaRings="1,vx1vilihnM6hR7mEl2Q"/></cx:geoPolygons></cx:geoData></cx:clear></cx:geoCache></cx:geography></cx:layoutPr></cx:series></cx:plotAreaRegion></cx:plotArea></cx:chart></cx:chartSpace>`,
	);
	const presentation = await zip.file('ppt/presentation.xml')!.async('string');
	const ids = [...presentation.matchAll(/<p:sldId\b[^>]*?(?:\/>|>\s*<\/p:sldId>)/gu)].map(
		(m) => m[0],
	);
	if (!ids[19]) throw new Error('Missing region map slide in the gallery manifest');
	zip.file(
		'ppt/presentation.xml',
		presentation.replace(
			/<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/u,
			`<p:sldIdLst>${ids[19]}</p:sldIdLst>`,
		),
	);
	await resetTabSession(page);
	await page.goto('/');
	await page.locator('#file-input').setInputFiles({
		name: 'cached-map.pptx',
		mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
		buffer: await zip.generateAsync({ type: 'nodebuffer' }),
	});
	const path = page
		.locator('[data-pptx-element="true"] svg path')
		.filter({ has: page.locator('title', { hasText: 'Sample: 42' }) })
		.first();
	await expect(path).toBeVisible();
	await expect(path).toHaveAttribute('d', /M.*L.*Z/u);
	await expect(path).toHaveAttribute('fill', '#00aa55');
	await testInfo.attach('cached-map.png', {
		body: await page.locator('[data-pptx-viewport]').first().screenshot(),
		contentType: 'image/png',
	});
});
