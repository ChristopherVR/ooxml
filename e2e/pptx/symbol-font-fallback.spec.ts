/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec */
import { readFile, writeFile } from 'node:fs/promises';

import { expect, test } from '@playwright/test';
import JSZip from 'jszip';

import { fixture, loadDeck, slideElements } from './support/deck';

test('symbol bullets have an offline last-resort font', async ({ page }, testInfo) => {
	// Simulate a device without Wingdings even on Windows. Keep real font loading
	// and rasterisation for the bundled fallback; only the installed-font probe changes.
	await page.addInitScript(() => {
		const measure = CanvasRenderingContext2D.prototype.measureText;
		CanvasRenderingContext2D.prototype.measureText = function (text) {
			const font = this.font;
			this.font = font.replace(/Wingdings/giu, 'Missing Symbol Face');
			try {
				return measure.call(this, text);
			} finally {
				this.font = font;
			}
		};
	});
	await page.route('**/fonts.googleapis.com/**', (route) => route.abort());
	await page.route('**/fonts.gstatic.com/**', (route) => route.abort());
	const zip = await JSZip.loadAsync(await readFile(fixture('sample-deck.pptx')));
	const markers = ['§', 'q', 'v', 'Ø', 'ü'];
	const paragraphs = markers
		.map(
			(char, index) =>
				`<a:p><a:pPr marL="342900" indent="-285750"><a:buFont typeface="Wingdings"/><a:buChar char="${char}"/></a:pPr><a:r><a:rPr sz="3200"><a:latin typeface="Georgia"/></a:rPr><a:t>Symbol ${index + 1}</a:t></a:r></a:p>`,
		)
		.join('');
	zip.file(
		'ppt/slides/slide1.xml',
		`<?xml version="1.0" encoding="UTF-8"?><p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="Symbols"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="7315200" cy="4572000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/>${paragraphs}</p:txBody></p:sp></p:spTree></p:cSld></p:sld>`,
	);
	const path = testInfo.outputPath('symbols.pptx');
	await writeFile(path, await zip.generateAsync({ type: 'nodebuffer' }));
	await loadDeck(page, path);
	const content = slideElements(page).filter({ hasText: 'Symbol 1' }).first();
	await expect(content).toContainText('▪');
	for (const marker of ['❑', '❖', '➢', '✔']) await expect(content).toContainText(marker);
	await expect(content).not.toContainText('§');
	const report = await page.evaluate(async () => {
		const faces = await document.fonts.load('32px "OOXML Noto Symbols"', '▪❑❖➢✔');
		return {
			loaded: faces.length > 0 && faces.every((face) => face.status === 'loaded'),
			inline: document
				.getElementById('ooxml-symbol-fallback')
				?.textContent?.includes('data:font/woff2;base64,'),
		};
	});
	expect(report).toEqual({ loaded: true, inline: true });
	await content.screenshot({ path: testInfo.outputPath('symbols.png') });
});
