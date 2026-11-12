/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright API */
import { readFile } from 'node:fs/promises';

import { expect, test } from '@playwright/test';
import type { Locator } from '@playwright/test';
import JSZip from 'jszip';

import { savePptxViaBackstage } from './save-pptx';
import {
	elementWithText,
	fixture,
	loadDeck,
	resetTabSession,
	ribbonTab,
	selectElement,
	slideStage,
} from './support/deck';
import { downloadBytes } from './support/exports';
import { extractElementBlock, readZipPartText } from './support/pptx-xml';
import { chooseSelectValue, expectSelectValue } from './support/select-control';

test.use({ viewport: { width: 1600, height: 1000 } });

/** Adapt an ordinary deck in memory, without regenerating committed fixtures. */
async function weightPickerDeck(): Promise<Buffer> {
	const zip = await JSZip.loadAsync(await readFile(fixture('calibri-metric-webfont.pptx')));
	const theme = zip.file('ppt/theme/theme1.xml')!;
	zip.file(
		theme.name,
		(await theme.async('string'))
			.replace('typeface="Calibri Light"', 'typeface="DM Sans"')
			.replace('typeface="Calibri"', 'typeface="ADLaM Display"'),
	);
	const slide = zip.file('ppt/slides/slide1.xml')!;
	zip.file(
		slide.name,
		(await slide.async('string')).replaceAll('typeface="Calibri"', 'typeface="DM Sans"'),
	);
	return zip.generateAsync({ type: 'nodebuffer' });
}

/** Read computed CSS from a real text node, not the shape's outer selection wrapper. */
async function renderedTextStyle(element: Locator) {
	return element.evaluate((shape) => {
		const walker = document.createTreeWalker(shape, NodeFilter.SHOW_TEXT);
		let node: Node | null;
		while ((node = walker.nextNode())) {
			if (node.textContent?.trim() && node.parentElement) {
				const style = getComputedStyle(node.parentElement);
				return { fontFamily: style.fontFamily, fontWeight: style.fontWeight };
			}
		}
		throw new Error('The selected shape has no rendered text');
	});
}

test('font weight picker renders Medium at 500 and preserves its family through save and reopen', async ({
	page,
}, info) => {
	// The property under test is the requested CSS weight, not a live CDN response.
	await page.route('**/fonts.googleapis.com/css2**', (route) =>
		route.fulfill({ contentType: 'text/css', body: '' }),
	);
	await resetTabSession(page);
	await page.goto('/');
	await page.locator('#file-input').setInputFiles({
		name: 'font-weight-picker.pptx',
		mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
		buffer: await weightPickerDeck(),
	});
	await slideStage(page).waitFor();
	const shape = elementWithText(page, 'Box A');
	await expect(shape).toBeVisible();
	await ribbonTab(page, 'Home').click();
	await selectElement(page, shape);
	const family = page.locator('[data-ribbon-control="home.font.fontFamily"]').getByRole('combobox');
	await expectSelectValue(family, 'DM Sans');
	await expect.poll(() => renderedTextStyle(shape)).toMatchObject({ fontWeight: '400' });

	const rows = await page
		.locator('[data-font-picker="family"] option')
		.evaluateAll((options) => options.map((option) => (option as HTMLOptionElement).value));
	for (const face of ['DM Sans Light', 'DM Sans Medium', 'DM Sans SemiBold']) {
		expect(rows).toContain(face);
	}
	expect(rows).not.toContain('DM Sans Regular');
	expect(rows).toContain('ADLaM Display');
	expect(rows).not.toContain('ADLaM Display Medium');

	await chooseSelectValue(page, family, 'DM Sans Medium');
	await expectSelectValue(family, 'DM Sans Medium');
	await expect.poll(() => renderedTextStyle(shape)).toMatchObject({ fontWeight: '500' });
	expect((await renderedTextStyle(shape)).fontFamily).toContain('DM Sans');
	await family.click();
	await page.getByRole('option', { name: 'DM Sans Headings', exact: true }).click();
	await expect.poll(() => renderedTextStyle(shape)).toMatchObject({ fontWeight: '400' });
	await chooseSelectValue(page, family, 'DM Sans Medium');
	await expect.poll(() => renderedTextStyle(shape)).toMatchObject({ fontWeight: '500' });

	const download = await savePptxViaBackstage(page);
	const bytes = await downloadBytes(download);
	const xml = await readZipPartText(bytes, 'ppt/slides/slide1.xml');
	const savedShape = extractElementBlock(xml, 'p:sp', 'Box A');
	expect(savedShape).toMatch(/<a:latin\b[^>]*\btypeface="DM Sans Medium"/u);
	const saved = info.outputPath('font-weight-medium.pptx');
	await download.saveAs(saved);
	await loadDeck(page, saved);
	const reopened = elementWithText(page, 'Box A');
	await expect.poll(() => renderedTextStyle(reopened)).toMatchObject({ fontWeight: '500' });
	await ribbonTab(page, 'Home').click();
	await selectElement(page, reopened);
	await expectSelectValue(family, 'DM Sans Medium');
});
