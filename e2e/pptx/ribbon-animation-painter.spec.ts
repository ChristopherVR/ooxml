/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright API */
import { expect, test } from '@playwright/test';
import JSZip from 'jszip';

import { savePptxViaBackstage } from './save-pptx';
import {
	elementWithText,
	fixture,
	loadDeck,
	openRibbonTab,
	selectElement,
	thumbnail,
	viewport,
} from './support/deck';
import { downloadBytes } from './support/exports';

test.use({ viewport: { width: 1440, height: 900 } });

test('painting a native effect across slides copies its embedded sound', async ({ page }) => {
	await loadDeck(page, fixture('native-animation-painter.pptx'));
	await selectElement(page, elementWithText(page, 'Native animation timing'));
	await openRibbonTab(page, 'Animations');
	const painter = page.getByRole('button', { name: 'Animation Painter', exact: true });
	await painter.click();
	await expect(painter).toHaveAttribute('aria-pressed', 'true');
	await thumbnail(page, 2).click();
	await selectElement(page, elementWithText(page, 'Native animation target'));
	await expect(painter).toHaveAttribute('aria-pressed', 'false');

	const bytes = await downloadBytes(await savePptxViaBackstage(page));
	const zip = await JSZip.loadAsync(bytes);
	const slide = await zip.file('ppt/slides/slide2.xml')?.async('string');
	const rels = await zip.file('ppt/slides/_rels/slide2.xml.rels')?.async('string');
	expect(slide).toMatch(/<p:cTn\b(?=[^>]*presetID="1")(?=[^>]*presetClass="entr")[^>]*>/u);
	const id = /<p:sndTgt\b[^>]*r:embed="(rId\d+)"/u.exec(slide ?? '')?.[1];
	const audioRel = rels
		?.match(/<Relationship\b[^>]*\/audio"[^>]*>/gu)
		?.find((entry) => entry.includes(`Id="${id}"`));
	const target = /Target="\.\.\/media\/([^"]+)"/u.exec(audioRel ?? '')?.[1];
	expect(target).toBeDefined();
	expect(target).not.toBe('audio1.wav');
	const original = await zip.file('ppt/media/audio1.wav')?.async('uint8array');
	expect(original?.byteLength).toBeGreaterThan(44);
	expect(await zip.file(`ppt/media/${target}`)?.async('uint8array')).toEqual(original);
});

test('paints animation timing, undoing each application independently', async ({ page }) => {
	await loadDeck(page);
	const source = elementWithText(page, 'Product Overview');
	const first = viewport(page).getByRole('group', { name: 'Project Atlas', exact: true });
	const second = elementWithText(page, 'Q2 2026');
	await selectElement(page, source);
	await openRibbonTab(page, 'Animations');

	const tab = page.locator('pptx-ui-ribbon-animations:visible');
	const timing = (field: string) =>
		tab.locator(`[data-ribbon-control="animations.timing.${field}"]`);
	const painter = tab.getByRole('button', { name: 'Animation Painter', exact: true });
	const duration = timing('duration');
	const delay = timing('delay');
	const start = timing('start').getByRole('combobox');
	await expect(painter).toBeDisabled();
	await expect(duration).toBeDisabled();

	await tab.getByRole('button', { name: 'Fly In', exact: true }).click();
	await expect(painter).toBeEnabled();
	await duration.fill('1.25');
	await duration.press('Tab');
	await expect(duration).toHaveValue('1.25');
	await delay.fill('0.4');
	await delay.press('Tab');
	await expect(delay).toHaveValue('0.4');

	await painter.click();
	await expect(painter).toHaveAttribute('aria-pressed', 'true');
	await selectElement(page, first);
	await expect(painter).toHaveAttribute('aria-pressed', 'false');
	await expect(duration).toHaveValue('1.25');
	await expect(delay).toHaveValue('0.4');
	await page.getByRole('button', { name: 'Undo', exact: true }).first().click();
	await selectElement(page, first);
	await expect(duration).toBeDisabled();

	await selectElement(page, source);
	await expect(duration).toHaveValue('1.25');
	await painter.dblclick();
	await selectElement(page, first);
	await expect(painter).toHaveAttribute('aria-pressed', 'true');
	await selectElement(page, second);
	await expect(duration).toHaveValue('1.25');
	await expect(painter).toHaveAttribute('aria-pressed', 'true');
	await page.keyboard.press('Escape');
	await expect(painter).toHaveAttribute('aria-pressed', 'false');

	await page.getByRole('button', { name: 'Undo', exact: true }).first().click();
	await selectElement(page, second);
	await expect(duration).toBeDisabled();
	await selectElement(page, first);
	await expect(duration).toHaveValue('1.25');
	await page.getByRole('button', { name: 'Redo', exact: true }).first().click();
	await selectElement(page, second);
	await expect(duration).toHaveValue('1.25');

	await start.click();
	await page.getByRole('option', { name: 'After Previous', exact: true }).click();
	await expect(start).toContainText('After Previous');
	const reorder = timing('reorder');
	await expect(reorder.getByRole('button', { name: 'Move later' })).toBeDisabled();
	await reorder.getByRole('button', { name: 'Move earlier' }).click();
	await expect(reorder.getByRole('button', { name: 'Move later' })).toBeEnabled();

	const bytes = await downloadBytes(await savePptxViaBackstage(page));
	const zip = await JSZip.loadAsync(bytes);
	const xml = await zip.file('ppt/slides/slide1.xml')!.async('string');
	expect(xml.match(/<p:cTn[^>]*presetClass="entr"[^>]*>/gu)).toHaveLength(3);
	expect(xml).toContain('nodeType="afterEffect"');
	expect(xml).toContain('dur="1250"');
	expect(xml).toContain('delay="400"');
});
