/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright API */
import { expect, test } from '@playwright/test';
import JSZip from 'jszip';

import { savePptxViaBackstage } from './save-pptx';
import { elementWithText, loadDeck, openRibbonTab, selectElement, viewport } from './support/deck';
import { downloadBytes } from './support/exports';

test.use({ viewport: { width: 1440, height: 900 } });

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
