/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec */
import { writeFile } from 'node:fs/promises';

import { expect, test } from '@playwright/test';
import JSZip from 'jszip';

import { savePptxViaBackstage } from './save-pptx';
import { afterEffectDeck } from './support/after-effect-deck';
import { elementWithText, inspector, loadDeck, ribbonTab, selectElement } from './support/deck';
import { downloadBytes } from './support/exports';
import { chooseSelectValue, expectSelectValue } from './support/select-control';

test.use({ viewport: { width: 1440, height: 900 } });
for (const [text, action, clicks] of [
	['RectDimRGB', 'dimToColor', 2],
	['TextDimScheme', 'dimToColor', 3],
	['RectHide', 'hideAfterAnimation', 3],
	['RectHideClick', 'hideOnNextClick', 5],
] as const) {
	test(`explicit Don't dim clears imported ${action} on ${text} with history and export`, async ({
		page,
	}, info) => {
		const sourcePath = info.outputPath('native-after-effects-with-following-click.pptx');
		await writeFile(sourcePath, await afterEffectDeck());
		await loadDeck(page, sourcePath);
		await selectElement(page, elementWithText(page, text));
		await ribbonTab(page, 'Animations').click();
		if (!(await inspector(page).isVisible()))
			await page
				.locator('[data-ribbon-control="animations.advancedAnimation.animationPane"] button')
				.click();
		const control = inspector(page).locator('pptx-ui-select[aria-label="After animation"]');
		await expectSelectValue(control, action);
		if (text === 'TextDimScheme')
			await expect(inspector(page).getByLabel('Dim color', { exact: true })).not.toBeVisible();
		await chooseSelectValue(page, control, 'none');
		await expectSelectValue(control, 'none');
		await page.getByRole('button', { name: 'Undo', exact: true }).first().click();
		await selectElement(page, elementWithText(page, text));
		await expectSelectValue(control, action);
		await page.getByRole('button', { name: 'Redo', exact: true }).first().click();
		await selectElement(page, elementWithText(page, text));
		await expectSelectValue(control, 'none');
		const bytes = await downloadBytes(await savePptxViaBackstage(page));
		const zip = await JSZip.loadAsync(bytes);
		const xml = await zip.file('ppt/slides/slide1.xml')!.async('string');
		expect(xml).toContain('afterAnimation="none"');
		const path = info.outputPath('after-effect-cleared.pptx');
		await writeFile(path, bytes);
		await loadDeck(page, path);
		await selectElement(page, elementWithText(page, text));
		await ribbonTab(page, 'Animations').click();
		if (!(await inspector(page).isVisible()))
			await page
				.locator('[data-ribbon-control="animations.advancedAnimation.animationPane"] button')
				.click();
		await expectSelectValue(control, 'none');
		const show = page.getByRole('button', { name: /^slide show$/iu });
		if (await show.count()) await show.last().click();
		else
			await page
				.getByRole('button', { name: /present/iu })
				.first()
				.click();
		await expect(page.locator('pptx-ui-present-toolbar:visible')).toHaveCount(1);
		const rendered = page.locator(`[data-element-id][aria-label="${text}"]`).last();
		const order = ['RectDimRGB', 'TextDimScheme', 'RectHide', 'RectHideClick', 'Next trigger'];
		for (let i = 0; i < clicks; i++) {
			await page.keyboard.press('PageDown');
			const current = page.locator(`[data-element-id][aria-label="${order[i]!}"]`).last();
			await current.evaluate(async (node) => {
				await Promise.all(
					node
						.getAnimations({ subtree: true })
						.map((animation) => animation.finished.catch(() => undefined)),
				);
			});
		}
		await expect(rendered).toBeVisible();
		await expect(page.locator('pptx-ui-present-toolbar:visible')).toHaveCount(1);
		await expect
			.poll(() => rendered.evaluate((node) => getComputedStyle(node).animationName))
			.not.toMatch(/dim/u);
		await page.keyboard.press('Escape');
	});
}
