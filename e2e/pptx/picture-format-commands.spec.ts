/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec */
import { expect, test } from '@playwright/test';
import type { Page } from '@playwright/test';

import { elementsOfType, fixture, loadDeck, selectElement } from './support/deck';

test.use({ viewport: { width: 2048, height: 1080 } });

const control = (page: Page, id: string) =>
	page.locator(`[data-ribbon-control="pictureFormat.${id}"]`).getByRole('button');
const dialog = (page: Page) => page.locator('office-ui-dialog[data-picture-dialog]');

async function pictureFormat(page: Page) {
	await loadDeck(page, fixture('merge-crop.pptx'));
	const picture = elementsOfType(page, 'image').first();
	await selectElement(page, picture);
	await page.locator('[data-ribbon-contextual-tab="pictureFormat"]').first().click();
	return picture;
}

test('picture commands replace the source, compress it and edit the outline', async ({ page }) => {
	const picture = await pictureFormat(page);
	await expect(control(page, 'pictureStyles.layout')).toBeDisabled();
	await control(page, 'adjust.changePicture').click();
	const bytes = await page.evaluate(() => {
		const canvas = document.createElement('canvas');
		canvas.width = 1200;
		canvas.height = 800;
		const ctx = canvas.getContext('2d')!;
		ctx.fillStyle = 'white';
		ctx.fillRect(0, 0, 1200, 800);
		ctx.fillStyle = 'orange';
		ctx.fillRect(300, 200, 600, 400);
		return canvas.toDataURL().split(',')[1]!;
	});
	await dialog(page)
		.locator('input[type="file"]')
		.setInputFiles({
			name: 'replacement.png',
			mimeType: 'image/png',
			buffer: Buffer.from(bytes, 'base64'),
		});
	await dialog(page).getByRole('button', { name: 'Apply' }).click();
	await expect(dialog(page)).toHaveCount(0);

	await page
		.locator('[data-ribbon-group="pictureFormat.size"] [data-pptx-ribbon-control="crop-menu"]')
		.click();
	await page.locator('[data-pptx-crop-aspect="1:1"]:visible').first().click();
	await expect
		.poll(async () => {
			const box = await picture.boundingBox();
			return box ? Math.abs(box.width - box.height) : Infinity;
		})
		.toBeLessThan(3);

	await control(page, 'adjust.compress').click();
	await dialog(page).getByLabel('All pictures in this presentation').check();
	await dialog(page).getByLabel('Resolution (pixels per inch)').selectOption('96');
	await dialog(page).getByLabel('Delete cropped areas of pictures').check();
	await dialog(page).getByRole('button', { name: 'Apply' }).click();
	await expect(dialog(page).getByRole('status')).toContainText(/Picture bytes saved [1-9]/);
	await dialog(page).getByRole('button', { name: 'OK' }).click();

	await control(page, 'pictureStyles.border').click();
	await dialog(page).getByLabel('Weight (pt)').fill('3');
	await dialog(page).getByLabel('Stroke Dash').selectOption('dash');
	await dialog(page).getByRole('button', { name: 'Apply' }).click();
	await expect(dialog(page)).toHaveCount(0);
	await control(page, 'adjust.resetPicture').click();
	await control(page, 'adjust.resetSize').click();
	await expect(dialog(page)).toHaveCount(0);
});

test('background removal supports marking, discard and commit, and picture tools reach the panes', async ({
	page,
}) => {
	await pictureFormat(page);
	await control(page, 'adjust.removeBackground').click();
	const preview = dialog(page).getByRole('button', { name: 'Picture preview' });
	const keep = dialog(page).getByRole('button', { name: 'Keep Changes' });
	await expect(keep).toBeEnabled();
	await expect(preview.locator('[data-pptx-removal-marquee]')).toBeVisible();
	await dialog(page).getByRole('button', { name: 'Mark Areas to Remove' }).click();
	await preview.focus();
	await preview.press('ArrowLeft');
	await preview.press('Space');
	await expect(keep).toBeEnabled();
	await dialog(page).getByRole('button', { name: 'Discard' }).click();
	await expect(dialog(page)).toHaveCount(0);

	await control(page, 'adjust.removeBackground').click();
	await expect(keep).toBeEnabled();
	await preview.focus();
	await preview.press('Space');
	await expect(keep).toBeEnabled();
	await keep.click();
	await expect(dialog(page)).toHaveCount(0);
	await control(page, 'accessibility.altText').click();
	await expect(page.getByTestId('element-alt-text').first()).toBeVisible();
	await control(page, 'arrange.rotateRight').click();
	await control(page, 'arrange.selectionPane').click();
});
