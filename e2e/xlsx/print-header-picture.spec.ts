import { expect, test } from '@playwright/test';
import { openFixture } from './fixture-files';
import { editor, FRAMEWORKS, pageErrors, ribbon } from './helpers';

// excel-header-images.xlsx (generate-excel-header-images.ps1): sheet "Print" with a different first page,
// whose left header is a picture (&G); the other sections carry pictures too, shown on later pages.

for (const framework of FRAMEWORKS)
	test(`${framework}: print preview shows the first-page header picture`, async ({ page }) => {
		const errors = pageErrors(page);
		await openFixture(page, framework, 'excel-header-images.xlsx');
		await ribbon(page).getByRole('tab', { name: 'Page Layout', exact: true }).click();
		await ribbon(page).locator('[data-command="page.preview"]').click();
		const preview = editor(page).locator('[data-dialog="print-preview"]');
		const first = preview.frameLocator('iframe').locator('.xve-print-page').first();
		await expect(first).toBeVisible();
		const header = first.locator('.xve-print-header');
		// A picture renders from the package image, not as the literal &G code.
		const pictures = header.locator('img');
		await expect(pictures.first()).toBeVisible();
		await expect
			.poll(() => pictures.first().evaluate((img: HTMLImageElement) => img.naturalWidth))
			.toBeGreaterThan(0);
		await expect(header).not.toContainText('&G');
		expect(errors).toEqual([]);
	});
