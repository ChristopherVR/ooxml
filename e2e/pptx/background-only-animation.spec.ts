/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec */
import { expect, test } from '@playwright/test';

import { backgroundAnimationDeck } from './support/background-animation-deck';
import { resetTabSession, slideStage } from './support/deck';

test('a native background target animates shape paint while its text stays visible', async ({
	page,
}, testInfo) => {
	await resetTabSession(page);
	await page.goto('/');
	await page.locator('#file-input').setInputFiles({
		name: 'background-target.pptx',
		mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
		buffer: await backgroundAnimationDeck(),
	});
	await expect(slideStage(page)).toContainText('Persistent text');
	const show = page.getByRole('button', { name: /^slide show$/iu });
	if (await show.count()) await show.last().click();
	else
		await page
			.getByRole('button', { name: /present/iu })
			.first()
			.click();
	const paint = page.locator('[data-pptx-animation-layer="background"]').last();
	await expect(paint).toBeAttached();
	await expect
		.poll(() => paint.evaluate((node) => getComputedStyle(node).visibility))
		.toBe('hidden');
	const shape = paint.locator('xpath=..');
	await expect(shape).toContainText('Persistent text');
	await expect
		.poll(() => shape.evaluate((node) => getComputedStyle(node).visibility))
		.toBe('visible');
	await expect(shape.getByText('Persistent text', { exact: true })).toBeVisible();
	await page.screenshot({ path: testInfo.outputPath('text-before-background.png') });
	await page.keyboard.press('PageDown');
	await expect(paint).toBeVisible();
	await expect
		.poll(() => paint.evaluate((node) => getComputedStyle(node).animationName))
		.not.toBe('none');
	await expect(shape.getByText('Persistent text', { exact: true })).toBeVisible();
	await page.screenshot({ path: testInfo.outputPath('animated-background.png') });
	await page.keyboard.press('Escape');
	await expect(slideStage(page)).toContainText('Persistent text');
	await expect(slideStage(page).locator('[data-pptx-animation-layer="background"]')).toHaveCount(0);
	await expect.poll(() => slideStage(page).locator('[data-element-id]').filter({ hasText: 'Persistent text' }).first()
		.evaluate((node) => getComputedStyle(node).backgroundColor)).toBe('rgb(0, 170, 0)');
});
