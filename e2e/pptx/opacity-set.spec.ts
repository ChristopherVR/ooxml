/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec */
import { expect, test } from '@playwright/test';

import { resetTabSession, slideStage } from './support/deck';
import { opacitySetDeck } from './support/opacity-set-deck';

for (const hold of [true, false]) {
	test(`a native fixed whole opacity ${hold ? 'holds through seek and backward navigation' : 'restores after its active duration'}`, async ({
		page,
	}) => {
		await resetTabSession(page);
		await page.goto('/');
		await page.locator('#file-input').setInputFiles({
			name: 'opacity-set.pptx',
			mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
			buffer: await opacitySetDeck(hold),
		});
		await expect(slideStage(page)).toContainText('Persistent text');
		const show = page.getByRole('button', { name: /^slide show$/iu });
		if (await show.count()) await show.last().click();
		else
			await page
				.getByRole('button', { name: /present/iu })
				.first()
				.click();
		const shape = () =>
			page.locator('[data-element-id]').filter({ hasText: 'Persistent text' }).last();
		await expect.poll(() => shape().evaluate((node) => getComputedStyle(node).opacity)).toBe('1');
		await page.keyboard.press('PageDown');
		if (hold) {
			await expect.poll(() => shape().evaluate((node) => getComputedStyle(node).opacity)).toBe('1');
			await page.keyboard.press('PageDown');
		}
		await expect.poll(() => shape().evaluate((node) => getComputedStyle(node).opacity)).toBe('0.4');
		await expect
			.poll(() => shape().evaluate((node) => getComputedStyle(node).animationName))
			.toBe('none');
		await expect(shape().getByText('Persistent text', { exact: true })).toBeVisible();
		if (hold) {
			await page.keyboard.press('PageDown');
			await page.keyboard.press('PageDown');
			await expect(slideStage(page)).toContainText('Next slide');
			await page.keyboard.press('PageUp');
			await expect
				.poll(() => shape().evaluate((node) => getComputedStyle(node).opacity))
				.toBe('0.4');
		} else
			await expect.poll(() => shape().evaluate((node) => getComputedStyle(node).opacity)).toBe('1');
		await page.keyboard.press('Escape');
		await expect
			.poll(() =>
				slideStage(page)
					.locator('[data-element-id]')
					.filter({ hasText: 'Persistent text' })
					.first()
					.evaluate((node) => getComputedStyle(node).opacity),
			)
			.toBe('1');
	});
}
