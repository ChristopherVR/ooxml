/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec */
import { expect, test } from '@playwright/test';
import type { Locator } from '@playwright/test';

import { elementsOfType, fixture, loadDeck, selectElement } from './support/deck';

const formats = [
	{ tab: 'videoFormat', deck: 'media-bookmark-trigger-editable.pptx' },
	{ tab: 'audioFormat', deck: 'audio-embed.pptx' },
	{ tab: 'smartArtFormat', deck: 'smartart-build-reveal.pptx' },
] as const;

async function expectControlsInside(group: Locator, row: Locator): Promise<void> {
	await expect(group).toBeVisible();
	await expect
		.poll(async () => {
			const box = await group.boundingBox();
			const ribbon = await row.boundingBox();
			if (!box || !ribbon) return ['missing bounds'];
			const controls = group.locator('button:visible, input:visible, [role="combobox"]:visible');
			const outside = await controls.evaluateAll(
				(elements, bounds) =>
					elements.flatMap((element) => {
						const rect = element.getBoundingClientRect();
						const inside =
							rect.left >= bounds.x - 1 &&
							rect.top >= bounds.y - 1 &&
							rect.right <= bounds.x + bounds.width + 1 &&
							rect.bottom <= bounds.y + bounds.height + 1;
						return inside ? [] : [element.getAttribute('aria-label') ?? element.textContent];
					}),
				box,
			);
			if (box.x < ribbon.x - 1 || box.x + box.width > ribbon.x + ribbon.width + 1) {
				outside.push('group extends past ribbon');
			}
			return outside;
		})
		.toEqual([]);
}

for (const { tab, deck } of formats) {
	test(`${tab} fits every ribbon group at 1600px and 1280px`, async ({ page }) => {
		await page.setViewportSize({ width: 1600, height: 960 });
		await loadDeck(page, fixture(deck));
		const selected =
			tab === 'smartArtFormat'
				? page
						.locator('[data-pptx-viewport] [data-smartart-node-id]')
						.filter({ hasText: 'Alpha' })
						.first()
				: elementsOfType(page, 'media player').first();
		await selectElement(page, selected);
		await page.locator(`[data-ribbon-contextual-tab="${tab}"]`).click();
		const row = page.locator('[data-pptx-chrome="ribbon-content"]:visible').filter({
			has: page.locator(`[data-ribbon-group="${tab}.arrange"]`),
		});
		const groups = row.locator(`[data-ribbon-group^="${tab}."]`);
		await expect(groups).toHaveCount(tab === 'audioFormat' ? 4 : 5);
		for (const width of [1600, 1280]) {
			await page.setViewportSize({ width, height: 960 });
			await expect
				.poll(() => row.evaluate((el) => el.scrollWidth - el.clientWidth))
				.toBeLessThanOrEqual(1);
			for (const group of await groups.all()) await expectControlsInside(group, row);
		}
	});
}

test('Audio Format Size exposes editable width and height inputs at 1600px', async ({ page }) => {
	await page.setViewportSize({ width: 1600, height: 960 });
	await loadDeck(page, fixture('audio-embed.pptx'));
	await selectElement(page, elementsOfType(page, 'media player').first());
	await page.locator('[data-ribbon-contextual-tab="audioFormat"]').click();
	const size = page.locator('[data-ribbon-group="audioFormat.size"]');
	for (const name of ['Width', 'Height']) {
		const field = size.getByRole('spinbutton', { name, exact: true });
		await expect(field).toBeVisible();
		await expect(field).toBeEnabled();
		await expect(field).toHaveCSS('border-top-width', '1px');
		await expect(field).toHaveValue(/^[1-9]\d*(\.\d+)?$/u);
	}
	const width = size.getByRole('spinbutton', { name: 'Width', exact: true });
	const height = size.getByRole('spinbutton', { name: 'Height', exact: true });
	await width.fill('120');
	await width.press('Tab');
	await height.fill('80');
	await height.press('Tab');
	await page.locator('[data-ribbon-contextual-tab="playback"]').click();
	await page.locator('[data-ribbon-contextual-tab="audioFormat"]').click();
	await expect(width).toHaveValue('120');
	await expect(height).toHaveValue('80');
});
