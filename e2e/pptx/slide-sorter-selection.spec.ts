/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright owns the test globals */
import { expect, test } from '@playwright/test';

import { loadDeckAt, SAMPLE_DECK } from './support/deck';

test('sorter pastes retained slides after a different destination and undoes the batch', async ({
	page,
	baseURL,
}) => {
	await loadDeckAt(page, baseURL!, SAMPLE_DECK);
	await page
		.getByRole('button', { name: /slide sorter/iu })
		.last()
		.click();
	const tiles = page.locator('[data-pptx-chrome="sorter-tile"]');
	const count = await tiles.count();
	const content = (index: number) =>
		tiles
			.nth(index)
			.evaluate((node) =>
				Array.from(node.querySelectorAll('[data-pptx-element="true"]')).map(
					(element) => element.textContent,
				),
			);
	const first = await content(0),
		second = await content(1);
	await tiles.nth(0).click();
	await tiles.nth(1).click({ modifiers: ['Shift'] });
	await page.keyboard.press('Control+c');
	await tiles.nth(count - 1).click();
	await page.keyboard.press('Control+v');
	await expect(tiles).toHaveCount(count + 2);
	await expect.poll(() => content(count)).toEqual(first);
	await expect.poll(() => content(count + 1)).toEqual(second);
	await expect(
		page.locator('[data-pptx-chrome="sorter-tile"][data-pptx-selected="true"]'),
	).toHaveCount(2);
	await page.keyboard.press('Control+z');
	await expect(tiles).toHaveCount(count);
	await page.keyboard.press('Control+y');
	await expect(tiles).toHaveCount(count + 2);
});

test('sorter range selection, multi-slide clipboard and zoom stay in sync', async ({
	page,
	baseURL,
}) => {
	await loadDeckAt(page, baseURL!, SAMPLE_DECK);
	await page
		.getByRole('button', { name: /slide sorter/iu })
		.last()
		.click();
	const tiles = page.locator('[data-pptx-chrome="sorter-tile"]');
	const count = await tiles.count();
	expect(count).toBeGreaterThan(2);
	await tiles.nth(0).click();
	await tiles.nth(2).click({ modifiers: ['Control'] });
	await tiles.nth(0).click({ modifiers: ['Control', 'Shift'] });
	await expect(
		page.locator('[data-pptx-chrome="sorter-tile"][data-pptx-selected="true"]'),
	).toHaveCount(3);
	await tiles.nth(0).click();
	await tiles.nth(2).click({ modifiers: ['Shift'] });
	await expect(
		page.locator('[data-pptx-chrome="sorter-tile"][data-pptx-selected="true"]'),
	).toHaveCount(3);
	await tiles.nth(1).click({ button: 'right' });
	await page.getByRole('menuitem', { name: 'Copy (3)', exact: true }).click();
	await page.keyboard.press('Control+v');
	await expect(tiles).toHaveCount(count + 3);
	const zoom = page.getByRole('slider', { name: /zoom/iu });
	await expect(zoom).toHaveValue('100');
	await page.keyboard.press('Control+=');
	await expect(zoom).toHaveValue('110');
	await page.keyboard.press('Escape');
	await expect(tiles).toHaveCount(count + 3);
	await expect(
		page.locator('[data-pptx-chrome="sorter-tile"][data-pptx-selected="true"]'),
	).toHaveCount(1);
	await page.keyboard.press('Escape');
	await expect(tiles).toHaveCount(0);
});
