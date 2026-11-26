/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec */
import { expect, test } from '@playwright/test';
import type { Page } from '@playwright/test';

import { openRibbonTab, resetTabSession } from './support/deck';

test.use({ viewport: { width: 1920, height: 1000 } });

function canvas(page: Page) {
	return page.locator('[data-pptx-viewport] [aria-roledescription="slide"]').first();
}

async function draw(page: Page, coordinates: readonly [number, number][]) {
	const stage = await canvas(page).boundingBox();
	if (!stage) throw new Error('Editable slide is missing');
	for (const [index, [x, y]] of coordinates.entries()) {
		await page.mouse.move(stage.x + stage.width * x, stage.y + stage.height * y, { steps: 5 });
		if (index === 0) await page.mouse.down();
	}
	await page.mouse.up();
}

test.beforeEach(async ({ page }) => {
	await resetTabSession(page);
	await page.goto('/');
	await page
		.getByRole('button', { name: /new presentation/iu })
		.first()
		.click();
	await expect(canvas(page)).toBeVisible();
	await openRibbonTab(page, 'Draw');
});

test('draws a rectangle and lasso-converts it to an editable shape', async ({ page }) => {
	await expect(page.getByRole('button', { name: 'Ink to Math' })).toHaveAttribute(
		'aria-description',
		'No handwriting recognizer is configured.',
	);
	await page.getByRole('button', { name: 'Pen', exact: true }).click();
	await draw(page, [
		[0.3, 0.3],
		[0.55, 0.3],
		[0.55, 0.55],
		[0.3, 0.55],
		[0.3, 0.3],
	]);
	const ink = canvas(page).locator(
		'[data-element-id^="ink-"]:not([data-element-id^="ink-conversion-"])',
	);
	await expect(ink).toHaveCount(1);
	const lasso = page.getByRole('button', { name: 'Lasso Select' });
	await lasso.focus();
	await lasso.press('Enter');
	await expect(lasso).toHaveAttribute('aria-pressed', 'true');
	await draw(page, [
		[0.23, 0.23],
		[0.62, 0.23],
		[0.62, 0.62],
		[0.23, 0.62],
		[0.23, 0.23],
	]);
	const convert = page.getByRole('button', { name: 'Ink to Shape' });
	await expect(convert).toBeEnabled();
	await convert.click();
	await expect(ink).toHaveCount(0);
	const shape = canvas(page).locator('[data-element-id^="ink-conversion-"]');
	await expect(shape).toHaveCount(1);
	await page.keyboard.press('Control+z');
	await expect(shape).toHaveCount(0);
	await expect(ink).toHaveCount(1);
	await page.keyboard.press('Control+y');
	await expect(shape).toHaveCount(1);
});

test('keeps a pen preset and exposes keyboard ruler and replay controls', async ({ page }) => {
	const type = page.getByRole('combobox', { name: 'Pen type' });
	await type.click();
	await page.getByRole('option', { name: 'Pencil', exact: true }).click();
	await page.getByRole('button', { name: 'Add Pen' }).click();
	const remove = page.getByRole('button', { name: 'Remove preset' });
	await expect(remove).toBeEnabled();
	const saved = page.getByRole('combobox', { name: 'Pen presets' });
	const pencil = page.getByRole('option', { name: /^Pencil /u });
	await saved.click();
	await expect(pencil).toBeVisible();
	await saved.press('Escape');
	await openRibbonTab(page, 'Home');
	await openRibbonTab(page, 'Draw');
	await saved.click();
	await pencil.click();
	await remove.click();
	await saved.click();
	await expect(pencil).toHaveCount(0);
	await saved.press('Escape');
	await page.getByRole('button', { name: 'Ruler', exact: true }).click();
	const angle = page.getByRole('slider', { name: 'Ruler angle' });
	await angle.press('ArrowRight');
	await expect(angle).toHaveValue('1');
	await expect(canvas(page).locator('[data-pptx-draw-ruler]')).toBeVisible();
	await draw(page, [
		[0.39, 0.5],
		[0.48, 0.52],
		[0.59, 0.5],
	]);
	await page.getByRole('button', { name: 'Ink Replay', exact: true }).click();
	const playOrPause = page.getByRole('button', { name: /^(pause|play)$/iu });
	await expect(playOrPause).toBeVisible();
	const seek = page.getByRole('slider', { name: 'Replay position' });
	await seek.press('Home');
	await expect(canvas(page).locator('[data-pptx-draw-replay]')).toHaveCount(1);
	await page.getByRole('button', { name: 'Close replay' }).click();
	await expect(canvas(page).locator('[data-pptx-draw-replay]')).toHaveCount(0);
});

test('arrow keys move focus between Draw commands and still adjust sliders', async ({ page }) => {
	const draw = page.locator('pptx-ui-ribbon-draw');
	const select = draw.getByRole('button', { name: 'Select', exact: true });
	const pen = draw.getByRole('button', { name: 'Pen', exact: true });
	await select.focus();
	await page.keyboard.press('ArrowRight');
	await expect(pen).toBeFocused();
	await page.keyboard.press('Enter');
	await expect(pen).toHaveAttribute('aria-pressed', 'true');
	await page.keyboard.press('Home');
	await expect(select).toBeFocused();
	const width = draw.getByRole('slider', { name: 'Width', exact: true });
	await width.focus();
	const before = Number(await width.inputValue());
	await page.keyboard.press('ArrowRight');
	await expect(width).toHaveValue(String(before + 1));
	await expect(width).toBeFocused();
});

test('updates the lasso path in place during a drag', async ({ page }) => {
	await page.getByRole('button', { name: 'Lasso Select' }).click();
	const stage = await canvas(page).boundingBox();
	if (!stage) throw new Error('Editable slide is missing');
	const x = stage.x + stage.width * 0.3;
	const y = stage.y + stage.height * 0.3;
	await page.mouse.move(x, y);
	await page.mouse.down();
	const path = canvas(page).locator('[data-pptx-draw-lasso] path');
	await path.evaluate((node) => node.setAttribute('data-test-gesture', 'original'));
	const first = await path.getAttribute('d');
	await page.mouse.move(x + 50, y + 40, { steps: 10 });
	await expect(path).not.toHaveAttribute('d', first ?? '');
	await expect(path).toHaveAttribute('data-test-gesture', 'original');
	await page.mouse.up();
	await expect(path).toHaveCount(0);
});

test('cancels an armed Draw gesture when entering Slide Master', async ({ page }) => {
	const lasso = page.getByRole('button', { name: 'Lasso Select' });
	await lasso.click();
	await expect(lasso).toHaveAttribute('aria-pressed', 'true');
	await openRibbonTab(page, 'View');
	await page.getByRole('button', { name: 'Slide Master', exact: true }).click();
	const masterTabs = page
		.getByRole('tablist')
		.filter({
			has: page.getByRole('tab', { name: 'Handout', exact: true }),
		})
		.first();
	await expect(masterTabs).toBeVisible();
	await expect(lasso).toBeHidden();
	await masterTabs.locator('..').getByRole('button').first().click();
	await expect(masterTabs).toHaveCount(0);
	await openRibbonTab(page, 'Draw');
	await expect(lasso).toHaveAttribute('aria-pressed', 'false');
	await expect(lasso).toBeEnabled();
});
