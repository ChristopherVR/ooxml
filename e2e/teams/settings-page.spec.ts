import { expect, test } from '@playwright/test';

test('keeps the shell available and preserves the conversation behind the settings page', async ({
	page,
}) => {
	await page.goto(`/?local=1&name=Ada&room=settings-page-${Date.now()}`);
	await expect(page.getByRole('heading', { name: '# General', exact: true })).toBeVisible({
		timeout: 15_000,
	});
	const composer = page.getByRole('textbox', { name: 'Message', exact: true });
	await composer.fill('Keep this unsent conversation draft');
	const more = page.getByRole('button', { name: 'Settings and more', exact: true });
	const open = async () => {
		await more.click();
		await page.getByRole('menuitem', { name: 'Settings', exact: true }).click();
	};
	await open();
	const settings = page.getByRole('dialog', { name: 'Settings', exact: true });
	const search = page.getByRole('searchbox', { name: 'Search settings', exact: true });
	await expect(search).toBeFocused();
	await expect(composer).toBeHidden();
	expect(await settings.evaluate((element) => element.matches(':modal'))).toBe(false);
	const topbar = page.locator('teams-app').locator('.topbar');
	const rail = page.getByRole('navigation', { name: 'App bar', exact: true });
	const bounds = (await settings.boundingBox())!;
	const topBounds = (await topbar.boundingBox())!;
	const railBounds = (await rail.boundingBox())!;
	expect(bounds.y).toBeGreaterThanOrEqual(topBounds.y + topBounds.height - 1);
	expect(bounds.x).toBeGreaterThanOrEqual(railBounds.x + railBounds.width - 1);
	await search.press('Escape');
	await expect(settings).toBeHidden();
	await expect(more).toBeFocused();
	await expect(composer).toHaveValue('Keep this unsent conversation draft');
	await open();
	await rail.getByRole('button', { name: 'Files', exact: true }).click();
	await expect(settings).toBeHidden();
	await expect(page.getByRole('heading', { name: 'Files', exact: true })).toBeVisible();
	await expect(page.getByRole('main')).not.toHaveAttribute('inert');
	await open();
	await page.setViewportSize({ width: 390, height: 844 });
	await expect(more).toBeVisible();
	await expect(rail.getByRole('button', { name: 'Files', exact: true })).toBeVisible();
	await expect(page.getByRole('button', { name: 'Close settings', exact: true })).toBeVisible();
	const mobile = (await settings.boundingBox())!;
	expect(mobile.x + mobile.width).toBeLessThanOrEqual(390);
	expect(mobile.y + mobile.height).toBeLessThanOrEqual(844);
	await page.screenshot({ path: test.info().outputPath('settings-page-mobile.png') });
});
