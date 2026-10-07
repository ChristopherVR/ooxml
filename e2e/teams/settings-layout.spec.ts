import { expect, test } from '@playwright/test';

test('keeps settings navigation accessible while scrolling and supports keyboard switches', async ({
	page,
}) => {
	await page.goto(`/?local=1&name=Ada&room=settings-layout-${Date.now()}`);
	await expect(page.getByRole('heading', { name: '# General', exact: true })).toBeVisible({
		timeout: 15_000,
	});
	await page.getByRole('button', { name: 'Settings and more', exact: true }).click();
	await page.getByRole('menuitem', { name: 'Settings', exact: true }).click();
	const settings = page.locator('teams-settings');
	const close = page.getByRole('button', { name: 'Close settings', exact: true });
	for (const viewport of [
		{ width: 1280, height: 600 },
		{ width: 390, height: 600 },
	]) {
		await page.setViewportSize(viewport);
		await page.getByRole('tab', { name: 'Connection', exact: true }).click();
		const before = (await close.boundingBox())!;
		await page.getByRole('textbox', { name: /^Sync URL/ }).fill('ws://draft.invalid/sync');
		await page
			.getByRole('button', { name: 'Apply and reconnect', exact: true })
			.scrollIntoViewIfNeeded();
		await expect(close).toBeInViewport();
		await expect(
			page.getByRole('searchbox', { name: 'Search settings', exact: true }),
		).toBeInViewport();
		const after = (await close.boundingBox())!;
		expect(after.y).toBeCloseTo(before.y, 1);
		expect(
			await settings.locator('.settings-content').evaluate((element) => element.scrollTop),
		).toBeGreaterThan(0);
		const dialog = (await page
			.getByRole('dialog', { name: 'Settings', exact: true })
			.boundingBox())!;
		expect(dialog.y).toBeGreaterThanOrEqual(0);
		expect(dialog.y + dialog.height).toBeLessThanOrEqual(viewport.height);
		await page.getByRole('tab', { name: 'Notifications and activity', exact: true }).click();
		const toggle = page.getByRole('switch', { name: 'Threads I start', exact: true });
		await toggle.focus();
		await toggle.press('Space');
		await expect(toggle).not.toBeChecked();
		await toggle.press('Enter');
		await expect(toggle).toBeChecked();
		await page.screenshot({
			path: test.info().outputPath(`settings-switches-${viewport.width}.png`),
		});
		await page.getByRole('tab', { name: 'Connection', exact: true }).click();
		await expect(page.getByRole('textbox', { name: /^Sync URL/ })).toHaveValue(
			'ws://draft.invalid/sync',
		);
		await page.screenshot({
			path: test.info().outputPath(`settings-scroll-${viewport.width}.png`),
		});
	}
	await close.click();
	await expect(page.getByRole('dialog', { name: 'Settings', exact: true })).not.toBeVisible();
});
