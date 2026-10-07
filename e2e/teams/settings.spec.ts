import { expect, test } from '@playwright/test';
import type { TeamsApp } from '../../src/ui/src/teams/app/teams-app';

test('organizes settings, saves personal themes and thread choices, and preserves chat on reconnect', async ({
	page,
}) => {
	const room = `settings-${Date.now()}`;
	await page.goto(`/?local=1&name=Ada&room=${room}`);
	await expect(page.getByRole('heading', { name: '# General', exact: true })).toBeVisible({
		timeout: 15_000,
	});
	await page.getByRole('button', { name: 'Settings', exact: true }).click();
	await page.getByRole('tab', { name: 'Connection', exact: true }).click();
	await page.getByRole('combobox', { name: 'Mode', exact: true }).selectOption('server');
	await page.getByRole('button', { name: 'Cancel', exact: true }).click();
	await page.getByRole('button', { name: 'Settings', exact: true }).click();
	await expect(page.getByRole('combobox', { name: 'Mode', exact: true })).toHaveValue('local');
	await expect(page.getByRole('dialog', { name: 'Settings', exact: true })).toBeVisible();
	await page.getByRole('tab', { name: 'Appearance and accessibility', exact: true }).click();
	await page.getByRole('button', { name: 'Light', exact: true }).click();
	const rootBackground = await page.evaluate(() =>
		getComputedStyle(document.documentElement).getPropertyValue('--office-background'),
	);
	const light = await page
		.locator('teams-app')
		.evaluate((element) => getComputedStyle(element).backgroundColor);
	await page.getByRole('button', { name: 'Dark', exact: true }).click();
	await expect(page.locator('teams-app')).toHaveAttribute('data-office-theme', 'dark');
	const dark = await page
		.locator('teams-app')
		.evaluate((element) => getComputedStyle(element).backgroundColor);
	expect(dark).not.toBe(light);
	expect(
		await page.evaluate(() =>
			getComputedStyle(document.documentElement).getPropertyValue('--office-background'),
		),
	).toBe(rootBackground);
	await page.screenshot({ path: test.info().outputPath('settings-appearance.png') });
	await page.getByRole('tab', { name: 'Notifications and activity', exact: true }).click();
	await page.getByRole('checkbox', { name: 'Threads I start', exact: true }).uncheck();
	await page.getByRole('checkbox', { name: 'Threads I reply to', exact: true }).uncheck();
	await page.getByRole('button', { name: 'Close settings', exact: true }).click();
	await page.reload();
	await expect(page.getByRole('heading', { name: '# General', exact: true })).toBeVisible({
		timeout: 15_000,
	});
	await expect(page.locator('teams-app')).toHaveAttribute('data-office-theme', 'dark');
	await page.getByRole('button', { name: 'Settings', exact: true }).click();
	await page.getByRole('tab', { name: 'Notifications and activity', exact: true }).click();
	await expect(
		page.getByRole('checkbox', { name: 'Threads I start', exact: true }),
	).not.toBeChecked();
	await expect(
		page.getByRole('checkbox', { name: 'Threads I reply to', exact: true }),
	).not.toBeChecked();
	await page
		.getByRole('tab', { name: 'Notifications and activity', exact: true })
		.press('ArrowDown');
	await expect(page.getByRole('tab', { name: 'Files and links', exact: true })).toBeFocused();
	await expect(page.getByRole('heading', { name: 'Files and links', exact: true })).toBeVisible();
	await page.getByRole('button', { name: 'Close settings', exact: true }).click();
	await page.locator('teams-app').evaluate(async (element) => {
		const app = element as TeamsApp;
		app.client!.createChannel('Keep changes');
		await Promise.resolve();
		await app.client!.send({ text: 'Keep this when applying connection settings' });
	});
	await page.getByRole('button', { name: 'Settings', exact: true }).click();
	await page.getByRole('tab', { name: 'Connection', exact: true }).click();
	await page.getByRole('button', { name: 'Apply and reconnect', exact: true }).click();
	await page.getByRole('button', { name: 'Keep changes', exact: true }).click();
	await expect(
		page.getByText('Keep this when applying connection settings', { exact: true }),
	).toBeVisible();
	await page.getByRole('button', { name: 'Settings', exact: true }).click();
	await page.setViewportSize({ width: 390, height: 844 });
	await page.getByRole('tab', { name: 'Appearance and accessibility', exact: true }).click();
	await page.getByRole('button', { name: 'Follow system', exact: true }).click();
	await expect(page.locator('teams-app')).not.toHaveAttribute('data-office-theme');
	await expect(page.getByRole('button', { name: 'Close settings', exact: true })).toBeVisible();
	await page.screenshot({ path: test.info().outputPath('settings-mobile.png') });
	await page.getByRole('button', { name: 'Close settings', exact: true }).click();
	await expect(page.getByRole('dialog', { name: 'Settings', exact: true })).not.toBeVisible();
});
