import { expect, test } from '@playwright/test';
import type { TeamsApp } from '../../src/ui/src/teams/app/teams-app';

test('finds working settings and saves app labels per user and workspace', async ({ page }) => {
	const room = `settings-search-${Date.now()}`;
	await page.goto(`/?local=1&name=Ada&room=${room}`);
	await expect(page.getByRole('heading', { name: '# General', exact: true })).toBeVisible({
		timeout: 15_000,
	});
	const open = async () => {
		await page.getByRole('button', { name: 'Settings and more', exact: true }).click();
		await page.getByRole('menuitem', { name: 'Settings', exact: true }).click();
	};
	await open();
	const search = page.getByRole('searchbox', { name: 'Search settings', exact: true });
	await search.fill('APPEARANCE compact');
	await expect(page.getByText('1 setting found', { exact: true })).toBeVisible();
	await page.getByRole('button', { name: /^Chat density Appearance/ }).click();
	await expect(
		page.getByRole('tab', { name: 'Appearance and accessibility', exact: true }),
	).toHaveAttribute('aria-selected', 'true');
	await expect(page.locator('teams-settings').locator('legend')).toBeFocused();
	await expect(search).toHaveValue('');
	await search.fill('app bar icons');
	await page.getByRole('button', { name: /^Show app names Appearance/ }).press('Enter');
	await expect(page.getByRole('heading', { name: 'Show app names', exact: true })).toBeFocused();
	await page.getByRole('checkbox', { name: 'Show app names', exact: true }).uncheck();
	const rail = page.getByRole('navigation', { name: 'App bar', exact: true });
	await expect(rail).toHaveAttribute('hide-labels', '');
	for (const label of await rail.locator('.label').all()) await expect(label).toBeHidden();
	await page.screenshot({ path: test.info().outputPath('settings-app-names.png') });
	await search.fill('not-a-supported-setting');
	await expect(
		page.getByText('No matching settings. Try another keyword.', { exact: true }),
	).toBeVisible();
	await expect(page.getByRole('heading', { name: 'Theme', exact: true })).not.toBeVisible();
	await search.fill('ICE servers');
	await page.getByRole('button', { name: /^Connection Connection/ }).click();
	await expect(page.getByRole('heading', { name: 'Connection', exact: true })).toBeFocused();
	await page.getByRole('textbox', { name: /^Sync URL/ }).fill('ws://keep-draft.invalid/sync');
	await search.fill('browser');
	await page.getByRole('button', { name: /^File open preference Files/ }).click();
	await page.getByRole('tab', { name: 'Connection', exact: true }).click();
	await expect(page.getByRole('textbox', { name: /^Sync URL/ })).toHaveValue(
		'ws://keep-draft.invalid/sync',
	);
	await search.fill('theme');
	await page.screenshot({ path: test.info().outputPath('settings-search.png') });
	await page.getByRole('button', { name: 'Close settings', exact: true }).click();
	await expect(rail.getByRole('button', { name: 'Files', exact: true })).toHaveAttribute(
		'title',
		'Files',
	);
	await rail.getByRole('button', { name: 'Files', exact: true }).click();
	await expect(page.getByRole('heading', { name: 'Files', exact: true })).toBeVisible();
	await page.reload();
	await expect(page.getByRole('heading', { name: '# General', exact: true })).toBeVisible({
		timeout: 15_000,
	});
	await expect(rail).toHaveAttribute('hide-labels', '');
	await open();
	await expect(search).toHaveValue('');
	await page.setViewportSize({ width: 390, height: 844 });
	await search.fill('theme');
	const dialog = page.getByRole('dialog', { name: 'Settings', exact: true });
	const bounds = (await dialog.boundingBox())!;
	expect(bounds.x).toBeGreaterThanOrEqual(0);
	expect(bounds.x + bounds.width).toBeLessThanOrEqual(390);
	await page.screenshot({ path: test.info().outputPath('settings-search-mobile.png') });
	await page.getByRole('button', { name: /^Theme Appearance/ }).click();
	await expect(page.getByRole('heading', { name: 'Theme', exact: true })).toBeFocused();
	await page.getByRole('button', { name: 'Close settings', exact: true }).click();
	await page.locator('teams-app').evaluate(async (element, otherRoom) => {
		const app = element as TeamsApp;
		app.workspaceId = otherRoom;
		await app.updateComplete;
	}, `${room}-other`);
	await expect(rail).not.toHaveAttribute('hide-labels');
	await page.locator('teams-app').evaluate(async (element, originalRoom) => {
		const app = element as TeamsApp;
		app.workspaceId = originalRoom;
		await app.updateComplete;
	}, room);
	await expect(rail).toHaveAttribute('hide-labels', '');
	await page.locator('teams-app').evaluate(async (element) => {
		const app = element as TeamsApp;
		app.userId = 'another-user';
		await app.updateComplete;
	});
	await expect(rail).not.toHaveAttribute('hide-labels');
});
