import { readFile } from 'node:fs/promises';
import { expect, test } from '@playwright/test';
import type { TeamsApp } from '../../src/ui/src/teams/app/teams-app';

test('opens actual workbook bytes in a browser viewer, downloads and copies links with saved preferences', async ({
	page,
	context,
}) => {
	const bytes = await readFile(
		new URL('../../src/core/xlsx/__fixtures__/openpyxl-features.xlsx', import.meta.url),
	);
	await context.route('**/content/Budget.xlsx', (route) => route.fulfill({ body: bytes }));
	await page.goto(`/?local=1&name=Ada&room=file-actions-${Date.now()}`);
	await expect(page.getByRole('heading', { name: '# General', exact: true })).toBeVisible({
		timeout: 15_000,
	});
	await page.locator('teams-app').evaluate((element) => {
		const app = element as TeamsApp;
		app.client!.workspace.chat.post(app.client!.getState().selectedChannelId, {
			text: 'Review budget',
			attachments: [
				{ name: 'Budget.xlsx', kind: 'xlsx', url: `${location.origin}/content/Budget.xlsx` },
			],
		});
		Object.defineProperty(navigator, 'clipboard', {
			configurable: true,
			value: {
				writeText: async (text: string) => {
					(window as unknown as { copied: string }).copied = text;
				},
			},
		});
	});
	await page.getByRole('tab', { name: 'Shared', exact: true }).click();
	await page.getByRole('button', { name: 'More actions for Budget.xlsx', exact: true }).click();
	await page.screenshot({ path: test.info().outputPath('shared-file-actions.png') });
	await page.getByRole('button', { name: 'Copy link', exact: true }).click();
	await expect(page.getByRole('status')).toHaveText('Link copied');
	await page
		.locator('teams-app')
		.evaluate((element) => (element as TeamsApp).client!.setAvailability('busy'));
	await expect(page.getByRole('button', { name: 'Download', exact: true })).toBeVisible();
	expect(await page.evaluate(() => (window as unknown as { copied: string }).copied)).toBe(
		`${new URL(page.url()).origin}/content/Budget.xlsx`,
	);
	const downloadEvent = page.waitForEvent('download');
	await page.getByRole('button', { name: 'Download', exact: true }).click();
	const download = await downloadEvent;
	expect(download.suggestedFilename()).toBe('Budget.xlsx');
	expect(await readFile((await download.path())!)).toEqual(bytes);
	const browserEvent = context.waitForEvent('page');
	await page.getByRole('button', { name: 'Open in browser', exact: true }).click();
	const browser = await browserEvent;
	await browser.waitForLoadState();
	await expect(browser.getByRole('button', { name: 'Edit workbook', exact: true })).toBeVisible();
	expect(
		await browser.locator('teams-app').evaluate((element) => (element as TeamsApp).client),
	).toBeNull();
	expect(await browser.evaluate(() => window.opener)).toBeNull();
	await expect(browser.locator('xlsx-editor')).toBeVisible();
	await browser.screenshot({ path: test.info().outputPath('browser-workbook.png') });
	await browser.close();
	await page.getByRole('button', { name: 'Settings', exact: true }).click();
	await page.getByRole('tab', { name: 'Files and links', exact: true }).click();
	await page
		.getByRole('combobox', { name: 'Office file open preference', exact: true })
		.selectOption('browser');
	await page.screenshot({ path: test.info().outputPath('settings-files.png') });
	await page.getByRole('button', { name: 'Close settings', exact: true }).click();
	await page.reload();
	await expect(page.getByRole('heading', { name: '# General', exact: true })).toBeVisible({
		timeout: 15_000,
	});
	await page.getByRole('tab', { name: 'Shared', exact: true }).click();
	const preferredEvent = context.waitForEvent('page');
	await page.getByRole('button', { name: 'Open Budget.xlsx', exact: true }).click();
	const preferred = await preferredEvent;
	await expect(preferred.getByRole('button', { name: 'Edit workbook', exact: true })).toBeVisible();
	await preferred.close();
	await page.getByRole('button', { name: 'More actions for Budget.xlsx', exact: true }).click();
	await page.getByRole('button', { name: 'Open in OpenTeams', exact: true }).click();
	await expect(page.getByRole('button', { name: 'Edit workbook', exact: true })).toBeVisible();
});
