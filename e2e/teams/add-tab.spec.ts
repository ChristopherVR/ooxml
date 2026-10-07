import { readFile } from 'node:fs/promises';
import { expect, test } from '@playwright/test';
import type { TeamsApp } from '../../src/ui/src/teams/app/teams-app';

test('configures tabs from shared workbooks, Markdown and sites while rejecting stale selections', async ({
	page,
}) => {
	const bytes = await readFile(
		new URL('../../src/core/xlsx/__fixtures__/openpyxl-features.xlsx', import.meta.url),
	);
	await page.route('**/content/Budget.xlsx', (route) => route.fulfill({ body: bytes }));
	await page.route('**/content/Notes.md', (route) =>
		route.fulfill({ body: '# Project notes\n\nShared decisions.', contentType: 'text/markdown' }),
	);
	await page.route('**/content/index.html', (route) =>
		route.fulfill({ body: '<h1>Project roadmap</h1>', contentType: 'text/html' }),
	);
	await page.goto(`/?local=1&name=Ada&room=add-tab-${Date.now()}`);
	await expect(page.getByRole('heading', { name: '# General', exact: true })).toBeVisible({
		timeout: 15_000,
	});
	const messageId = await page.locator('teams-app').evaluate((element) => {
		const client = (element as TeamsApp).client!;
		return client.workspace.chat.post(client.getState().selectedChannelId, {
			text: 'Shared resources',
			attachments: [
				{ name: 'Budget.xlsx', kind: 'xlsx', url: `${location.origin}/content/Budget.xlsx` },
				{ name: 'Missing.xlsx', kind: 'xlsx' },
				{ name: 'Notes.md', kind: 'other', url: `${location.origin}/content/Notes.md` },
				{ name: 'index.html', kind: 'other', url: `${location.origin}/content/index.html` },
			],
		})!.id;
	});
	await page.getByRole('button', { name: 'Add tab', exact: true }).click();
	const dialog = page.getByRole('dialog', { name: 'Add a tab', exact: true });
	await expect(dialog).toBeVisible();
	await expect(dialog.getByRole('button', { name: 'Choose Word', exact: true })).toBeFocused();
	await page.screenshot({ path: test.info().outputPath('add-tab-apps.png') });
	await dialog.getByRole('button', { name: 'Choose Word', exact: true }).click();
	await expect(dialog.getByText(/No Word files are shared/)).toBeVisible();
	await expect(dialog.getByRole('button', { name: 'Save', exact: true })).toBeDisabled();
	await dialog.getByRole('button', { name: 'Back', exact: true }).click();
	await expect(dialog.getByRole('button', { name: 'Choose Word', exact: true })).toBeFocused();
	await dialog.getByRole('button', { name: 'Choose Excel', exact: true }).click();
	await expect(dialog.getByRole('textbox', { name: 'Tab name', exact: true })).toBeFocused();
	await expect(dialog.getByRole('radio')).toHaveCount(1);
	await dialog.getByRole('radio', { name: /Budget.xlsx/ }).check();
	await page.locator('teams-app').evaluate((element, id) => {
		const client = (element as TeamsApp).client!;
		client.workspace.chat.remove(client.getState().selectedChannelId, id);
	}, messageId);
	await expect(dialog.getByRole('alert')).toHaveText(
		'This file is no longer shared in this channel.',
	);
	await expect(dialog.getByRole('button', { name: 'Save', exact: true })).toBeDisabled();
	await dialog.getByRole('button', { name: 'Cancel', exact: true }).click();
	await expect(page.getByRole('button', { name: 'Add tab', exact: true })).toBeFocused();
	await page.locator('teams-app').evaluate((element) => {
		const client = (element as TeamsApp).client!;
		client.workspace.chat.post(client.getState().selectedChannelId, {
			text: 'Current resources',
			attachments: [
				{ name: 'Budget.xlsx', kind: 'xlsx', url: `${location.origin}/content/Budget.xlsx` },
				{ name: 'Notes.md', kind: 'other', url: `${location.origin}/content/Notes.md` },
				{ name: 'index.html', kind: 'other', url: `${location.origin}/content/index.html` },
			],
		});
	});
	await page.getByRole('button', { name: 'Add tab', exact: true }).click();
	await dialog.getByRole('button', { name: 'Choose Excel', exact: true }).click();
	await dialog
		.getByRole('searchbox', { name: 'Search shared files', exact: true })
		.fill('not found');
	await expect(dialog.getByText('No matching files.', { exact: true })).toBeVisible();
	await dialog.getByRole('searchbox', { name: 'Search shared files', exact: true }).fill('budget');
	await dialog.getByRole('radio', { name: /Budget.xlsx/ }).check();
	await page.screenshot({ path: test.info().outputPath('add-tab-workbook.png') });
	await dialog.getByRole('button', { name: 'Save', exact: true }).click();
	await expect(page.getByRole('tab', { name: 'Budget.xlsx', exact: true })).toHaveAttribute(
		'aria-selected',
		'true',
	);
	await expect(page.getByRole('button', { name: 'Edit workbook', exact: true })).toBeVisible();
	await page.getByRole('button', { name: 'Edit workbook', exact: true }).click();
	await page
		.locator('xlsx-editor')
		.evaluate((element) => (element as HTMLElement & { select(ref: string): void }).select('A1'));
	const formula = page
		.locator('xlsx-editor')
		.getByRole('textbox', { name: 'Formula Bar', exact: true });
	await formula.fill('Keep this edit');
	await formula.press('Enter');
	await expect
		.poll(() =>
			page
				.locator('xlsx-editor')
				.evaluate((element) => (element as HTMLElement & { dirty: boolean }).dirty),
		)
		.toBe(true);
	await page.getByRole('button', { name: 'Add tab', exact: true }).click();
	await dialog.getByRole('button', { name: 'Choose Markdown', exact: true }).click();
	await dialog.getByRole('radio', { name: /Notes.md/ }).check();
	await dialog.getByRole('textbox', { name: 'Tab name', exact: true }).fill('Project notes');
	page.once('dialog', (prompt) => prompt.dismiss());
	await dialog.getByRole('button', { name: 'Save', exact: true }).click();
	await expect(dialog).toBeVisible();
	expect(
		await page
			.locator('teams-app')
			.evaluate((element) => (element as TeamsApp).client!.getState().tabs.length),
	).toBe(1);
	page.once('dialog', (prompt) => prompt.accept());
	await dialog.getByRole('button', { name: 'Save', exact: true }).click();
	await expect(page.getByRole('heading', { name: 'Project notes', exact: true })).toBeVisible();
	await page.getByRole('button', { name: 'Add tab', exact: true }).click();
	await dialog.getByRole('button', { name: 'Choose Static site', exact: true }).click();
	await dialog.getByRole('radio', { name: /index.html/ }).check();
	await dialog.getByRole('button', { name: 'Save', exact: true }).click();
	await expect(
		page.frameLocator('iframe').getByRole('heading', { name: 'Project roadmap', exact: true }),
	).toBeVisible();
	await page.getByRole('button', { name: 'Add tab', exact: true }).click();
	await dialog.getByRole('button', { name: 'Choose Website', exact: true }).click();
	await dialog.getByRole('textbox', { name: 'Tab name', exact: true }).fill('Changed channel');
	await dialog
		.getByRole('textbox', { name: 'Tab website URL', exact: true })
		.fill('javascript:alert(1)');
	await dialog.getByRole('button', { name: 'Save', exact: true }).click();
	await expect(dialog.getByRole('alert')).toHaveText('Enter a valid http or https website URL.');
	await dialog
		.getByRole('textbox', { name: 'Tab website URL', exact: true })
		.fill('https://example.com');
	await page
		.locator('teams-app')
		.evaluate((element) => (element as TeamsApp).client!.createChannel('Different'));
	await expect(dialog.getByRole('alert')).toHaveText(
		'The channel changed. Close this dialog and try again.',
	);
	await expect(dialog.getByRole('button', { name: 'Save', exact: true })).toBeDisabled();
	await dialog.getByRole('button', { name: 'Cancel', exact: true }).click();
	await page.getByRole('button', { name: 'Add tab', exact: true }).click();
	await page.setViewportSize({ width: 390, height: 844 });
	const bounds = await dialog.boundingBox();
	expect(bounds!.x).toBeGreaterThanOrEqual(0);
	expect(bounds!.x + bounds!.width).toBeLessThanOrEqual(390);
	await page.screenshot({ path: test.info().outputPath('add-tab-mobile.png') });
	await dialog.getByRole('button', { name: 'Choose Word', exact: true }).press('Escape');
	await expect(dialog).not.toBeVisible();
	expect(
		await page
			.locator('teams-app')
			.evaluate((element) => (element as TeamsApp).client!.getState().tabs),
	).toEqual([]);
});
