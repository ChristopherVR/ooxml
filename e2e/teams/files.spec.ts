import { readFile } from 'node:fs/promises';
import { expect, test } from '@playwright/test';
import type { TeamsApp } from '../../src/ui/src/teams/app/teams-app';

test('creates a workbook and uploads, retries and searches channel files', async ({ page }) => {
	test.setTimeout(90_000);
	await page.route('**/content/*', async (route) => {
		const name = new URL(route.request().url()).pathname.split('/').pop()!;
		const bytes = await page.evaluate(
			(name) => (window as unknown as { copies: Record<string, number[]> }).copies[name],
			name,
		);
		await route.fulfill({ body: Buffer.from(bytes ?? []) });
	});
	await page.goto('/?local=1&name=Ada&room=files-start');
	await page.locator('teams-app').evaluate((element) => {
		const app = element as TeamsApp;
		const storage = window as unknown as { copies: Record<string, number[]>; failNext?: boolean };
		storage.copies = {};
		app.uploadFile = async (file) => {
			if (storage.failNext) {
				storage.failNext = false;
				throw new Error('Storage offline');
			}
			storage.copies[file.name] = Array.from(new Uint8Array(await (file as File).arrayBuffer()));
			return { url: `${location.origin}/content/${file.name}` };
		};
		app.workspaceId = `files-${Date.now()}`;
	});
	await expect
		.poll(() =>
			page
				.locator('teams-app')
				.evaluate((el) => (el as TeamsApp).client?.getState().canUploadFiles),
		)
		.toBe(true);
	await page
		.locator('teams-app')
		.evaluate((element) => (element as TeamsApp).client!.createChannel('Finance'));
	await expect(page.getByRole('heading', { name: '# Finance' })).toBeVisible();
	await page.getByRole('tab', { name: 'Files', exact: true }).click();
	await page.getByRole('button', { name: 'New Excel workbook', exact: true }).click();
	await page.getByRole('textbox', { name: 'Workbook name', exact: true }).fill('Budget');
	await page.getByRole('button', { name: 'Create workbook', exact: true }).click();
	await expect(page.getByRole('button', { name: 'Edit workbook', exact: true })).toBeVisible({
		timeout: 60_000,
	});
	await expect
		.poll(() =>
			page
				.locator('xlsx-editor')
				.evaluate(
					(el) =>
						(el as HTMLElement & { workbook?: { sheets: { name: string }[] } }).workbook?.sheets[0]
							?.name,
				),
		)
		.toBe('Sheet1');
	await page.getByRole('button', { name: 'Close preview', exact: true }).click();
	const original = await readFile(
		new URL('../../src/core/xlsx/__fixtures__/openpyxl-features.xlsx', import.meta.url),
	);
	await page.evaluate(() => ((window as unknown as { failNext: boolean }).failNext = true));
	await page.getByLabel('Upload files', { exact: true }).setInputFiles([
		{
			name: 'Budget.xlsx',
			mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
			buffer: original,
		},
		{
			name: 'Notes.md',
			mimeType: 'text/markdown',
			buffer: Buffer.from('# Project notes\nUpload review'),
		},
	]);
	await expect(page.getByRole('alert')).toContainText('could not be uploaded');
	await page.getByRole('button', { name: 'Retry sharing files', exact: true }).click();
	await expect(page.getByRole('button', { name: 'Open', exact: true })).toHaveCount(3);
	await page.getByRole('searchbox', { name: 'Search files', exact: true }).fill('notes ada');
	await expect(page.getByRole('button', { name: 'Open', exact: true })).toHaveCount(1);
	await page.getByRole('button', { name: 'Open', exact: true }).click();
	await expect(page.getByRole('heading', { name: 'Project notes', exact: true })).toBeVisible();
	await page.getByRole('button', { name: 'Close preview', exact: true }).click();
	await page.getByRole('searchbox', { name: 'Search files', exact: true }).fill('budget');
	await expect(page.getByRole('button', { name: 'Open', exact: true })).toHaveCount(2);
	const budgets = await page.locator('teams-app').evaluate((element) =>
		(element as TeamsApp)
			.client!.getState()
			.files.filter((file) => file.name === 'Budget.xlsx')
			.map((file) => file.url),
	);
	expect(new Set(budgets).size).toBe(2);
	await page.getByRole('searchbox', { name: 'Search files', exact: true }).fill('absent');
	await expect(page.getByText('No matching files.', { exact: true })).toBeVisible();
});
