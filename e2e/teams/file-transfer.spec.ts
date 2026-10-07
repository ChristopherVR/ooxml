import { expect, test } from '@playwright/test';
import type { TeamsApp } from '../../src/ui/src/teams/app/teams-app';

test('cancels pending uploads, ignores late completion and retries with file progress', async ({
	page,
}) => {
	await page.goto('/?local=1&name=Ada&room=transfer-start');
	await page.locator('teams-app').evaluate((element) => {
		const app = element as TeamsApp;
		const state = window as unknown as {
			transfers: { signal?: AbortSignal; finish: () => void }[];
		};
		state.transfers = [];
		app.uploadFile = (file, context) =>
			new Promise((resolve) => {
				state.transfers.push({
					signal: context.signal,
					finish: () => resolve({ url: `${location.origin}/content/${file.name}` }),
				});
			});
		app.workspaceId = `transfer-${Date.now()}`;
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
		.evaluate((element) => (element as TeamsApp).client!.createChannel('Transfers'));
	await expect(page.getByRole('heading', { name: '# Transfers' })).toBeVisible();
	await page.getByRole('tab', { name: 'Files', exact: true }).click();
	await page.getByLabel('Upload files', { exact: true }).setInputFiles([
		{ name: 'Budget.xlsx', mimeType: 'application/octet-stream', buffer: Buffer.from('workbook') },
		{ name: 'Notes.md', mimeType: 'text/markdown', buffer: Buffer.from('# Notes') },
	]);
	await expect(page.getByRole('status')).toContainText('0 of 2 uploaded');
	await page.getByRole('button', { name: 'Cancel sharing', exact: true }).click();
	await expect(page.getByRole('status')).toContainText('canceled');
	expect(
		await page.evaluate(() => {
			const transfers = (
				window as unknown as { transfers: { signal?: AbortSignal; finish: () => void }[] }
			).transfers;
			transfers[0]!.finish();
			return transfers[0]!.signal?.aborted;
		}),
	).toBe(true);
	await expect(page.getByRole('button', { name: 'Open', exact: true })).toHaveCount(0);
	await page.getByRole('button', { name: 'Retry sharing files', exact: true }).click();
	await expect
		.poll(() =>
			page.evaluate(() => (window as unknown as { transfers: unknown[] }).transfers.length),
		)
		.toBe(2);
	await page.evaluate(() =>
		(window as unknown as { transfers: { finish: () => void }[] }).transfers[1]!.finish(),
	);
	await expect(page.getByRole('status')).toContainText('Notes.md (1 of 2 uploaded)');
	await expect(page.getByRole('progressbar', { name: 'Files uploaded' })).toHaveAttribute(
		'value',
		'1',
	);
	await expect
		.poll(() =>
			page.evaluate(() => (window as unknown as { transfers: unknown[] }).transfers.length),
		)
		.toBe(3);
	await page.evaluate(() =>
		(window as unknown as { transfers: { finish: () => void }[] }).transfers[2]!.finish(),
	);
	await expect(page.getByRole('button', { name: 'Open', exact: true })).toHaveCount(2);
	await page.getByRole('button', { name: 'New Excel workbook', exact: true }).click();
	await page.getByRole('textbox', { name: 'Workbook name', exact: true }).fill('Canceled budget');
	await page.getByRole('button', { name: 'Create workbook', exact: true }).click();
	await expect
		.poll(() =>
			page.evaluate(() => (window as unknown as { transfers: unknown[] }).transfers.length),
		)
		.toBe(4);
	await page.getByRole('tab', { name: 'Posts', exact: true }).click();
	expect(
		await page.evaluate(() => {
			const pending = (
				window as unknown as { transfers: { signal?: AbortSignal; finish: () => void }[] }
			).transfers[3]!;
			pending.finish();
			return pending.signal?.aborted;
		}),
	).toBe(true);
	await page.getByRole('tab', { name: 'Files', exact: true }).click();
	await expect(page.getByRole('button', { name: 'Open', exact: true })).toHaveCount(2);
	await expect(page.getByRole('button', { name: 'Close preview', exact: true })).toHaveCount(0);
});
