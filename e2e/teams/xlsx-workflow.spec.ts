import { readFile } from 'node:fs/promises';
import { expect, test } from '@playwright/test';
import { getCell, loadXlsx } from '../../src/core/xlsx/index';
import type { TeamsApp } from '../../src/ui/src/teams/app/teams-app';

test('pins a workbook, edits a cell, saves a new channel copy, and opens the saved bytes', async ({
	page,
}) => {
	test.setTimeout(120_000);
	const original = await readFile(
		new URL('../../src/core/xlsx/__fixtures__/openpyxl-features.xlsx', import.meta.url),
	);
	await page.route('**/content/Budget.xlsx', (route) => route.fulfill({ body: original }));
	await page.route('**/content/copy-*.xlsx', async (route) => {
		const copies = await page.evaluate(
			() => (window as unknown as { copies: Record<string, number[]> }).copies,
		);
		const name = new URL(route.request().url()).pathname.split('/').pop()!;
		await route.fulfill({ body: Buffer.from(copies[name] ?? []) });
	});
	await page.goto('/?local=1&name=Ada&room=xlsx-source');
	await page.locator('teams-app').evaluate((element) => {
		const app = element as TeamsApp;
		const copies: Record<string, number[]> = {};
		(window as unknown as { copies: Record<string, number[]> }).copies = copies;
		app.uploadFile = async (file) => {
			const control = window as unknown as {
				failNext?: boolean;
				pauseNext?: boolean;
				releaseUpload?: () => void;
			};
			if (control.failNext) {
				control.failNext = false;
				throw new Error('Storage unavailable');
			}
			if (control.pauseNext) {
				control.pauseNext = false;
				await new Promise<void>((resolve) => (control.releaseUpload = resolve));
			}
			copies[file.name] = Array.from(new Uint8Array(await (file as File).arrayBuffer()));
			return { url: `${location.origin}/content/${file.name}` };
		};
		app.workspaceId = 'xlsx-copies';
	});
	await expect
		.poll(() =>
			page
				.locator('teams-app')
				.evaluate((element) => (element as TeamsApp).client?.workspace.session.roomId),
		)
		.toBe('xlsx-copies');
	await page.locator('teams-app').evaluate((element) => {
		const client = (element as TeamsApp).client!;
		client.createChannel('Finance');
	});
	await expect(page.getByRole('heading', { name: '# Finance' })).toBeVisible();
	await page.locator('teams-app').evaluate((element) => {
		const client = (element as TeamsApp).client!;
		client.workspace.chat.post(client.getState().selectedChannelId, {
			text: 'Review budget',
			attachments: [
				{ name: 'Budget.xlsx', kind: 'xlsx', url: `${location.origin}/content/Budget.xlsx` },
			],
		});
	});
	await page.getByRole('tab', { name: 'Shared', exact: true }).click();
	await page.getByRole('button', { name: 'More actions for Budget.xlsx', exact: true }).click();
	await page.getByRole('button', { name: 'Pin as tab' }).click();
	await expect(page.getByRole('tab', { name: 'Budget.xlsx' })).toHaveAttribute(
		'aria-selected',
		'true',
	);
	await expect(page.getByRole('button', { name: 'Edit workbook', exact: true })).toBeVisible({
		timeout: 60_000,
	});
	await page.getByRole('button', { name: 'Edit workbook', exact: true }).click();
	await page
		.locator('xlsx-editor')
		.evaluate((el) => (el as HTMLElement & { select(ref: string): void }).select('A1'));
	const formula = page
		.locator('xlsx-editor')
		.getByRole('textbox', { name: 'Formula Bar', exact: true });
	await formula.click();
	await formula.press('ControlOrMeta+A');
	await formula.pressSequentially('Teams workbook edit');
	await formula.press('Enter');
	await expect
		.poll(() =>
			page.locator('xlsx-editor').evaluate((el) => (el as HTMLElement & { dirty: boolean }).dirty),
		)
		.toBe(true);
	page.once('dialog', (dialog) => dialog.dismiss());
	await page.getByRole('tab', { name: 'Posts', exact: true }).click();
	await expect(page.locator('xlsx-editor')).toBeVisible();
	await page.getByRole('button', { name: 'Save copy to channel', exact: true }).click();
	await expect(page.getByRole('status').filter({ hasText: 'Workbook copy shared' })).toBeVisible();
	await expect
		.poll(() =>
			page.locator('xlsx-editor').evaluate((el) => (el as HTMLElement & { dirty: boolean }).dirty),
		)
		.toBe(false);
	const copies = await page.evaluate(
		() => (window as unknown as { copies: Record<string, number[]> }).copies,
	);
	expect(Object.keys(copies)).toHaveLength(1);
	const saved = await loadXlsx(Uint8Array.from(Object.values(copies)[0]!));
	expect(getCell(saved.sheets[0]!, 0, 0)?.value).toBe('Teams workbook edit');
	expect(getCell((await loadXlsx(original)).sheets[0]!, 0, 0)?.value).not.toBe(
		'Teams workbook edit',
	);
	await page.getByRole('tab', { name: 'Shared', exact: true }).click();
	await page
		.getByRole('button', { name: /^Open / })
		.first()
		.click();
	await expect
		.poll(
			() =>
				page.locator('xlsx-editor').evaluate(
					(el) =>
						(
							el as HTMLElement & {
								workbook: { sheets: { rows: Map<number, Map<number, { value: unknown }>> }[] };
							}
						).workbook?.sheets[0]?.rows
							.get(0)
							?.get(0)?.value,
				),
			{ timeout: 60_000 },
		)
		.toBe('Teams workbook edit');
	await page.getByRole('button', { name: 'Edit workbook', exact: true }).click();
	await page
		.locator('xlsx-editor')
		.evaluate((el) => (el as HTMLElement & { select(ref: string): void }).select('A1'));
	await formula.click();
	await formula.press('ControlOrMeta+A');
	await formula.pressSequentially('Retry workbook edit');
	await formula.press('Enter');
	await page.evaluate(() => ((window as unknown as { failNext: boolean }).failNext = true));
	await page.getByRole('button', { name: 'Save copy to channel', exact: true }).click();
	await expect(page.getByRole('alert')).toContainText('could not be uploaded');
	await expect
		.poll(() =>
			page.locator('xlsx-editor').evaluate((el) => (el as HTMLElement & { dirty: boolean }).dirty),
		)
		.toBe(true);
	await page.evaluate(() => ((window as unknown as { pauseNext: boolean }).pauseNext = true));
	await page.getByRole('button', { name: 'Save copy to channel', exact: true }).click();
	await expect
		.poll(() =>
			page.evaluate(
				() => typeof (window as unknown as { releaseUpload?: () => void }).releaseUpload,
			),
		)
		.toBe('function');
	await page.getByRole('button', { name: 'Close preview', exact: true }).click();
	await expect(page.getByRole('alert')).toContainText('Wait for the workbook copy');
	await formula.click();
	await formula.press('ControlOrMeta+A');
	await formula.pressSequentially('Edit during upload');
	await formula.press('Enter');
	await page.evaluate(() => (window as unknown as { releaseUpload: () => void }).releaseUpload());
	await expect(page.getByRole('status').filter({ hasText: 'Workbook copy shared' })).toBeVisible();
	await expect
		.poll(() =>
			page.locator('xlsx-editor').evaluate((el) => (el as HTMLElement & { dirty: boolean }).dirty),
		)
		.toBe(true);
	const latest = await page.evaluate(
		() => (window as unknown as { copies: Record<string, number[]> }).copies,
	);
	expect(Object.keys(latest)).toHaveLength(2);
	const pendingCopy = await loadXlsx(Uint8Array.from(Object.values(latest)[1]!));
	expect(getCell(pendingCopy.sheets[0]!, 0, 0)?.value).toBe('Retry workbook edit');
	const sharedBeforeCancel = await page
		.locator('teams-app')
		.evaluate((el) => (el as TeamsApp).client!.getState().files.length);
	await page.evaluate(() => {
		const control = window as unknown as { pauseNext: boolean; releaseUpload?: () => void };
		control.pauseNext = true;
		delete control.releaseUpload;
	});
	await page.getByRole('button', { name: 'Save copy to channel', exact: true }).click();
	await expect
		.poll(() =>
			page.evaluate(
				() => typeof (window as unknown as { releaseUpload?: () => void }).releaseUpload,
			),
		)
		.toBe('function');
	await page.getByRole('button', { name: 'Cancel workbook save', exact: true }).click();
	await expect(page.getByRole('alert')).toContainText('Your edits remain local');
	await expect(
		page.getByRole('button', { name: 'Save copy to channel', exact: true }),
	).toBeEnabled();
	await page.evaluate(() => (window as unknown as { releaseUpload: () => void }).releaseUpload());
	await expect
		.poll(() =>
			page.evaluate(
				() =>
					Object.keys((window as unknown as { copies: Record<string, number[]> }).copies).length,
			),
		)
		.toBe(3);
	expect(
		await page
			.locator('teams-app')
			.evaluate((el) => (el as TeamsApp).client!.getState().files.length),
	).toBe(sharedBeforeCancel);
	expect(
		await page
			.locator('xlsx-editor')
			.evaluate((el) => (el as HTMLElement & { dirty: boolean }).dirty),
	).toBe(true);
	await page.getByRole('button', { name: 'Save copy to channel', exact: true }).click();
	await expect(page.getByRole('status').filter({ hasText: 'Workbook copy shared' })).toBeVisible();
	await expect
		.poll(() =>
			page.locator('teams-app').evaluate((el) => (el as TeamsApp).client!.getState().files.length),
		)
		.toBe(sharedBeforeCancel + 1);
	expect(
		await page
			.locator('xlsx-editor')
			.evaluate((el) => (el as HTMLElement & { dirty: boolean }).dirty),
	).toBe(false);
});
