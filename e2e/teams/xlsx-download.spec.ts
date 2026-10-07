import { readFile } from 'node:fs/promises';
import { expect, test } from '@playwright/test';
import { getCell, loadXlsx } from '../../src/core/xlsx/index';
import type { TeamsApp } from '../../src/ui/src/teams/app/teams-app';
import type { Workbook } from '../../src/core/xlsx/model';

test('edits and downloads an XLSX copy without storage or publishing another attachment', async ({
	page,
}) => {
	const original = await readFile(
		new URL('../../src/core/xlsx/__fixtures__/openpyxl-features.xlsx', import.meta.url),
	);
	await page.route('**/content/Local.xlsx', (route) => route.fulfill({ body: original }));
	await page.goto('/?local=1&name=Ada&room=xlsx-download');
	await expect(page.getByRole('heading', { name: '# General', exact: true })).toBeVisible({
		timeout: 15_000,
	});
	await page.locator('teams-app').evaluate(async (element) => {
		const app = element as TeamsApp;
		await app.updateComplete;
		app.previewContent({
			attachment: { name: 'Local.xlsx', kind: 'xlsx' },
			url: `${location.origin}/content/Local.xlsx`,
		});
	});
	await expect(page.getByRole('button', { name: 'Edit workbook', exact: true })).toBeVisible();
	await expect(page.getByRole('button', { name: 'Save copy to channel', exact: true })).toHaveCount(
		0,
	);
	await page.getByRole('button', { name: 'Edit workbook', exact: true }).click();
	await page
		.locator('xlsx-editor')
		.evaluate((element) => (element as HTMLElement & { select(ref: string): void }).select('A1'));
	const formula = page
		.locator('xlsx-editor')
		.getByRole('textbox', { name: 'Formula Bar', exact: true });
	await page.locator('xlsx-editor').evaluate((element) => {
		(element as HTMLElement & { workbook: Workbook }).workbook.sheets[0]!.dataValidations.push({
			ranges: [{ start: { row: 0, col: 0 }, end: { row: 0, col: 0 } }],
			type: 'whole',
			operator: 'between',
			formula1: '1',
			formula2: '10',
			errorStyle: 'stop',
			showErrorMessage: true,
			error: 'Between 1 and 10',
		});
	});
	let downloads = 0;
	page.on('download', () => downloads++);
	await formula.click();
	await formula.press('ControlOrMeta+A');
	await formula.pressSequentially('50');
	await page.getByRole('button', { name: 'Download workbook copy', exact: true }).click();
	await expect(
		page
			.getByRole('alert')
			.filter({ hasText: 'Finish or cancel the current cell edit before saving.' }),
	).toBeVisible();
	await expect(page.getByRole('alertdialog')).toContainText('Between 1 and 10');
	expect(downloads).toBe(0);
	await page.getByRole('alertdialog').getByRole('button', { name: 'Cancel', exact: true }).click();
	await page.locator('xlsx-editor').evaluate((element) => {
		(element as HTMLElement & { workbook: Workbook }).workbook.sheets[0]!.dataValidations.pop();
	});
	await formula.click();
	await formula.press('ControlOrMeta+A');
	await formula.pressSequentially('Downloaded from OpenTeams');
	// Export commits the pending cell edit rather than silently saving its previous value.
	const waiting = page.waitForEvent('download');
	await page.getByRole('button', { name: 'Download workbook copy', exact: true }).click();
	const download = await waiting;
	expect(download.suggestedFilename()).toBe('Local.xlsx');
	const workbook = await loadXlsx(await readFile((await download.path())!));
	expect(getCell(workbook.sheets[0]!, 0, 0)?.value).toBe('Downloaded from OpenTeams');
	expect(getCell((await loadXlsx(original)).sheets[0]!, 0, 0)?.value).not.toBe(
		'Downloaded from OpenTeams',
	);
	await expect(
		page.getByRole('status').filter({ hasText: 'Workbook download started' }),
	).toBeVisible();
	expect(
		await page
			.locator('teams-app')
			.evaluate((element) => (element as TeamsApp).client!.getState().files.length),
	).toBe(0);
	expect(
		await page
			.locator('xlsx-editor')
			.evaluate((element) => (element as HTMLElement & { dirty: boolean }).dirty),
	).toBe(true);
	page.once('dialog', (dialog) => dialog.dismiss());
	await page.getByRole('button', { name: 'Close preview', exact: true }).click();
	await expect(page.locator('xlsx-editor')).toBeVisible();
	page.once('dialog', (dialog) => dialog.accept());
	await page.getByRole('button', { name: 'Close preview', exact: true }).click();
	await expect(page.locator('xlsx-editor')).toHaveCount(0);
});
