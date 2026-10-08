import { expect, test } from '@playwright/test';
import type { Workbook } from 'ooxml-core/xlsx';
import {
	editor,
	editorProperty,
	FRAMEWORKS,
	grid,
	nameBox,
	newWorkbook,
	ribbon,
	typeInActiveCell,
} from './helpers';

interface SaveEditor extends HTMLElement {
	workbook: Workbook;
	saveBytes(format: 'xlsx' | 'csv'): Promise<Uint8Array>;
	download(): Promise<void>;
	focusGrid(): void;
}

for (const framework of FRAMEWORKS) {
	test(`${framework}: a rejected cell edit blocks exports and downloads`, async ({ page }) => {
		const downloads: string[] = [];
		page.on('download', (download) => downloads.push(download.suggestedFilename()));
		await newWorkbook(page, framework);
		await typeInActiveCell(page, '1');
		await nameBox(page).fill('A1');
		await nameBox(page).press('Enter');
		await ribbon(page).getByRole('tab', { name: 'Data', exact: true }).click();
		await ribbon(page).locator('[data-command="data.validation"]').first().click();
		const settings = editor(page).locator('[data-dialog="data-validation"]');
		await settings.getByLabel('Allow:', { exact: true }).selectOption('whole');
		await settings.getByLabel('Minimum:', { exact: true }).fill('1');
		await settings.getByLabel('Maximum:', { exact: true }).fill('2');
		await settings.getByRole('button', { name: 'OK', exact: true }).click();
		await expect(settings).toBeHidden();
		await editor(page).evaluate((node) => (node as SaveEditor).focusGrid());
		await page.keyboard.type('9');
		await expect(grid(page).getByRole('textbox')).toHaveValue('9');
		const exported = await editor(page).evaluate(async (node) => {
			const element = node as SaveEditor;
			const errors: string[] = [];
			for (const save of [
				() => element.saveBytes('xlsx'),
				() => element.saveBytes('csv'),
				() => element.download(),
			]) {
				try {
					await save();
					errors.push('saved');
				} catch (error) {
					errors.push(error instanceof Error ? error.message : String(error));
				}
			}
			return { errors, value: element.workbook.sheets[0]!.rows.get(0)?.get(0)?.value };
		});
		expect(exported.errors).toEqual(
			Array(3).fill('Finish or cancel the current cell edit before saving.'),
		);
		expect(exported.value).toBe(1);
		const alert = editor(page).getByRole('alertdialog');
		await expect(alert).toBeVisible();
		expect(await editorProperty(page, 'dirty')).toBe(true);
		expect(downloads).toEqual([]);
		await alert.getByRole('button', { name: 'Retry', exact: true }).click();
		await editor(page)
			.getByRole('toolbar', { name: 'Quick access', exact: true })
			.getByRole('button', { name: 'Save', exact: true })
			.click();
		await expect(alert).toBeVisible();
		expect(await editorProperty(page, 'dirty')).toBe(true);
		expect(downloads).toEqual([]);
		await alert.getByRole('button', { name: 'Cancel', exact: true }).click();
		const csv = await editor(page).evaluate(async (node) =>
			new TextDecoder().decode(await (node as SaveEditor).saveBytes('csv')).trim(),
		);
		expect(csv).toBe('1');
		const [download] = await Promise.all([
			page.waitForEvent('download'),
			editor(page).evaluate((node) => (node as SaveEditor).download()),
		]);
		expect(download.suggestedFilename()).toMatch(/\.xlsx$/);
		expect(await editorProperty(page, 'dirty')).toBe(false);
	});
}
