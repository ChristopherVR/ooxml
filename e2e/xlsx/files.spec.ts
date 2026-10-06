import { expect, test } from '@playwright/test';
import {
	editor,
	editorProperty,
	grid,
	landingFileInput,
	newWorkbook,
	openLanding,
	part,
	sheetTabs,
	typeInActiveCell,
} from './helpers';

test('a .csv file opens through the landing file input', async ({ page }) => {
	await openLanding(page);
	await (
		await landingFileInput(page)
	).setInputFiles({
		name: 'people.csv',
		mimeType: 'text/csv',
		buffer: Buffer.from('Name,Age\nAda,36\nGrace,45\n'),
	});
	await expect(grid(page)).toContainText('Grace');
	await expect(sheetTabs(page)).toContainText('people');
});

test('a downloaded .xlsx reopens with its edits (save round trip)', async ({ page }) => {
	await newWorkbook(page);
	await typeInActiveCell(page, 'Round trip');
	await typeInActiveCell(page, '=LEN(A1)');
	const download = page.waitForEvent('download');
	await editor(page).evaluate((node) =>
		(node as unknown as { download(name: string): Promise<void> }).download('trip.xlsx'),
	);
	const file = await download;
	expect(file.suggestedFilename()).toBe('trip.xlsx');
	const path = await file.path();
	await openLanding(page);
	await (await landingFileInput(page)).setInputFiles(path);
	await expect(grid(page)).toContainText('Round trip');
	await expect(grid(page)).toContainText('10');
	expect(await editorProperty<boolean>(page, 'dirty')).toBe(false);
});

test('the Editing / Viewing switch makes the editor read-only', async ({ page }) => {
	await newWorkbook(page);
	await part(page, 'title-bar')
		.getByRole('combobox', { name: 'Editing mode' })
		.selectOption('viewing');
	expect(await editorProperty<boolean>(page, 'readOnly')).toBe(true);
});
