import { fileURLToPath } from 'node:url';
import { expect, test } from '@playwright/test';
import type { Workbook } from 'ooxml-core/xlsx';
import {
	editor,
	FRAMEWORKS,
	goToCell,
	grid,
	landingFileInput,
	openLanding,
	pageErrors,
} from './helpers';

interface CommentEditor extends HTMLElement {
	workbook: Workbook;
	saveBytes(): Promise<Uint8Array>;
	load(bytes: Uint8Array, name: string): Promise<void>;
}
const fixture = fileURLToPath(
	new URL('../../src/core/xlsx/__fixtures__/excel-threaded-comments.xlsx', import.meta.url),
);

for (const framework of FRAMEWORKS) {
	test(`${framework}: edits a threaded reply, undoes it, and saves the thread`, async ({
		page,
	}) => {
		const errors = pageErrors(page);
		await openLanding(page, framework);
		await (await landingFileInput(page)).setInputFiles(fixture);
		await expect(grid(page)).toBeVisible();
		const comments = () =>
			editor(page).evaluate((node) => (node as CommentEditor).workbook.sheets[0]!.comments);
		const original = await comments();
		await goToCell(page, 'B2');
		await page.keyboard.press('Shift+F2');
		const dialog = editor(page).locator('[data-dialog="comment"]');
		await expect(dialog.getByLabel('Comment', { exact: true })).toHaveValue('Original comment');
		await dialog.getByLabel('Reply 2 (Parity Author)', { exact: true }).fill('Edited second reply');
		await dialog.getByRole('button', { name: 'OK', exact: true }).click();
		await expect(dialog).toBeHidden();
		const edited = structuredClone(original);
		edited[0]!.replies![1]!.text = 'Edited second reply';
		expect(await comments()).toEqual(edited);
		await page.keyboard.press('Control+z');
		await expect.poll(comments).toEqual(original);
		await page.keyboard.press('Control+y');
		await expect.poll(comments).toEqual(edited);
		await editor(page).evaluate(async (node) => {
			const el = node as CommentEditor;
			await el.load(await el.saveBytes(), 'edited-comments.xlsx');
		});
		expect(await comments()).toEqual(edited);
		await goToCell(page, 'B2');
		await page.keyboard.press('Shift+F2');
		await expect(dialog.getByLabel('Reply 2 (Parity Author)', { exact: true })).toHaveValue(
			'Edited second reply',
		);
		await dialog.getByRole('button', { name: 'Cancel', exact: true }).click();
		expect(errors).toEqual([]);
	});
}
