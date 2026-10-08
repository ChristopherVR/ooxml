import { expect, test, type Page } from '@playwright/test';
import { createEditSession, createWorkbook, saveXlsx } from 'ooxml-core/xlsx';
import {
	FRAMEWORKS,
	editor,
	grid,
	nameBoxValue,
	newWorkbook,
	openLanding,
	pageErrors,
	ribbon,
} from './helpers';

/** A workbook with a note on B2 and a threaded comment with a reply on D5. */
async function commentedWorkbook(): Promise<Buffer> {
	const book = createWorkbook();
	const session = createEditSession(book);
	session.setComment(0, { row: 4, col: 3 }, 'Check the Q3 total.', 'Ana');
	session.setComment(0, { row: 1, col: 1 }, 'Source: ledger export.', 'Ben');
	book.sheets[0]!.comments.find((c) => c.address.row === 4)!.replies = [
		{ author: 'Ben', text: 'Fixed in the ledger.' },
	];
	return Buffer.from(await saveXlsx(book));
}

async function showComments(page: Page, framework: string) {
	await openLanding(page, framework);
	await page.locator('#landing-file').setInputFiles({
		name: 'comments.xlsx',
		mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
		buffer: await commentedWorkbook(),
	});
	await expect(grid(page)).toBeVisible();
	await ribbon(page).getByRole('tab', { name: 'Review', exact: true }).click();
	await ribbon(page).getByRole('button', { name: 'Show Comments', exact: true }).click();
	return editor(page).locator('[data-dialog="comments-list"]');
}

for (const framework of FRAMEWORKS)
	test(`${framework}: Show Comments lists the threads and selects the chosen cell`, async ({
		page,
	}) => {
		const errors = pageErrors(page);
		const dialog = await showComments(page, framework);
		const pane = dialog.locator('office-ui-comments-pane');
		await expect(pane).toBeVisible();
		const threads = pane.locator('.thread');
		// Sheet order, keyed by address, with every author and text.
		await expect(threads).toHaveCount(2);
		await expect(threads.locator('.anchor-label')).toHaveText(['B2', 'D5']);
		await expect(threads.nth(0).locator('.author')).toHaveText(['Ben']);
		await expect(threads.nth(0).locator('.text')).toHaveText(['Source: ledger export.']);
		await expect(threads.nth(1).locator('.author')).toHaveText(['Ana', 'Ben']);
		await expect(threads.nth(1).locator('.text')).toHaveText([
			'Check the Q3 total.',
			'Fixed in the ledger.',
		]);
		// A navigation list: no composer, reply box or actions.
		await expect(pane.locator('.composer, .reply-box, .actions, textarea')).toHaveCount(0);

		await threads.nth(1).click();
		await expect(dialog).toBeHidden();
		await expect.poll(() => nameBoxValue(page)).toBe('D5');
		expect(errors).toEqual([]);
	});

test('Show Comments on a sheet without comments says so', async ({ page }) => {
	await newWorkbook(page);
	await ribbon(page).getByRole('tab', { name: 'Review', exact: true }).click();
	await ribbon(page).getByRole('button', { name: 'Show Comments', exact: true }).click();
	const dialog = editor(page).locator('[data-dialog="comments-list"]');
	await expect(dialog.locator('.empty')).toHaveText('There are no comments on this sheet.');
});
