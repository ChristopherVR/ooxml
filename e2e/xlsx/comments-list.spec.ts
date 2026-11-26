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
	return editor(page).locator('.xve-comments-pane');
}

for (const framework of FRAMEWORKS)
	test(`${framework}: Show Comments lists the threads and selects the chosen cell`, async ({
		page,
	}) => {
		const errors = pageErrors(page);
		const panel = await showComments(page, framework);
		const pane = panel.locator('office-ui-comments-pane');
		await expect(pane).toBeVisible();
		const threads = pane.locator('.thread');
		// Sheet order, keyed by address, with every author and text.
		await expect(threads).toHaveCount(1);
		await expect(threads.locator('.anchor-label')).toHaveText(['D5']);
		await expect(threads.first().locator('.author')).toHaveText(['Ana', 'Ben']);
		await expect(threads.first().locator('.text')).toHaveText([
			'Check the Q3 total.',
			'Fixed in the ledger.',
		]);
		// A navigation list: no composer, reply box or actions.
		await expect(pane.locator('.composer, .reply-box, .actions, textarea')).toHaveCount(0);

		await threads.first().click();
		await expect(panel).toBeVisible();
		await expect.poll(() => nameBoxValue(page)).toBe('D5');
		expect(errors).toEqual([]);
	});

for (const framework of FRAMEWORKS)
	test(`${framework}: header Comments toggles an empty pane and creates a saved thread`, async ({
		page,
	}) => {
		await newWorkbook(page, framework);
		const host = editor(page);
		const toggle = host
			.locator('office-ui-ribbon-actions')
			.getByRole('button', { name: 'Comments', exact: true });
		const pane = host.getByRole('complementary', { name: 'Comments', exact: true });
		await toggle.press('Enter');
		await expect(toggle).toHaveAttribute('aria-pressed', 'true');
		await expect(pane.locator('.empty')).toHaveText('There are no comments on this sheet.');
		const add = pane.getByRole('button', { name: 'New Comment', exact: true });
		await expect(add).toBeFocused();
		await page.keyboard.press('Escape');
		await expect(pane).toBeHidden();
		await expect(toggle).toBeFocused();
		await expect(toggle).toHaveAttribute('aria-pressed', 'false');
		await toggle.click();
		await toggle.click();
		await expect(pane).toBeHidden();
		await toggle.click();
		await add.click();
		const dialog = host.locator('[data-dialog="comment"]');
		await dialog.getByLabel('Comment', { exact: true }).fill('A new thread');
		await dialog.getByRole('button', { name: 'OK', exact: true }).click();
		await expect(pane.locator('.text')).toHaveText('A new thread');
		await host.evaluate(async (node) => {
			const control = node as unknown as {
				save(): Promise<Blob>;
				load(data: Uint8Array, name: string): Promise<void>;
			};
			const saved = await control.save();
			await control.load(new Uint8Array(await saved.arrayBuffer()), 'comments-reopened.xlsx');
		});
		await expect(pane.locator('.text')).toHaveText('A new thread');
	});
