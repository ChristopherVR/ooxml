import { fileURLToPath } from 'node:url';
import { expect, test } from '@playwright/test';
import type { Workbook } from 'ooxml-core/xlsx';
import { metadataFixture } from '../../src/core/xlsx/__fixtures__/threaded-comment-metadata';
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
	test(`${framework}: preserves an imported mention while editing and reopening its thread`, async ({
		page,
	}) => {
		const errors = pageErrors(page);
		await openLanding(page, framework);
		await (
			await landingFileInput(page)
		).setInputFiles({
			name: 'mentions.xlsx',
			mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
			buffer: Buffer.from(await metadataFixture()),
		});
		await expect(grid(page)).toBeVisible();
		await goToCell(page, 'B2');
		await page.keyboard.press('Shift+F2');
		const dialog = editor(page).locator('[data-dialog="comment"]');
		await expect(dialog.getByRole('checkbox', { name: 'Resolved', exact: true })).toBeChecked();
		await dialog.getByLabel('Comment', { exact: true }).fill('Before Hello @Parity Author! After');
		await dialog.getByRole('checkbox', { name: 'Resolved', exact: true }).uncheck();
		await dialog.getByRole('button', { name: 'OK', exact: true }).click();
		await editor(page).evaluate(async (node) => {
			const el = node as CommentEditor;
			await el.load(await el.saveBytes(), 'edited-mention.xlsx');
		});
		const result = await editor(page).evaluate((node) => {
			const comment = (node as CommentEditor).workbook.sheets[0]!.comments[0]!;
			const xml = new DOMParser().parseFromString(comment.threadedXml!, 'text/xml');
			const mention = xml.getElementsByTagNameNS('*', 'mention')[0]!;
			return {
				text: comment.text,
				resolved: comment.resolved,
				mention: mention.getAttribute('mentionId'),
				start: mention.getAttribute('startIndex'),
			};
		});
		expect(result).toEqual({
			text: 'Before Hello @Parity Author! After',
			resolved: false,
			mention: '{11111111-1111-4111-8111-111111111111}',
			start: '13',
		});
		expect(errors).toEqual([]);
	});

	test(`${framework}: edits a threaded reply, undoes it, and saves the thread`, async ({
		page,
	}) => {
		const errors = pageErrors(page);
		await openLanding(page, framework);
		await (await landingFileInput(page)).setInputFiles(fixture);
		await expect(grid(page)).toBeVisible();
		const comments = () =>
			editor(page).evaluate(
				(node) =>
					JSON.parse(
						JSON.stringify(
							(node as CommentEditor).workbook.sheets[0]!.comments,
							(key, value: unknown) => (key === 'threadedXml' ? undefined : value),
						),
					) as Workbook['sheets'][number]['comments'],
			);
		const original = await comments();
		await goToCell(page, 'B2');
		await page.keyboard.press('Shift+F2');
		const dialog = editor(page).locator('[data-dialog="comment"]');
		await expect(dialog.getByLabel('Comment', { exact: true })).toHaveValue('Original comment');
		await dialog.getByLabel('Reply 2 (Parity Author)', { exact: true }).fill('Edited second reply');
		await dialog.getByRole('checkbox', { name: 'Resolved', exact: true }).check();
		await dialog.getByRole('button', { name: 'OK', exact: true }).click();
		await expect(dialog).toBeHidden();
		const edited = structuredClone(original);
		edited[0]!.replies![1]!.text = 'Edited second reply';
		edited[0]!.resolved = true;
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
		await expect(dialog.getByRole('checkbox', { name: 'Resolved', exact: true })).toBeChecked();
		await dialog.getByRole('checkbox', { name: 'Resolved', exact: true }).uncheck();
		await dialog.getByRole('button', { name: 'OK', exact: true }).click();
		await editor(page).evaluate(async (node) => {
			const el = node as CommentEditor;
			await el.load(await el.saveBytes(), 'reopened-thread.xlsx');
		});
		expect((await comments())[0]!.resolved).toBe(false);
		expect(errors).toEqual([]);
	});
}
