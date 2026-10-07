import { expect, test } from '@playwright/test';
import JSZip from 'jszip';
import type { DocxEditorElement } from '../../viewers/docx/packages/web-component/src';

for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid']) {
	test(`${framework}: Yjs review pane synchronizes threads, replies and local undo`, async ({
		page,
	}) => {
		const errors: string[] = [];
		page.on('pageerror', (error) => errors.push(error.message));
		await page.goto(`/collaboration.html?framework=${framework}&mode=yjs`);
		const a = page.locator('#peer-a docx-editor');
		const b = page.locator('#peer-b docx-editor');
		await expect(a.locator('.ProseMirror')).toContainText('Shared document');
		await expect(b.locator('.ProseMirror')).toContainText('Shared document');
		await a.locator('.ProseMirror').click();
		await page.keyboard.press('Control+Home');
		await page.keyboard.press('Control+Shift+ArrowRight');
		await a.getByRole('button', { name: 'Show comments', exact: true }).click();
		const paneA = a.locator('.dve-comments-panel');
		await paneA.getByRole('textbox', { name: 'New comment', exact: true }).fill('Shared review');
		await paneA.getByRole('button', { name: 'Add comment', exact: true }).click();
		await b.getByRole('button', { name: 'Show comments', exact: true }).click();
		const paneB = b.locator('.dve-comments-panel');
		await expect(paneB).toContainText('Shared review');
		await page.getByRole('button', { name: 'Pause delivery', exact: true }).click();
		for (const [pane, text] of [
			[paneA, 'Alice reply'],
			[paneB, 'Bob reply'],
		] as const) {
			await pane.getByRole('textbox', { name: 'Reply', exact: true }).fill(text);
			await pane.getByRole('button', { name: 'Reply', exact: true }).click();
		}
		await paneB.getByRole('button', { name: 'Resolve', exact: true }).click();
		await page.getByRole('button', { name: 'Resume delivery', exact: true }).click();
		for (const pane of [paneA, paneB]) {
			await expect(pane).toContainText('Alice reply');
			await expect(pane).toContainText('Bob reply');
			await expect(pane.getByRole('button', { name: 'Reopen', exact: true })).toBeVisible();
		}
		await a.locator('.ProseMirror').click();
		await page.keyboard.press('Control+z');
		for (const pane of [paneA, paneB]) {
			await expect(pane).not.toContainText('Alice reply');
			await expect(pane).toContainText('Bob reply');
		}
		const bytes = await b.evaluate(async (element) =>
			Array.from(await (element as DocxEditorElement).saveBytes()),
		);
		const zip = await JSZip.loadAsync(new Uint8Array(bytes));
		const comments = await zip.file('word/comments.xml')!.async('string');
		expect(comments).toContain('Shared review');
		expect(comments).toContain('Bob reply');
		expect(comments).not.toContain('Alice reply');
		await page.getByLabel('Editor B read only', { exact: true }).check();
		// Reopening refreshes the existing pane against the current write permission.
		await paneB.getByRole('button', { name: 'Close comments', exact: true }).click();
		await b.getByRole('button', { name: 'Show comments', exact: true }).click();
		await expect(paneB.getByRole('button', { name: 'Reopen', exact: true })).toBeDisabled();
		await expect(paneB.getByRole('button', { name: 'Reply', exact: true })).toBeDisabled();
		await paneA
			.locator('.dve-comment-thread')
			.filter({ hasText: 'Shared review' })
			.getByRole('button', { name: 'Delete', exact: true })
			.click();
		for (const pane of [paneA, paneB]) {
			await expect(pane).not.toContainText('Shared review');
			await expect(pane).not.toContainText('Bob reply');
		}
		await a.locator('.ProseMirror').click();
		await page.keyboard.press('Control+z');
		for (const pane of [paneA, paneB]) {
			await expect(pane).toContainText('Shared review');
			await expect(pane).toContainText('Bob reply');
		}
		expect(errors).toEqual([]);
	});
}
