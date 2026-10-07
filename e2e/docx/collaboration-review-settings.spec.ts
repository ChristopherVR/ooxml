import { expect, test } from '@playwright/test';
import JSZip from 'jszip';
import type { DocxEditorElement } from '../../viewers/docx/packages/web-component/src';

for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid']) {
	test(`${framework}: rejecting a peer revision preserves the other author's insertion`, async ({
		page,
	}) => {
		const errors: string[] = [];
		page.on('pageerror', (error) => errors.push(error.message));
		await page.setViewportSize({ width: 3200, height: 1000 });
		await page.goto(`/collaboration.html?framework=${framework}&mode=yjs`);
		const a = page.locator('#peer-a docx-editor');
		const b = page.locator('#peer-b docx-editor');
		await expect(b.locator('.ProseMirror')).toContainText('Shared document');
		await a.getByRole('tab', { name: 'Review', exact: true }).click();
		await a.getByRole('button', { name: 'Track changes', exact: true }).click();
		await page.getByRole('button', { name: 'Pause delivery', exact: true }).click();
		for (const [editor, text] of [
			[a, ' AdaEdit'],
			[b, ' GraceEdit'],
		] as const) {
			await editor.locator('.ProseMirror').click();
			await page.keyboard.press('Control+End');
			await page.keyboard.insertText(text);
		}
		await page.getByRole('button', { name: 'Resume delivery', exact: true }).click();
		const grace = a.locator('.ProseMirror ins[data-author="Grace"]');
		await expect(grace).toContainText('GraceEdit');
		await grace.dblclick();
		await a.getByRole('button', { name: 'Reject', exact: true }).click();
		for (const editor of [a, b]) {
			await expect(editor.locator('.ProseMirror')).not.toContainText('GraceEdit');
			await expect(editor.locator('.ProseMirror ins[data-author="Ada"]')).toContainText('AdaEdit');
			await expect(editor.locator('.ProseMirror del')).toHaveCount(0);
		}
		await a.locator('.ProseMirror').click();
		await page.keyboard.press('Control+z');
		for (const editor of [a, b]) {
			await expect(editor.locator('.ProseMirror ins[data-author="Grace"]')).toContainText(
				'GraceEdit',
			);
			await expect(editor.locator('.ProseMirror ins[data-author="Ada"]')).toContainText('AdaEdit');
		}
		await page.keyboard.press('Control+Shift+z');
		await expect(b.locator('.ProseMirror')).not.toContainText('GraceEdit');
		await expect(b.locator('.ProseMirror ins[data-author="Ada"]')).toContainText('AdaEdit');
		expect(errors).toEqual([]);
	});
	test(`${framework}: shared Track Changes records peer edits and exports its setting`, async ({
		page,
	}) => {
		const errors: string[] = [];
		page.on('pageerror', (error) => errors.push(error.message));
		await page.setViewportSize({ width: 2400, height: 1000 });
		await page.goto(`/collaboration.html?framework=${framework}&mode=yjs`);
		const a = page.locator('#peer-a docx-editor');
		const b = page.locator('#peer-b docx-editor');
		await expect(a.locator('.ProseMirror')).toContainText('Shared document');
		await expect(b.locator('.ProseMirror')).toContainText('Shared document');
		await a.getByRole('tab', { name: 'Review', exact: true }).click();
		const toggle = a.getByRole('button', { name: 'Track changes', exact: true });
		await toggle.click();
		await expect(toggle).toHaveAttribute('aria-pressed', 'true');
		await expect
			.poll(() =>
				b.evaluate((element) => (element as DocxEditorElement).documentModel!.trackChanges),
			)
			.toBe(true);
		await b.locator('.ProseMirror').click();
		await page.keyboard.press('Control+End');
		await page.keyboard.type(' peer revision');
		await expect(a.locator('.ProseMirror')).toContainText('peer revision');
		const revisions = await a.evaluate((element) =>
			(element as DocxEditorElement).documentModel!.blocks.flatMap((block) =>
				block.type === 'paragraph'
					? block.runs
							.filter((run) => run.revision)
							.map((run) => ({
								text: run.text,
								kind: run.revision!.kind,
								author: run.revision!.author,
								date: run.revision!.date,
							}))
					: [],
			),
		);
		expect(revisions.map((revision) => revision.text).join('')).toContain('peer revision');
		expect(
			revisions.every((revision) => revision.kind === 'insert' && revision.author === 'Grace'),
		).toBe(true);
		expect(
			revisions.every((revision) => revision.date && Number.isFinite(Date.parse(revision.date))),
		).toBe(true);
		const bytes = await a.evaluate(async (element) =>
			Array.from(await (element as DocxEditorElement).saveBytes()),
		);
		const zip = await JSZip.loadAsync(new Uint8Array(bytes));
		expect(await zip.file('word/settings.xml')!.async('string')).toContain('trackRevisions');
		expect(await zip.file('word/document.xml')!.async('string')).toContain('w:ins');
		expect(await zip.file('word/document.xml')!.async('string')).toContain('w:date=');
		await a.locator('.ProseMirror').click();
		await page.keyboard.press('Control+z');
		await expect(toggle).toHaveAttribute('aria-pressed', 'false');
		await expect
			.poll(() =>
				b.evaluate((element) => (element as DocxEditorElement).documentModel!.trackChanges),
			)
			.toBe(false);
		await expect(b.locator('.ProseMirror')).toContainText('peer revision');
		expect(errors).toEqual([]);
	});
}
