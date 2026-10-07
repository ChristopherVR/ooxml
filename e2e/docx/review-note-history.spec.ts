import { test, expect } from '@playwright/test';
import { fileURLToPath } from 'node:url';
import type { DocxEditorElement } from '../../viewers/docx/packages/web-component/src';

const fixture = fileURLToPath(
	new URL(
		'../../src/core/docx/__fixtures__/review-inline/note-delete-tracked.docx',
		import.meta.url,
	),
);
for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid'])
	test(`${framework}: Accept All resolves the note story and undo restores it`, async ({
		page,
	}) => {
		await page.setViewportSize({ width: 2400, height: 1000 });
		await page.goto(`/?framework=${framework}`);
		await page.locator('input[type=file]').first().setInputFiles(fixture);
		const editor = page.locator('docx-editor');
		const body = editor.locator('.dve-paper > .ProseMirror');
		await expect(body).toContainText('Before');
		await expect(editor.locator('.dve-note-body')).toHaveCount(1);
		const source = await editor.evaluate((el) => (el as DocxEditorElement).documentModel);
		await editor.getByRole('tab', { name: 'Review', exact: true }).click();
		await editor.getByRole('button', { name: 'Accept all', exact: true }).click();
		await expect(body.locator('.dve-note-reference')).toHaveCount(0);
		await expect(editor.locator('.dve-note-body')).toHaveCount(0);
		await body.press('Control+z');
		expect(await editor.evaluate((el) => (el as DocxEditorElement).documentModel)).toEqual(source);
		await expect(editor.locator('.dve-note-body')).toHaveCount(1);
		await editor.getByRole('button', { name: 'Reject all', exact: true }).click();
		await expect(body.locator('.dve-note-reference')).toHaveCount(1);
		await expect(editor.locator('.dve-note-body')).toHaveCount(1);
		const note = await editor.evaluate(
			(el) => (el as DocxEditorElement).documentModel!.footnotes![0],
		);
		for (const block of note!.blocks) {
			if (block.type !== 'paragraph') continue;
			expect(block.markRevision).toBeUndefined();
			expect(block.formatRevision).toBeUndefined();
			for (const run of block.runs) {
				expect(run.revision).toBeUndefined();
				expect(run.formatRevision).toBeUndefined();
			}
		}
		await body.press('Control+z');
		expect(await editor.evaluate((el) => (el as DocxEditorElement).documentModel)).toEqual(source);
	});
