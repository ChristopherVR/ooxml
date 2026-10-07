import { test, expect } from '@playwright/test';
import { fileURLToPath } from 'node:url';
import type { DocxEditorElement } from '../../viewers/docx/packages/web-component/src';
import { newDocument } from './helpers';

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

for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid'])
	test(`${framework}: notes with the same id keep their separate text`, async ({ page }) => {
		await page.setViewportSize({ width: 2400, height: 1000 });
		await page.goto(`/?framework=${framework}`);
		await newDocument(page);
		const editor = page.locator('docx-editor');
		await expect(editor.locator('.dve-paper > .ProseMirror')).toBeVisible();
		await editor.evaluate((el) => {
			const host = el as DocxEditorElement;
			host.documentModel = {
				...host.documentModel!,
				footnotes: [
					{ id: '1', blocks: [{ type: 'paragraph', id: 'fn', runs: [{ text: 'Footnote text' }] }] },
				],
				endnotes: [
					{ id: '1', blocks: [{ type: 'paragraph', id: 'en', runs: [{ text: 'Endnote text' }] }] },
				],
			};
		});
		const footnote = editor.locator('.dve-notes-footnote .dve-note-body');
		const endnote = editor.locator('.dve-notes-endnote .dve-note-body');
		await expect(footnote).toContainText('Footnote text');
		await expect(endnote).toContainText('Endnote text');
		await endnote.dblclick();
		const active = endnote.locator('.ProseMirror');
		await expect(active).toContainText('Endnote text');
		await active.fill('Updated endnote');
		await active.press('Escape');
		await editor.getByRole('tab', { name: 'Review', exact: true }).click();
		await editor
			.getByRole('combobox', { name: 'Display for review', exact: true })
			.selectOption('original', { force: true });
		await expect(footnote).toContainText('Footnote text');
		await expect(endnote).toContainText('Updated endnote');
		const model = await editor.evaluate((el) => (el as DocxEditorElement).documentModel);
		expect(model!.footnotes![0]!.blocks[0]).toMatchObject({ runs: [{ text: 'Footnote text' }] });
		expect(model!.endnotes![0]!.blocks[0]).toMatchObject({ runs: [{ text: 'Updated endnote' }] });
	});
