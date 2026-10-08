import { test, expect } from '@playwright/test';
import { fileURLToPath } from 'node:url';
import type { DocxEditorElement } from '../../viewers/docx/packages/web-component/src';
import { fileInput } from './helpers';

const fixture = fileURLToPath(
	new URL(
		'../../src/core/docx/__fixtures__/review-stories/all-stories-tracked.docx',
		import.meta.url,
	),
);
for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid'])
	test(`${framework}: document-wide review resolves all five native stories`, async ({ page }) => {
		await page.setViewportSize({ width: 2400, height: 1000 });
		await page.goto(`/?framework=${framework}`);
		await (await fileInput(page)).setInputFiles(fixture);
		const editor = page.locator('docx-editor');
		const body = editor.locator('.dve-paper > .ProseMirror');
		await expect(body).toContainText('Body');
		await expect(editor.locator('.dve-header')).toContainText('Header text');
		await expect(editor.locator('.dve-footer')).toContainText('Footer text');
		const source = await editor.evaluate((el) => (el as DocxEditorElement).documentModel);
		const pending = () =>
			editor.evaluate((el) => {
				const model = (el as DocxEditorElement).documentModel!;
				const blocks: NonNullable<DocxEditorElement['documentModel']>['blocks'] = [
					...model.blocks,
					...(model.sections ?? []).flatMap((section) =>
						[
							...Object.values(section.headers ?? {}),
							...Object.values(section.footers ?? {}),
						].flatMap((part) => part?.blocks ?? []),
					),
					...[...(model.footnotes ?? []), ...(model.endnotes ?? [])].flatMap((note) => note.blocks),
				];
				return blocks
					.flatMap((block) =>
						block.type === 'paragraph'
							? [block]
							: block.rows.flatMap((row) => row.flatMap((cell) => cell.paragraphs)),
					)
					.reduce(
						(sum, paragraph) =>
							sum +
							Number(!!paragraph.markRevision) +
							Number(!!paragraph.formatRevision) +
							paragraph.runs.reduce(
								(count, run) => count + Number(!!run.revision) + Number(!!run.formatRevision),
								0,
							),
						0,
					);
			});
		expect(await pending()).toBe(5);
		await editor.getByRole('tab', { name: 'Review', exact: true }).click();
		await editor.getByRole('button', { name: 'Accept all', exact: true }).click();
		expect(await pending()).toBe(0);
		await expect(editor.locator('.dve-header strong')).toContainText('Header text');
		await expect(editor.locator('.dve-notes-footnote strong')).toContainText('Footnote text');
		await body.press('Control+z');
		expect(await editor.evaluate((el) => (el as DocxEditorElement).documentModel)).toEqual(source);
		await editor.getByRole('button', { name: 'Reject all', exact: true }).click();
		expect(await pending()).toBe(0);
		await expect(editor.locator('.dve-header strong')).toHaveCount(0);
		await expect(editor.locator('.dve-footer strong')).toHaveCount(0);
		await expect(editor.locator('.dve-notes-footnote strong')).toHaveCount(0);
		await expect(editor.locator('.dve-notes-endnote strong')).toHaveCount(0);
		await body.press('Control+z');
		expect(await editor.evaluate((el) => (el as DocxEditorElement).documentModel)).toEqual(source);
	});
