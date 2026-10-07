import { test, expect } from '@playwright/test';
import { fileURLToPath } from 'node:url';
import type { DocxEditorElement } from '../../viewers/docx/packages/web-component/src';
import { fileInput } from './helpers';

for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid'])
	for (const name of ['picture', 'note', 'break', 'field'])
		test(`${framework}: native ${name} formatting review retains content and undo`, async ({
			page,
		}) => {
			await page.setViewportSize({ width: 2400, height: 1000 });
			await page.goto(`/?framework=${framework}`);
			await (
				await fileInput(page)
			).setInputFiles(
				fileURLToPath(
					new URL(
						`../../src/core/docx/__fixtures__/review-object-formatting/${name}-tracked.docx`,
						import.meta.url,
					),
				),
			);
			const editor = page.locator('docx-editor');
			const body = editor.locator('.dve-paper > .ProseMirror');
			await expect(body).toContainText('Before');
			const source = await editor.evaluate((el) => (el as DocxEditorElement).documentModel!);
			const objectRuns = () =>
				editor.evaluate((el) =>
					(el as DocxEditorElement).documentModel!.blocks.flatMap((block) =>
						block.type === 'paragraph'
							? block.runs.filter(
									(run) =>
										run.image || run.break || run.noteReference || run.fieldCode !== undefined,
								)
							: [],
					),
				);
			expect((await objectRuns()).filter((run) => run.revision || run.formatRevision)).toHaveLength(
				1,
			);
			await editor.getByRole('tab', { name: 'Review', exact: true }).click();
			await editor.getByRole('button', { name: 'Accept all', exact: true }).click();
			const accepted = await objectRuns();
			expect(accepted.filter((run) => run.revision || run.formatRevision)).toHaveLength(0);
			expect(accepted.filter((run) => run.bold)).toHaveLength(1);
			await body.press('Control+z');
			expect(await editor.evaluate((el) => (el as DocxEditorElement).documentModel)).toEqual(
				source,
			);
			await editor.getByRole('button', { name: 'Reject all', exact: true }).click();
			const rejected = await objectRuns();
			expect(rejected).toHaveLength(accepted.length);
			expect(rejected.filter((run) => run.revision || run.formatRevision || run.bold)).toHaveLength(
				0,
			);
			await expect(body).toContainText('After');
			if (name === 'picture') await expect(body.locator('.dve-picture')).toBeVisible();
			if (name === 'note') await expect(body.locator('.dve-note-reference')).toBeVisible();
			await body.press('Control+z');
			expect(await editor.evaluate((el) => (el as DocxEditorElement).documentModel)).toEqual(
				source,
			);
		});
