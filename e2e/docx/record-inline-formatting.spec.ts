import { test, expect } from '@playwright/test';
import { fileURLToPath } from 'node:url';
import type { DocxEditorElement } from '../../viewers/docx/packages/web-component/src';
import { fileInput } from './helpers';

for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid'])
	for (const name of ['picture', 'note', 'line-break'])
		test(`${framework}: ribbon formatting records text and native ${name} in one history operation`, async ({
			page,
		}) => {
			await page.setViewportSize({ width: 2400, height: 1000 });
			await page.goto(`/?framework=${framework}`);
			await (
				await fileInput(page)
			).setInputFiles(
				fileURLToPath(
					new URL(
						`../../src/core/docx/__fixtures__/${name === 'line-break' ? 'review-line-break-formatting' : 'review-object-formatting'}/${name}-before.docx`,
						import.meta.url,
					),
				),
			);
			const editor = page.locator('docx-editor');
			const body = editor.locator('.dve-paper > .ProseMirror');
			await expect(body).toContainText('Before');
			await editor.getByRole('tab', { name: 'Review', exact: true }).click();
			await editor.getByRole('button', { name: 'Track changes', exact: true }).click();
			await expect
				.poll(() => editor.evaluate((el) => (el as DocxEditorElement).documentModel!.trackChanges))
				.toBe(true);
			const source = await editor.evaluate((el) => (el as DocxEditorElement).documentModel);
			// ProseMirror coalesces Word's adjacent physical runs with identical properties.
			if (name === 'line-break')
				for (const block of source!.blocks) {
					if (block.type !== 'paragraph') continue;
					const merged: typeof block.runs = [];
					for (const run of block.runs) {
						const previous = merged.at(-1);
						const { text, ...properties } = run;
						const { text: _previousText, ...previousProperties } = previous ?? { text: '' };
						if (previous && JSON.stringify(properties) === JSON.stringify(previousProperties))
							previous.text += text;
						else merged.push({ ...run });
					}
					block.runs = merged;
				}
			await body.click();
			await body.press('Control+a');
			await editor.getByRole('tab', { name: 'Home', exact: true }).click();
			await editor.getByRole('button', { name: 'Bold', exact: true }).click();
			const tracked = await editor.evaluate((el) => (el as DocxEditorElement).documentModel!);
			const runs = tracked.blocks.flatMap((block) =>
				block.type === 'paragraph' ? block.runs : [],
			);
			expect(runs.every((run) => run.bold)).toBe(true);
			expect(
				runs.find((run) => run.image || run.noteReference || run.text.includes('\n'))!
					.formatRevision?.kind,
			).toBe('formatChange');
			await body.press('Control+z');
			expect(await editor.evaluate((el) => (el as DocxEditorElement).documentModel)).toEqual(
				source,
			);
			await body.press('Control+Shift+z');
			expect(await editor.evaluate((el) => (el as DocxEditorElement).documentModel)).toEqual(
				tracked,
			);
			await editor.getByRole('tab', { name: 'Review', exact: true }).click();
			await editor.getByRole('button', { name: 'Reject all', exact: true }).click();
			const rejected = await editor.evaluate((el) => (el as DocxEditorElement).documentModel!);
			expect(
				rejected.blocks
					.flatMap((block) => (block.type === 'paragraph' ? block.runs : []))
					.every((run) => !run.bold && !run.revision && !run.formatRevision),
			).toBe(true);
			await body.press('Control+z');
			expect(await editor.evaluate((el) => (el as DocxEditorElement).documentModel)).toEqual(
				tracked,
			);
		});
