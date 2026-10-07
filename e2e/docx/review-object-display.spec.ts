import { test, expect } from '@playwright/test';
import { fileURLToPath } from 'node:url';
import type { DocxEditorElement } from '../../viewers/docx/packages/web-component/src';
import { fileInput } from './helpers';

for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid'])
	for (const [name, selector] of [
		['picture', '.dve-picture'],
		['note', '.dve-note-reference'],
		['break', '.dve-break-marker'],
		['field', '[data-field-marker="code"]'],
	])
		test(`${framework}: Original ${name} properties match the native before reference`, async ({
			page,
		}) => {
			await page.setViewportSize({ width: 2400, height: 1000 });
			await page.goto(`/?framework=${framework}`);
			const upload = async (state: string) =>
				(await fileInput(page)).setInputFiles(
					fileURLToPath(
						new URL(
							`../../src/core/docx/__fixtures__/review-object-formatting/${name}-${state}.docx`,
							import.meta.url,
						),
					),
				);
			await upload('before');
			const editor = page.locator('docx-editor');
			const body = editor.locator('.dve-paper > .ProseMirror');
			await expect(body).toContainText('Before');
			const appearance = () =>
				body.locator(selector!).evaluate((el) => {
					const style = getComputedStyle(el);
					return {
						weight: style.fontWeight,
						font: style.fontFamily,
						size: style.fontSize,
						align: style.verticalAlign,
						color: style.color,
					};
				});
			const before = await appearance();
			await upload('tracked');
			await expect
				.poll(() =>
					editor.evaluate(
						(el) =>
							(el as DocxEditorElement).documentModel!.blocks.flatMap((block) =>
								block.type === 'paragraph'
									? block.runs.filter((run) => run.revision || run.formatRevision)
									: [],
							).length,
					),
				)
				.toBe(1);
			const source = await editor.evaluate((el) => (el as DocxEditorElement).documentModel);
			await editor.getByRole('tab', { name: 'Review', exact: true }).click();
			const mode = editor.getByRole('combobox', { name: 'Display for review', exact: true });
			await mode.selectOption('original', { force: true });
			expect(await appearance()).toEqual(before);
			expect(await editor.evaluate((el) => (el as DocxEditorElement).documentModel)).toEqual(
				source,
			);
			await mode.selectOption('all', { force: true });
			expect((await appearance()).weight).toBe('700');
			expect(await editor.evaluate((el) => (el as DocxEditorElement).documentModel)).toEqual(
				source,
			);
		});
