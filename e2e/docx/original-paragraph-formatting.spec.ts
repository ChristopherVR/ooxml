import { fileURLToPath } from 'node:url';
import { test, expect } from '@playwright/test';
import JSZip from 'jszip';
import { fileInput } from './helpers';
import type { DocxEditorElement } from '../../viewers/docx/packages/web-component/src';

const fixture = (suffix: string) =>
	fileURLToPath(
		new URL(
			`../../src/core/docx/__fixtures__/review-paragraph-formatting/multiple-${suffix}.docx`,
			import.meta.url,
		),
	);

for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid'])
	test(`${framework}: Original displays native prior paragraph formatting without resolving changes`, async ({
		page,
	}) => {
		await page.setViewportSize({ width: 2400, height: 1000 });
		const errors: string[] = [];
		page.on('pageerror', (error) => errors.push(error.message));
		await page.goto(`/?framework=${framework}`);
		const input = await fileInput(page);
		const editor = page.locator('docx-editor');
		const paragraph = editor.locator('.ProseMirror p').first();
		const appearance = () =>
			paragraph.evaluate((element) => {
				const style = getComputedStyle(element);
				return {
					align: style.textAlign,
					before: style.marginTop,
					after: style.marginBottom,
					left: style.marginLeft,
					right: style.marginRight,
					indent: style.textIndent,
					line: style.lineHeight,
				};
			});
		await input.setInputFiles(fixture('before'));
		await expect(paragraph).toContainText('Paragraph formatting');
		const before = await appearance();
		await input.setInputFiles(fixture('tracked'));
		await expect(paragraph).toHaveClass(/dve-revision-format-markup/);
		const current = await appearance();
		expect(current).not.toEqual(before);
		const source = await editor.evaluate((element) => (element as DocxEditorElement).documentModel);
		await editor.getByRole('tab', { name: 'Review', exact: true }).click();
		const mode = editor.getByRole('combobox', { name: 'Display for review', exact: true });
		await mode.selectOption('original', { force: true });
		await expect.poll(appearance).toEqual(before);
		expect(
			await editor.evaluate((element) => (element as DocxEditorElement).documentModel),
		).toEqual(source);
		const bytes = await editor.evaluate(async (element) =>
			Array.from(await (element as DocxEditorElement).saveBytes()),
		);
		const xml = await (
			await JSZip.loadAsync(new Uint8Array(bytes))
		)
			.file('word/document.xml')!
			.async('string');
		expect(xml).toContain('pPrChange');
		await mode.selectOption('final', { force: true });
		await expect.poll(appearance).toEqual(current);
		expect(errors).toEqual([]);
	});
