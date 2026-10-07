import { fileURLToPath } from 'node:url';
import { test, expect } from '@playwright/test';
import JSZip from 'jszip';
import { fileInput } from './helpers';
import type { DocxEditorElement } from '../../viewers/docx/packages/web-component/src';

const fixture = fileURLToPath(
	new URL(
		'../../src/core/docx/__fixtures__/review-paragraph-formatting/multiple-before.docx',
		import.meta.url,
	),
);
for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid'])
	test(`${framework}: records paragraph formatting with review, export and undo`, async ({
		page,
	}) => {
		await page.setViewportSize({ width: 2400, height: 1000 });
		const errors: string[] = [];
		page.on('pageerror', (error) => errors.push(error.message));
		await page.goto(`/?framework=${framework}`);
		await (await fileInput(page)).setInputFiles(fixture);
		const editor = page.locator('docx-editor');
		const body = editor.locator('.ProseMirror');
		await expect(body).toContainText('Paragraph');
		await editor.getByRole('tab', { name: 'Review', exact: true }).click();
		await editor.getByRole('button', { name: 'Track changes', exact: true }).click();
		await body.click();
		await page.keyboard.press('Control+Home');
		await page.keyboard.press('Control+e');
		const revision = body.locator('p.dve-revision-format-markup');
		await expect(revision).toHaveCount(1);
		const bytes = await editor.evaluate(async (element) =>
			Array.from(await (element as DocxEditorElement).saveBytes()),
		);
		const xml = await (
			await JSZip.loadAsync(new Uint8Array(bytes))
		)
			.file('word/document.xml')!
			.async('string');
		expect(xml).toContain('pPrChange');
		expect(xml).toContain('dateUtc');
		await page.keyboard.press('Control+z');
		await expect(revision).toHaveCount(0);
		await page.keyboard.press('Control+Shift+z');
		await expect(revision).toHaveCount(1);
		await editor.getByRole('button', { name: 'Next change', exact: true }).click();
		await editor.getByRole('button', { name: 'Reject', exact: true }).click();
		await expect(revision).toHaveCount(0);
		const paragraph = await editor.evaluate(
			(element) => (element as DocxEditorElement).documentModel!.blocks[0],
		);
		expect(paragraph?.type).toBe('paragraph');
		if (paragraph?.type !== 'paragraph') throw new Error('Expected paragraph');
		expect(paragraph.align).toBeUndefined();
		expect(paragraph.sourceParagraphPropertiesXml).toContain('Arial');
		expect(errors).toEqual([]);
	});
