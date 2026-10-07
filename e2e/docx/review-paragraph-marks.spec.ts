import { test, expect } from '@playwright/test';
import { fileURLToPath } from 'node:url';
import JSZip from 'jszip';
import type { DocxEditorElement } from '../../viewers/docx/packages/web-component/src';
import { fileInput } from './helpers';

const fixture = fileURLToPath(
	new URL(
		'../../src/core/docx/__fixtures__/review-inline/break-insert-tracked.docx',
		import.meta.url,
	),
);
for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid'])
	test(`${framework}: native paragraph marks resolve with the inserted page break`, async ({
		page,
	}) => {
		await page.setViewportSize({ width: 2400, height: 1000 });
		await page.goto(`/?framework=${framework}`);
		await (await fileInput(page)).setInputFiles(fixture);
		const editor = page.locator('docx-editor');
		const body = editor.locator('.dve-paper > .ProseMirror');
		await expect(body).toContainText('Before');
		await expect(body.locator(':scope > p')).toHaveCount(3);
		const source = await editor.evaluate((el) => (el as DocxEditorElement).documentModel);
		await editor.getByRole('tab', { name: 'Review', exact: true }).click();
		await editor.getByRole('button', { name: 'Accept all', exact: true }).click();
		await expect(body.locator(':scope > p')).toHaveCount(3);
		await expect(body.locator('.dve-break-marker')).toHaveCount(1);
		await body.press('Control+z');
		expect(await editor.evaluate((el) => (el as DocxEditorElement).documentModel)).toEqual(source);
		await editor.getByRole('button', { name: 'Reject all', exact: true }).click();
		await expect(body.locator(':scope > p')).toHaveCount(1);
		await expect(body).toContainText('BeforeAfter');
		await expect(body.locator('.dve-break-marker')).toHaveCount(0);
		const bytes = await editor.evaluate(async (el) =>
			Array.from(await (el as DocxEditorElement).saveBytes()),
		);
		const xml = await (
			await JSZip.loadAsync(new Uint8Array(bytes))
		)
			.file('word/document.xml')!
			.async('string');
		expect(xml).not.toMatch(/<w:(?:ins|del)(?:\s|\/|>)/);
		await body.press('Control+z');
		expect(await editor.evaluate((el) => (el as DocxEditorElement).documentModel)).toEqual(source);
	});
