import { fileURLToPath } from 'node:url';
import { expect, test } from '@playwright/test';
import JSZip from 'jszip';
import { fileInput } from './helpers';
import type { DocxEditorElement } from '../../viewers/docx/packages/web-component/src';

const fixture = fileURLToPath(
	new URL(
		'../../src/core/docx/__fixtures__/review-formatting/multiple-tracked.docx',
		import.meta.url,
	),
);
for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid']) {
	test(`${framework}: text editing preserves native formatting history`, async ({ page }) => {
		await page.setViewportSize({ width: 2400, height: 1000 });
		const errors: string[] = [];
		page.on('pageerror', (error) => errors.push(error.message));
		await page.goto(`/?framework=${framework}`);
		await (await fileInput(page)).setInputFiles(fixture);
		const editor = page.locator('docx-editor');
		const body = editor.locator('.ProseMirror');
		await expect(body).toContainText('Format me');
		await editor.getByRole('tab', { name: 'Review', exact: true }).click();
		await editor.getByRole('button', { name: 'Track changes', exact: true }).click();
		await expect
			.poll(() =>
				editor.evaluate((element) => (element as DocxEditorElement).documentModel!.trackChanges),
			)
			.toBe(false);
		await body.click();
		await page.keyboard.press('Control+End');
		await page.keyboard.type('!');
		await expect(body).toContainText('Format me!');
		const bytes = await editor.evaluate(async (element) =>
			Array.from(await (element as DocxEditorElement).saveBytes()),
		);
		const xml = await (
			await JSZip.loadAsync(new Uint8Array(bytes))
		)
			.file('word/document.xml')!
			.async('string');
		expect(xml).toContain('rPrChange');
		expect(xml).toContain('Format me!');
		expect(xml).toMatch(/<w:rPrChange[^>]*>[\s\S]*?<w:color w:val="0000FF"/);
		expect(errors).toEqual([]);
	});
}
