import { fileURLToPath } from 'node:url';
import { test, expect } from '@playwright/test';
import { fileInput } from './helpers';

const fixture = fileURLToPath(
	new URL(
		'../../src/core/docx/__fixtures__/page-break-style/page-break-style.docx',
		import.meta.url,
	),
);
for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid'])
	test(`${framework}: respects inherited page breaks and explicit off in print layout`, async ({
		page,
	}) => {
		await page.goto(`/?framework=${framework}`);
		await (await fileInput(page)).setInputFiles(fixture);
		const editor = page.locator('docx-editor');
		await expect(editor.locator('.ProseMirror')).toContainText('Explicitly disabled');
		await editor.getByRole('tab', { name: 'View', exact: true }).click();
		await editor.getByRole('combobox', { name: 'Layout view', exact: true }).selectOption('print');
		const pages = editor.locator('.dve-canvas > .dve-print-pages .dve-print-page');
		await expect(pages).toHaveCount(2);
		await expect(pages.nth(0)).toContainText('First paragraph');
		await expect(pages.nth(0)).not.toContainText('Inherited page break');
		await expect(pages.nth(1)).toContainText('Inherited page break');
		await expect(pages.nth(1)).toContainText('Explicitly disabled');
	});
