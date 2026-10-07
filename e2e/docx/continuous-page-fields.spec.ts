import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { expect, test } from '@playwright/test';
import { fileInput } from './helpers';

const fixture = (name: string) =>
	new URL(`../../src/core/docx/layout/fixtures/continuous-page-fields/${name}`, import.meta.url);
const evidence = JSON.parse(await readFile(fixture('evidence.json'), 'utf8')) as {
	cases: { name: string; pages: { header: string; footer: string }[] }[];
};

for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid'])
	for (const reference of evidence.cases)
		test(`${framework}: continuous page fields ${reference.name}`, async ({ page }) => {
			const errors: string[] = [];
			page.on('pageerror', (error) => errors.push(error.message));
			page.on('dialog', (dialog) => void dialog.accept());
			await page.goto(`/?framework=${framework}`);
			await (await fileInput(page)).setInputFiles(fileURLToPath(fixture(`${reference.name}.docx`)));
			const editor = page.locator('docx-editor');
			await expect(editor.locator('.ProseMirror')).toContainText('After120');
			await editor.getByRole('tab', { name: 'View', exact: true }).click();
			await editor.getByRole('button', { name: 'Print Layout', exact: true }).click();
			const sheets = editor.locator('.dve-canvas > .dve-print-pages .dve-print-page');
			await expect(sheets).toHaveCount(reference.pages.length);
			expect(
				await sheets.evaluateAll((elements) =>
					elements.map((sheet) => ({
						header: sheet.querySelector('.dve-print-header')!.textContent!.trim(),
						footer: sheet.querySelector('.dve-print-footer')!.textContent!.trim(),
					})),
				),
			).toEqual(reference.pages);
			expect(errors).toEqual([]);
		});
