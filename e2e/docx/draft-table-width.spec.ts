import { expect, test } from '@playwright/test';
import { readFile } from 'node:fs/promises';
import { fileInput } from './helpers';

// The table report fixtures (generated in the core, see their README): an auto-fit table of
// twelve columns, a centred one with thick borders and large cell margins, and a fixed one whose
// cells all claim the table's width (#322).
const fixture = (name: string) =>
	readFile(
		new URL(`../../src/core/docx/layout/fixtures/table-reports/${name}.docx`, import.meta.url),
	);

for (const name of [
	't02-autofit-12-colonnes',
	't15-centre-bordures-epaisses',
	't25-largeurs-cellules-perimees',
]) {
	test(`Draft view keeps the table of ${name} inside the paper, as Print Layout does (#322)`, async ({
		page,
	}) => {
		page.on('dialog', (dialog) => void dialog.accept());
		await page.goto('/?framework=vanilla');
		await (
			await fileInput(page)
		).setInputFiles({
			name: `${name}.docx`,
			mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
			buffer: await fixture(name),
		});
		const editor = page.locator('docx-editor');
		await editor.getByRole('tab', { name: 'View', exact: true }).click();
		await editor.getByRole('combobox', { name: 'Layout view', exact: true }).selectOption('draft');
		const table = editor.locator('.ProseMirror table');
		await expect(table).toHaveCount(1);
		const paper = await editor.locator('.dve-paper').evaluate((element) => {
			const box = element.getBoundingClientRect();
			const style = getComputedStyle(element);
			return {
				left: box.left + parseFloat(style.paddingLeft),
				right: box.right - parseFloat(style.paddingRight),
			};
		});
		const box = (await table.boundingBox())!;
		// Half of a collapsed outer border may sit outside the text column, as in Word.
		expect(box.x).toBeGreaterThanOrEqual(paper.left - 4);
		expect(box.x + box.width).toBeLessThanOrEqual(paper.right + 4);
	});
}
