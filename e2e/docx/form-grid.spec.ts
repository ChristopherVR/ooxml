import { expect, test } from '@playwright/test';
import { savedDocument } from './saved-document';
import { clickInEditor, fileInput } from './helpers';
import { formGridFixture, formGridRowCellCounts } from './support/form-grid-fixture';

const cellCounts = (xml: string) =>
	[...xml.matchAll(/<w:tr\b[\s\S]*?<\/w:tr>/g)].map(
		(row) => (row[0].match(/<w:tc\b/g) ?? []).length,
	);

for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid']) {
	test(`${framework}: values typed into a merged form grid and a toggled check box are saved`, async ({
		page,
	}) => {
		await page.goto(`/?framework=${framework}`);
		await (
			await fileInput(page)
		).setInputFiles({
			name: 'form-grid.docx',
			mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
			buffer: Buffer.from(await formGridFixture()),
		});
		const editor = page.locator('docx-editor');
		// The value cell beside the first label; its text changes, so it is located by position.
		const value = editor.locator('.ProseMirror td').nth(1);
		await expect(value).toHaveText('Number');
		await clickInEditor(value);
		await page.keyboard.press('End');
		await page.keyboard.type(' DOC-0042');
		await expect(value).toHaveText('Number DOC-0042');
		const box = editor.locator('.ProseMirror .dve-sdt[data-sdt-kind="checkbox"]');
		await box.click();
		await expect(box).toHaveText('☒');

		const { xml } = await savedDocument(page, `${framework}-form-grid.docx`);
		expect(xml).toContain('Number DOC-0042');
		expect(xml).toContain('<w14:checked w14:val="1"/>');
		expect(xml).toContain('<w:gridSpan w:val="3"/>');
		expect(cellCounts(xml)).toEqual(formGridRowCellCounts);
	});
}
