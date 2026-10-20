import { expect, test, type Page } from '@playwright/test';
import { savedDocument } from './saved-document';
import { fileInput } from './helpers';
import { formFieldsFixture } from './support/form-fields-fixture';

async function open(page: Page, framework: string, protect = false) {
	await page.goto(`/?framework=${framework}`);
	await (
		await fileInput(page)
	).setInputFiles({
		name: 'form-fields.docx',
		mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
		buffer: Buffer.from(await formFieldsFixture(protect)),
	});
	const editor = page.locator('docx-editor');
	await expect(editor.locator('.ProseMirror')).toBeVisible();
	return editor;
}

for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid']) {
	test(`${framework}: legacy form fields are shaded, filled, toggled, picked and saved`, async ({
		page,
	}) => {
		const editor = await open(page, framework);
		const name = editor.locator('.ProseMirror span[data-field="FORMTEXT"]');
		const box = editor.locator('.ProseMirror .dve-form-field[data-form-field="checkbox"]');
		const status = editor.locator('.ProseMirror .dve-form-field[data-form-field="dropDown"]');
		await expect(name).toHaveText('Name');
		await expect(box).toHaveText('☐');
		await expect(status).toHaveText('Revised');
		await expect(name).toHaveCSS('background-color', 'rgb(217, 217, 217)');
		await expect(box).toHaveCSS('background-color', 'rgb(217, 217, 217)');

		await name.dblclick();
		await page.keyboard.type('Ada Lovelace');
		await expect(name).toHaveText('Ada Lovelace');
		await box.click();
		await expect(box).toHaveText('☒');
		await status.click();
		await editor.getByRole('option', { name: 'Withdrawn', exact: true }).click();
		await expect(status).toHaveText('Withdrawn');

		const { xml } = await savedDocument(page, `${framework}-form-fields.docx`);
		expect(xml).toContain('<w:ffData><w:name w:val="Name"/>');
		expect(xml).toContain('>Ada Lovelace<');
		expect(xml).toContain(
			'<w:checkBox><w:sizeAuto/><w:default w:val="0"/><w:checked/></w:checkBox>',
		);
		expect(xml).toContain('<w:result w:val="2"/>');
		expect(xml.indexOf('<w:bookmarkStart w:id="1" w:name="Name"/>')).toBeLessThan(
			xml.indexOf('Ada Lovelace'),
		);
	});

	test(`${framework}: a document protected for forms only lets its fields change`, async ({
		page,
	}) => {
		const editor = await open(page, framework, true);
		const label = editor.locator('.ProseMirror > p').first();
		await label.click({ position: { x: 4, y: 4 } });
		await page.keyboard.press('Home');
		await page.keyboard.type('Refused ');
		await expect(label).not.toContainText('Refused');
		const box = editor.locator('.ProseMirror .dve-form-field[data-form-field="checkbox"]');
		await box.click();
		await expect(box).toHaveText('☒');
		const name = editor.locator('.ProseMirror span[data-field="FORMTEXT"]');
		await name.dblclick();
		await page.keyboard.type('Grace');
		await expect(name).toHaveText('Grace');
	});
}
