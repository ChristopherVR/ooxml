import { expect, test, type Page } from '@playwright/test';
import { savedDocument } from './saved-document';
import { fileInput } from './helpers';
import { formFieldsFixture, protectedSectionsFixture } from './support/form-fields-fixture';

async function open(page: Page, framework: string, docx: Promise<Uint8Array>) {
	await page.goto(`/?framework=${framework}`);
	await (
		await fileInput(page)
	).setInputFiles({
		name: 'form-fields.docx',
		mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
		buffer: Buffer.from(await docx),
	});
	const editor = page.locator('docx-editor');
	await expect(editor.locator('.ProseMirror')).toBeVisible();
	return editor;
}

for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid']) {
	test(`${framework}: legacy form fields are shaded, filled, toggled, picked and saved`, async ({
		page,
	}) => {
		const editor = await open(page, framework, formFieldsFixture());
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
		const editor = await open(page, framework, formFieldsFixture(true));
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

	test(`${framework}: a protected form keeps locked fields and leaves unprotected sections editable`, async ({
		page,
	}) => {
		const editor = await open(page, framework, protectedSectionsFixture());
		const boxes = editor.locator('.ProseMirror .dve-form-field[data-form-field="checkbox"]');
		await expect(boxes).toHaveCount(2);
		await boxes.nth(1).click();
		await expect(boxes.nth(1)).toHaveText('☐');
		await boxes.nth(0).click();
		await expect(boxes.nth(0)).toHaveText('☒');
		const end = editor.locator('.ProseMirror > p', { hasText: 'End of form' });
		await end.click();
		await page.keyboard.press('End');
		await page.keyboard.type('!');
		await expect(end).toHaveText('End of form');
		const notes = editor.locator('.ProseMirror > p', { hasText: 'Notes' });
		await notes.click();
		await page.keyboard.press('End');
		await page.keyboard.type(' to add');
		await expect(notes).toHaveText('Notes to add');
		// A paragraph added at the end of the last section is in that section too.
		await page.keyboard.press('Enter');
		await page.keyboard.type('More notes');
		await expect(editor.locator('.ProseMirror > p').last()).toHaveText('More notes');

		const { xml } = await savedDocument(page, `${framework}-protected-sections.docx`);
		expect(xml).toContain('>Notes to add<');
		expect(xml).toContain('>More notes<');
		const field = (name: string) =>
			xml.slice(
				xml.indexOf(`<w:name w:val="${name}"/>`),
				xml.indexOf('</w:ffData>', xml.indexOf(name)),
			);
		expect(field('Free')).toContain('<w:checked/>');
		expect(field('Locked')).not.toContain('<w:checked');
	});
}
