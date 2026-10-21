import { expect, test, type Page } from '@playwright/test';
import type { EditorView } from 'prosemirror-view';
import { savedDocument } from './saved-document';
import { fileInput } from './helpers';
import {
	blankFieldFixture,
	formFieldsFixture,
	protectedSectionsFixture,
} from './support/form-fields-fixture';

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

	test(`${framework}: a protected form's check box and drop-down work from the keyboard`, async ({
		page,
	}) => {
		const editor = await open(page, framework, formFieldsFixture(true));
		const box = editor.locator('.ProseMirror .dve-form-field[data-form-field="checkbox"]');
		const status = editor.locator('.ProseMirror .dve-form-field[data-form-field="dropDown"]');
		// Tab from a label goes to the next cell, the field, as in Word; Space toggles it there.
		await editor.locator('.ProseMirror td', { hasText: 'Approved' }).click();
		await expect
			.poll(() =>
				editor.evaluate(
					(element) =>
						(element as unknown as { view: EditorView }).view.state.selection.$from.parent
							.textContent,
				),
			)
			.toBe('Approved');
		await page.keyboard.press('Tab');
		await page.keyboard.press('Space');
		await expect(box).toHaveText('☒');
		await page.keyboard.press('Space');
		await expect(box).toHaveText('☐');
		await page.keyboard.press('Tab');
		await page.keyboard.press('Tab');
		// The list opens on the current entry; the arrows move through it and Enter picks one.
		await page.keyboard.press('Alt+ArrowDown');
		await expect(editor.getByRole('option', { name: 'Revised', exact: true })).toBeFocused();
		await page.keyboard.press('ArrowDown');
		await page.keyboard.press('Enter');
		await expect(status).toHaveText('Withdrawn');

		const { xml } = await savedDocument(page, `${framework}-form-keys.docx`);
		expect(xml).toContain('<w:result w:val="2"/>');
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

	test(`${framework}: typing into a blank text field of a protected form replaces its en spaces`, async ({
		page,
	}) => {
		const editor = await open(page, framework, blankFieldFixture());
		const code = editor.locator('.ProseMirror span[data-field="FORMTEXT"]');
		expect(await code.evaluate((node) => node.textContent)).toBe('\u2002'.repeat(5));
		// The editor has focus before the click, so the click only places the caret in the field.
		await editor.locator('.ProseMirror').evaluate((node: HTMLElement) => node.focus());
		await code.click();
		await page.keyboard.type('AB12!');
		// The first character takes the en spaces' place, and the field holds four at most.
		await expect(code).toHaveText('AB12');
		await expect(editor.locator('.ProseMirror > p').first()).toHaveText('Code: AB12 end');

		const { xml } = await savedDocument(page, `${framework}-blank-field.docx`);
		// What was typed keeps the result's properties, as Word's own result runs had them.
		expect(xml).toMatch(/<w:rPr[^>]*><w:noProof\/><\/w:rPr><w:t>AB12<\/w:t>/);
		expect(xml).not.toContain('\u2002');
		expect(xml).toContain('<w:maxLength w:val="4"/>');
	});
}
