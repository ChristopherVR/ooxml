import { expect, test, type Page } from '@playwright/test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { newDocument, reveal, saveButton, dialogByHeading, waitForEditorFocus } from './helpers';

async function savedPackage(page: Page, name: string): Promise<JSZip> {
	const downloadPromise = page.waitForEvent('download');
	await saveButton(page).click();
	const download = await downloadPromise;
	await download.saveAs(`test-results/${name}.docx`);
	return JSZip.loadAsync(await readFile((await download.path())!));
}

test('Quick Parts, Signature Line and the header and footer galleries write what Word reads', async ({
	page,
}) => {
	await page.goto('/?framework=vanilla');
	await newDocument(page);
	const editor = page.locator('docx-editor');
	const surface = editor.locator('.ProseMirror');
	await surface.click();
	await page.keyboard.type('Quarterly review');
	await editor.getByRole('tab', { name: 'Insert', exact: true }).click();

	// Header design with fields, then Remove Footer is a no-op on a document without one.
	const header = editor.getByRole('combobox', { name: 'Header', exact: true });
	await reveal(editor, header);
	await header.selectOption('titleAndAuthor');
	// The design shows placeholders, so the header is visible and clickable; Edit puts the caret in it.
	await expect(editor.locator('.dve-header')).toContainText('[Document title]');
	await header.selectOption('edit');
	await waitForEditorFocus(editor.locator('.dve-header .ProseMirror'));
	await page.keyboard.press('Escape');
	await editor.getByRole('tab', { name: 'Insert', exact: true }).click();

	// Document property: a control bound to the core properties part.
	await surface.click();
	await page.keyboard.press('Control+End');
	const quickParts = editor.getByRole('combobox', { name: 'Quick Parts', exact: true });
	await reveal(editor, quickParts);
	await quickParts.selectOption('property:author');
	await expect(surface).toContainText('[Author]');

	// Save the selection as a Quick Part and insert it again from the Organizer.
	await surface.click();
	await page.keyboard.press('Control+a');
	await reveal(editor, quickParts);
	await quickParts.selectOption('save');
	const save = dialogByHeading(editor, 'Create New Building Block');
	await save.getByRole('textbox', { name: 'Name', exact: true }).fill('Review line');
	await save.getByRole('button', { name: 'OK', exact: true }).click();
	await expect(save).toBeHidden();
	await surface.click();
	await page.keyboard.press('Control+End');
	await page.keyboard.press('Enter');
	await reveal(editor, quickParts);
	await quickParts.selectOption('organizer');
	const organizer = dialogByHeading(editor, 'Building Blocks Organizer');
	await expect(organizer.getByRole('listbox', { name: 'Building blocks' })).toContainText(
		'Review line',
	);
	await organizer.getByRole('button', { name: 'Insert', exact: true }).click();
	await expect(surface).toContainText('Quarterly review');

	// Signature Setup inserts the signature line; Sign stays disabled without the host's signer.
	const signature = editor.getByRole('combobox', { name: 'Signature line', exact: true });
	await reveal(editor, signature);
	await expect(signature.locator('option[value="sign"]')).toBeDisabled();
	await signature.selectOption('setup');
	const setup = dialogByHeading(editor, 'Signature Setup');
	await setup.getByRole('textbox', { name: 'Suggested signer', exact: true }).fill('Ada Lovelace');
	await setup.getByRole('button', { name: 'OK', exact: true }).click();
	await expect(surface.locator('img:not(.ProseMirror-separator)')).toBeVisible();

	// Online Video stays hidden until the host allows a video site.
	await expect(editor.getByRole('button', { name: 'Online Video', exact: true })).toBeHidden();

	const zip = await savedPackage(page, 'insert-parts');
	const body = await zip.file('word/document.xml')!.async('string');
	expect(body).toContain('issignatureline="t"');
	expect(body).toContain('o:suggestedsigner="Ada Lovelace"');
	expect(body).toContain('w:storeItemID="{6C3C8BC8-F283-45AE-878A-BAB7291924A1}"');
	const glossary = await zip.file('word/glossary/document.xml')!.async('string');
	expect(glossary).toContain('w:val="Review line"');
	expect(glossary).toContain('<w:gallery w:val="docParts"/>');
	const headerPart = Object.keys(zip.files).find((name) => /^word\/header\d+\.xml$/.test(name))!;
	expect(await zip.file(headerPart)!.async('string')).toContain('<w:alias w:val="Title"/>');
});

test('Online Video appears only when the host names a video site and refuses others', async ({
	page,
}) => {
	await page.goto('/?framework=vanilla');
	await newDocument(page);
	const editor = page.locator('docx-editor');
	const surface = editor.locator('.ProseMirror');
	await surface.click();
	await editor.getByRole('tab', { name: 'Insert', exact: true }).click();
	const video = editor.getByRole('button', { name: 'Online Video', exact: true });
	await expect(video).toBeHidden();
	await page.evaluate(() => {
		(
			document.querySelector('docx-editor') as unknown as { onlineVideoHosts: string[] }
		).onlineVideoHosts = ['video.example.com'];
	});
	await reveal(editor, video);
	await video.click();
	const dialog = dialogByHeading(editor, 'Insert Video');
	const code = dialog.getByRole('textbox', { name: 'Video embed code or address', exact: true });
	const insert = dialog.getByRole('button', { name: 'Insert', exact: true });
	await code.fill('https://other.example.org/embed/1');
	await expect(dialog).toContainText('not allowed');
	await expect(insert).toBeDisabled();
	await code.fill('<iframe src="https://video.example.com/embed/1"></iframe>');
	await expect(insert).toBeEnabled();
	await insert.click();
	await expect(surface.locator('img:not(.ProseMirror-separator)')).toBeVisible();
	const zip = await savedPackage(page, 'insert-video');
	const body = await zip.file('word/document.xml')!.async('string');
	expect(body).toContain('wp15:webVideoPr');
	expect(body.match(/embeddedHtml=/g)).toHaveLength(1);
	expect(body).toContain('video.example.com/embed/1');
});
