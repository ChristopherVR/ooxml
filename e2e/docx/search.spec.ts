import { expect, test } from '@playwright/test';
import {
	clickInEditor,
	editorSelectionText,
	openSample,
	newDocument,
	saveButton,
	setReadOnly,
} from './helpers';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';

for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid']) {
	test(`${framework}: shared find/replace respects formatting, history and readonly`, async ({
		page,
	}) => {
		const errors: string[] = [];
		page.on('pageerror', (error) => errors.push(error.message));
		await openSample(page, framework);
		const editor = page.locator('docx-editor');
		const surface = editor.locator('.ProseMirror');
		await expect(surface).toContainText('Document title');
		await newDocument(page);
		await surface.click();
		await page.keyboard.type('alpha ');
		await page.keyboard.press('Control+b');
		await page.keyboard.type('ALPHA');
		await page.keyboard.press('Control+b');
		await page.keyboard.insertText(' café 👩🏽‍💻');
		await page.keyboard.press('Control+f');
		await expect(editor.getByLabel('Find text', { exact: true })).toBeFocused();
		await editor.getByLabel('Find text', { exact: true }).fill('alpha');
		await editor.getByRole('button', { name: 'Find next', exact: true }).click();
		await expect(editor.getByRole('status', { name: 'Search results' })).toContainText('2');
		await editor.getByLabel('Replace with', { exact: true }).fill('$&');
		await editor.getByRole('button', { name: 'Replace all', exact: true }).click();
		await expect(surface).toHaveText('$& $& café 👩🏽‍💻');
		await expect(surface.locator('strong')).toHaveText('$&');
		await editor.getByRole('button', { name: 'Undo', exact: true }).click();
		await expect(surface).toHaveText('alpha ALPHA café 👩🏽‍💻');
		await editor.getByRole('button', { name: 'Redo', exact: true }).click();
		await expect(surface).toHaveText('$& $& café 👩🏽‍💻');
		await setReadOnly(page, true);
		await editor.getByLabel('Find text', { exact: true }).fill('café');
		await expect(editor.getByRole('button', { name: 'Replace all', exact: true })).toBeDisabled();
		await editor.getByRole('button', { name: 'Find next', exact: true }).click();
		await expect(surface).toHaveText('$& $& café 👩🏽‍💻');
		await editor.getByLabel('Find text', { exact: true }).press('Escape');
		await expect(editor.getByLabel('Find text', { exact: true })).toBeHidden();
		const pending = page.waitForEvent('download');
		await saveButton(page).click();
		const result = await pending;
		const zip = await JSZip.loadAsync(await readFile((await result.path())!));
		const xml = await zip.file('word/document.xml')!.async('string');
		expect(xml).toContain('$&amp;');
		expect(xml).toContain('café 👩🏽‍💻');
		expect(errors).toEqual([]);
	});
}

test('find and replace walk the body, headers, notes and text boxes in story order', async ({
	page,
}) => {
	const errors: string[] = [];
	page.on('pageerror', (error) => errors.push(error.message));
	await page.goto('/?framework=vanilla');
	await newDocument(page);
	const editor = page.locator('docx-editor');
	await editor.evaluate((element) => {
		const host = element as HTMLElement & { documentModel: Record<string, unknown> };
		const model = host.documentModel;
		const paragraph = (id: string, text: string) => ({ type: 'paragraph', id, runs: [{ text }] });
		const section = (endsAtBlockId: string, partName: string, text: string) => ({
			type: 'nextPage',
			orientation: 'portrait',
			pageWidthTwips: 12240,
			pageHeightTwips: 15840,
			marginTopTwips: 1440,
			marginLeftTwips: 1440,
			marginRightTwips: 1440,
			marginBottomTwips: 1440,
			columns: { count: 1, equalWidth: true },
			endsAtBlockId,
			headers: { default: { partName, blocks: [paragraph(`${endsAtBlockId}-h`, text)] } },
		});
		host.documentModel = {
			...model,
			blocks: [
				{
					type: 'paragraph',
					id: 'body',
					runs: [
						{ text: 'zq body ' },
						{ text: '', noteReference: { kind: 'footnote', id: '1' } },
						{
							text: '',
							image: {
								relId: '',
								partName: '',
								contentType: '',
								widthPx: 160,
								heightPx: 48,
								unsupported: 'Text box',
								textBoxText: ['zq box'],
								textBoxEditable: true,
								textBoxBorder: true,
							},
						},
					],
				},
				paragraph('body2', 'second body'),
			],
			footnotes: [
				{
					id: '1',
					blocks: [
						{
							type: 'paragraph',
							id: 'note',
							runs: [{ text: '1', noteMark: 'footnote' }, { text: ' zq note' }],
						},
					],
				},
			],
			sections: [
				section('body', 'word/header1.xml', 'zq header'),
				section('body2', 'word/header2.xml', 'zq second'),
			],
		};
	});
	const surface = editor.locator('.dve-paper > .ProseMirror');
	await expect(surface).toContainText('zq body');
	await clickInEditor(surface.locator('p').first(), { position: { x: 2, y: 4 } });
	await expect(surface).toBeFocused();
	await page.keyboard.press('Control+Home');
	await page.keyboard.press('Control+f');
	const find = editor.getByLabel('Find text', { exact: true });
	await find.fill('zq');
	const status = editor.getByRole('status', { name: 'Search results' });
	const where = editor.getByLabel('Search location', { exact: true });
	const headerEditor = editor.locator('.dve-header .ProseMirror');
	const noteEditor = editor.locator('.dve-notes-container .ProseMirror');
	await expect(status).toHaveText('5 matches');
	await find.press('Enter');
	await expect(status).toHaveText('1 of 5 matches');
	await expect(where).toHaveText('Main Document');
	// A header match opens that section's header with the match selected; the focus stays in Find.
	await find.press('Enter');
	await expect(where).toHaveText('Header -Section 1-');
	await expect(headerEditor).toContainText('zq header');
	await expect(find).toBeFocused();
	await find.press('Enter');
	await expect(where).toHaveText('Header -Section 2-');
	await expect(headerEditor).toContainText('zq second');
	await editor.getByRole('button', { name: 'Find next', exact: true }).click();
	await expect(status).toHaveText('4 of 5 matches');
	await expect(where).toHaveText('Footnotes');
	await expect(headerEditor).toHaveCount(0);
	await expect(noteEditor).toContainText('zq note');
	// Replace changes the match in the open note.
	await editor.getByLabel('Replace with', { exact: true }).fill('ok');
	await editor
		.getByRole('region', { name: 'Find and replace' })
		.getByRole('button', { name: 'Replace', exact: true })
		.click();
	await expect(noteEditor).toContainText('ok note');
	await expect(status).toHaveText('4 matches');
	await find.press('Enter');
	await expect(where).toHaveText('Text Boxes');
	await expect(noteEditor).toHaveCount(0);
	await find.press('Enter');
	await expect(where).toHaveText('Main Document');
	// Closing Find leaves the caret in the story it opened, so typing goes there.
	await find.press('Enter');
	await expect(where).toHaveText('Header -Section 1-');
	await find.press('Escape');
	await expect(find).toBeHidden();
	await page.keyboard.type('X');
	await expect(headerEditor).toContainText('X header');
	await expect(surface).not.toContainText('X');

	// The click's caret reaches the body through an asynchronous selectionchange that the editor
	// ignores once Ctrl+F has moved the focus, so wait for it, as a person's pause would.
	await clickInEditor(surface.locator('p').first(), { position: { x: 2, y: 4 } });
	await expect(headerEditor).toHaveCount(0);
	await expect.poll(() => editorSelectionText(editor)).toBe('');
	await page.keyboard.press('Control+f');
	await expect(status).toHaveText('3 matches');
	await editor.getByRole('button', { name: 'Replace all', exact: true }).click();
	await expect(status).toHaveText('0 matches');
	await expect(surface).toContainText('ok body');
	await expect(surface.getByText('ok box', { exact: true })).toBeVisible();
	// One undo step per story: the text boxes, replaced last, come back first.
	await editor.getByRole('button', { name: 'Undo', exact: true }).click();
	await expect(surface.getByText('zq box', { exact: true })).toBeVisible();
	await expect(surface).toContainText('ok body');
	await expect(status).toHaveText('1 match');
	await editor.getByRole('button', { name: 'Redo', exact: true }).click();
	await expect(surface.getByText('ok box', { exact: true })).toBeVisible();

	// From Print Layout, a header match opens the header on the continuous (Draft) surface.
	await editor.getByRole('button', { name: 'Print Layout', exact: true }).click();
	await page.keyboard.press('Control+f');
	await find.fill('ok second');
	await find.press('Enter');
	await expect(where).toHaveText('Header -Section 2-');
	await expect(headerEditor).toContainText('ok second');

	const pending = page.waitForEvent('download');
	await saveButton(page).click();
	const zip = await JSZip.loadAsync(await readFile((await (await pending).path())!));
	const names = Object.keys(zip.files).filter((name) =>
		/^word\/(document|footnotes|header\d+)\.xml$/.test(name),
	);
	expect(names).toHaveLength(4);
	const xml = (await Promise.all(names.map((name) => zip.file(name)!.async('string')))).join('');
	expect(xml).not.toContain('zq');
	for (const text of ['ok body', 'ok box', 'ok note', 'X header', 'ok second'])
		expect(xml).toContain(text);
	expect(errors).toEqual([]);
});
