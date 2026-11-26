import { readFile } from 'node:fs/promises';
import { expect, test } from '@playwright/test';
import JSZip from 'jszip';
import type { TextSelection } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { newDocument, reopenDocument, reveal, saveButton } from './helpers';

test('authors and reopens index, citation, bibliography and authority fields from References', async ({
	page,
}, testInfo) => {
	await page.setViewportSize({ width: 1600, height: 900 });
	await page.addInitScript(() => Object.defineProperty(crypto, 'randomUUID', { value: undefined }));
	await page.goto('/?framework=vanilla');
	await newDocument(page);
	const editor = page.locator('docx-editor');
	const surface = editor.locator('.ProseMirror');
	const dialog = editor.locator('.dve-references-dialog');
	const button = (name: string) => dialog.getByRole('button', { name, exact: true });
	const input = (name: string) => dialog.getByRole('textbox', { name, exact: true });
	const command = (name: string) => editor.getByRole('button', { name, exact: true });
	const selectText = (text: string) =>
		editor.evaluate((element, target) => {
			const { view } = element as unknown as { view: EditorView };
			const { state } = view;
			let from: number | undefined;
			state.doc.descendants((node, pos) => {
				const index = node.text?.indexOf(target) ?? -1;
				if (from === undefined && index >= 0) from = pos + index;
				return from === undefined;
			});
			if (from === undefined) throw new Error(`Missing text: ${target}`);
			const selection = state.selection.constructor as typeof TextSelection;
			view.focus();
			view.dispatch(state.tr.setSelection(selection.create(state.doc, from, from + target.length)));
		}, text);
	const run = async (name: string) => {
		const target = command(name);
		await reveal(editor, target);
		await target.click();
	};

	const structure = () =>
		editor.evaluate((element) =>
			(element as unknown as { view: EditorView }).view.state.doc.toJSON(),
		);
	const expectStableUpdate = async (name: string) => {
		const inserted = await structure();
		await run(name);
		expect(await structure()).toEqual(inserted);
		await run(name);
		expect(await structure()).toEqual(inserted);
	};

	await surface.click();
	await page.keyboard.insertText('Zebra meets Zebra and Doe v. City.');
	await page.keyboard.press('Enter');
	await page.keyboard.insertText('zebra ZEBRA Zebras');
	await editor.getByRole('tab', { name: 'References', exact: true }).click();
	await selectText('Zebra');
	await reveal(editor, command('Mark Entry'));
	await command('Mark Entry').focus();
	await page.keyboard.press('Enter');
	await expect(input('Main entry')).toHaveValue('Zebra');
	await dialog.getByRole('checkbox', { name: 'Bold page numbers', exact: true }).check();
	await button('Mark All').click();
	await expect(command('Mark Entry')).toBeFocused();

	await selectText('Doe v. City');
	await run('Mark Citation');
	await expect(input('Long citation')).toHaveValue('Doe v. City');
	await button('Mark').click();

	await run('Manage Sources');
	await input('Tag').fill('Doe2020');
	await input('Title').fill('Field studies');
	await input('Author').fill('Doe, Jane');
	await input('Year').fill('2020');
	await input('Publisher').fill('Example Press');
	await button('Save Source').click();
	await expect(dialog).toHaveCount(0);
	await run('Insert Citation');
	await input('Pages').fill('12');
	await button('Insert').click();
	await expect(surface).toContainText('(Doe, 2020, p. 12)');

	await selectText('Doe, 2020');
	await run('Edit Citation');
	await input('Pages').fill('13');
	await button('OK').click();
	await expect(surface).toContainText('(Doe, 2020, p. 13)');
	await surface.press('Control+z');
	await expect(surface).toContainText('(Doe, 2020, p. 12)');
	await surface.press('Control+y');
	await expect(surface).toContainText('(Doe, 2020, p. 13)');

	await surface.press('Control+End');
	await page.keyboard.press('Enter');
	await run('Insert Index');
	await dialog.getByRole('spinbutton', { name: 'Columns', exact: true }).fill('2');
	await button('OK').click();
	await expect(surface.locator('p').filter({ hasText: /^Zebra\s+1$/ })).toHaveCount(1);
	await expectStableUpdate('Update Index');

	await surface.press('Control+End');
	await run('Bibliography');
	await expect(surface).toContainText('Doe, J. (2020). Field studies. Example Press.');
	await expectStableUpdate('Update Bibliography');
	await surface.press('Control+End');
	await page.keyboard.press('Enter');
	await run('Insert Table of Authorities');
	await button('OK').click();
	await expect(surface.locator('p').filter({ hasText: /^Cases$/ })).toHaveCount(1);
	await expect(surface.locator('p').filter({ hasText: /^Doe v\. City\s+1$/ })).toHaveCount(1);
	await expectStableUpdate('Update Table of Authorities');

	await run('Citation Style');
	await dialog.getByRole('combobox', { name: 'Citation Style' }).selectOption('MLA');
	await button('OK').click();
	await expect(surface).toContainText('(Doe 13)');
	await surface.press('Control+z');
	await expect(surface).toContainText('(Doe, 2020, p. 13)');

	const downloaded = page.waitForEvent('download');
	await saveButton(page).click();
	const file = await downloaded;
	await file.saveAs(testInfo.outputPath('word-references.docx'));
	const bytes = await readFile(testInfo.outputPath('word-references.docx'));
	const zip = await JSZip.loadAsync(bytes);
	const xml = await zip.file('word/document.xml')?.async('string');
	expect(xml?.match(/ XE /g)).toHaveLength(1);
	expect(xml).toContain(' INDEX ');
	expect(xml).toContain(' CITATION ');
	expect(xml).toContain(' BIBLIOGRAPHY ');
	expect(xml).toContain(' TOA ');
	expect(xml).toContain(' TA ');
	expect(xml).toMatch(/<w:cols[^>]*w:num="2"/);
	expect(await zip.file('customXml/item1.xml')?.async('string')).toContain(
		'<b:Tag>Doe2020</b:Tag>',
	);
	await reopenDocument(page, 'word-references.docx', bytes);
	await expect(surface).toContainText('(Doe, 2020, p. 13)');
	await run('Manage Sources');
	await dialog.getByRole('combobox', { name: 'Source', exact: true }).selectOption('Doe2020');
	await expect(input('Title')).toHaveValue('Field studies');
});
