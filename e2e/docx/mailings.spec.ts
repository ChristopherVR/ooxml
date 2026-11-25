import { readFile } from 'node:fs/promises';
import { expect, test } from '@playwright/test';
import JSZip from 'jszip';
import { newDocument, reopenDocument } from './helpers';

test('merge a local CSV through the ribbon, preview and print it, and reopen the download', async ({
	page,
}) => {
	await page.setViewportSize({ width: 1600, height: 900 });
	await page.goto('/?framework=vanilla');
	await newDocument(page);
	const editor = page.locator('docx-editor');
	const ribbon = editor.locator('office-ui-ribbon');
	const command = (name: string) => ribbon.getByRole('button', { name, exact: true });
	const modal = editor.locator('.dve-mailings-dialog');
	const button = (name: string) => modal.getByRole('button', { name, exact: true });
	await ribbon.getByRole('tab', { name: 'Mailings' }).click();
	await expect(command('Preview Results')).toBeDisabled();
	await expect(command('Envelopes')).toBeDisabled();
	await command('Start Mail Merge').focus();
	await page.keyboard.press('Enter');
	await modal.getByRole('combobox', { name: 'Start Mail Merge' }).selectOption('formLetters');
	await button('OK').click();
	await expect(command('Start Mail Merge')).toBeFocused();
	await command('Select Recipients').click();
	await expect(button('Use host recipients')).toBeDisabled();
	const parser = page.waitForEvent('worker', (worker) =>
		worker.url().includes('mailings-recipient-worker'),
	);
	const chooser = page.waitForEvent('filechooser');
	await button('Choose a CSV, TSV, or XLSX file').click();
	await (
		await chooser
	).setFiles({
		name: 'recipients.csv',
		mimeType: 'text/csv',
		buffer: Buffer.from('Name,Group\nZoe,Current\nAna,Former\nCal,Current'),
	});
	await parser;
	await expect(command('Insert Merge Field')).toBeEnabled();
	await command('Edit Recipient List').click();
	await modal.getByRole('checkbox', { name: 'Include: Record 2' }).uncheck();
	await modal.getByRole('combobox', { name: 'Sort by' }).selectOption('Name');
	await button('OK').click();
	await command('Insert Merge Field').click();
	await button('OK').click();
	await expect(editor.locator('.ProseMirror')).toContainText('«Name»');
	await command('Highlight Merge Fields').click();
	await expect(command('Highlight Merge Fields')).toHaveAttribute('aria-pressed', 'true');

	const projection = editor.locator('[data-merge-projection]');
	await command('Preview Results').click();
	await expect(projection).toContainText('Cal');
	await expect(editor.locator('.ProseMirror')).toBeHidden();
	await command('Last recipient').click();
	await expect(projection).toContainText('Zoe');
	await command('Previous recipient').click();
	await expect(projection).toContainText('Cal');
	await command('Next recipient').click();
	await expect(projection).toContainText('Zoe');
	await command('First recipient').click();
	await expect(projection).toContainText('Cal');
	await command('Find Recipient').click();
	await modal.getByRole('textbox', { name: 'Search recipient values' }).fill('zoe');
	await button('Find next').click();
	await expect(projection).toContainText('Zoe');
	await command('Check for Errors').click();
	await expect(modal).toContainText('No merge errors found.');
	await page.keyboard.press('Escape');

	await page.evaluate(() => {
		window.print = () => {
			const host = document.querySelector('docx-editor');
			const pages = host?.shadowRoot?.querySelector('[data-merge-projection]');
			document.documentElement.dataset.printedMerge = pages?.textContent ?? '';
		};
	});
	await command('Finish & Merge').click();
	await expect(button('Send Email Messages')).toBeDisabled();
	await button('Print Documents').click();
	await expect(page.locator('html')).toHaveAttribute('data-printed-merge', /Cal.*Zoe/);
	await expect(projection).toContainText('Zoe');
	await command('Finish & Merge').click();
	const download = page.waitForEvent('download');
	await button('Edit Individual Documents').click();
	const file = await download;
	await command('Preview Results').click();
	await expect(editor.locator('.ProseMirror')).toContainText('«Name»');
	const path = await file.path();
	if (!path) throw new Error('The merge did not download a DOCX.');
	await reopenDocument(page, file.suggestedFilename(), await readFile(path));
	await expect(editor.locator('.ProseMirror')).toContainText('Cal');
	await expect(editor.locator('.ProseMirror')).toContainText('Zoe');
	await expect(editor.locator('.ProseMirror')).not.toContainText('«Name»');
	await expect(editor.locator('.ProseMirror')).not.toContainText('Ana');
});

test('reads the first XLSX sheet and uses cached values without calculating formulas', async ({
	page,
}) => {
	const zip = new JSZip();
	const spreadsheet = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
	const relationships = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
	zip.file(
		'xl/workbook.xml',
		`<workbook xmlns="${spreadsheet}" xmlns:r="${relationships}"><sheets><sheet name="Recipients" sheetId="1" r:id="rId1"/></sheets></workbook>`,
	);
	zip.file(
		'xl/_rels/workbook.xml.rels',
		`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${relationships}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
	);
	zip.file(
		'xl/worksheets/sheet1.xml',
		`<worksheet xmlns="${spreadsheet}"><sheetData>
			<row r="1"><c r="A1" t="inlineStr"><is><t>Amount</t></is></c></row>
			<row r="2"><c r="A2"><f>40+2</f><v>71</v></c></row>
		</sheetData></worksheet>`,
	);
	await page.setViewportSize({ width: 1600, height: 900 });
	await page.goto('/?framework=vanilla');
	await newDocument(page);
	const editor = page.locator('docx-editor');
	const ribbon = editor.locator('office-ui-ribbon');
	const dialog = editor.locator('.dve-mailings-dialog');
	await ribbon.getByRole('tab', { name: 'Mailings' }).click();
	await ribbon.getByRole('button', { name: 'Select Recipients' }).click();
	await dialog.locator('input[type=file]').setInputFiles({
		name: 'recipients.xlsx',
		mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
		buffer: await zip.generateAsync({ type: 'nodebuffer' }),
	});
	await ribbon.getByRole('button', { name: 'Insert Merge Field' }).click();
	await dialog.getByRole('button', { name: 'OK' }).click();
	await ribbon.getByRole('button', { name: 'Preview Results' }).click();
	await expect(editor.locator('[data-merge-projection]')).toContainText('71');
});
