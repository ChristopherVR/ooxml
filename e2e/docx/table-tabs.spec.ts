import { expect, test } from '@playwright/test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import {
	dialogByHeading,
	insertTableOfSize,
	newDocument,
	reopenDocument,
	reveal,
	saveButton,
} from './helpers';

test('a table without tblLook shows the options Word uses on open', async ({ page }) => {
	await page.goto('/?framework=vanilla');
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
		<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>Imported</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/>
		</w:body></w:document>`,
	);
	await reopenDocument(page, 'no-table-look.docx', await zip.generateAsync({ type: 'nodebuffer' }));
	const editor = page.locator('docx-editor');
	await editor.locator('.ProseMirror td').first().click();
	await editor.getByRole('tab', { name: 'Table Design', exact: true }).click();
	for (const name of ['Header Row', 'First Column', 'Banded Rows'])
		await expect(editor.getByRole('button', { name, exact: true })).toHaveAttribute(
			'aria-pressed',
			'true',
		);
	for (const name of ['Total Row', 'Last Column', 'Banded Columns'])
		await expect(editor.getByRole('button', { name, exact: true })).toHaveAttribute(
			'aria-pressed',
			'false',
		);
});

test('typing after Insert Row Below stays in the inserted row', async ({ page }) => {
	await page.goto('/?framework=vanilla');
	await newDocument(page);
	await insertTableOfSize(page);
	const editor = page.locator('docx-editor');
	const rows = editor.locator('.ProseMirror tr');
	await rows.nth(0).locator('td').first().click();
	await page.keyboard.type('Heading');
	await rows.nth(1).locator('td').first().click();
	await page.keyboard.type('Original');
	await rows.nth(0).locator('td').last().click();
	await editor.getByRole('tab', { name: 'Table Layout', exact: true }).click();
	await editor.getByRole('button', { name: 'Insert row below', exact: true }).click();
	await page.keyboard.type('Inserted');
	await expect(rows).toHaveCount(3);
	await expect(rows.nth(1).locator('td').first()).toHaveText('Inserted');
	await editor.getByRole('button', { name: 'Delete row', exact: true }).click();
	await expect(rows).toHaveCount(2);
	await expect(rows.nth(0)).toHaveText('Heading');
	await expect(rows.nth(1)).toHaveText('Original');
});

test('Table Design and Layout edit a table and save the Word properties', async ({ page }) => {
	await page.goto('/?framework=vanilla');
	await newDocument(page);
	const editor = page.locator('docx-editor');
	const design = editor.getByRole('tab', { name: 'Table Design', exact: true });
	const layout = editor.getByRole('tab', { name: 'Table Layout', exact: true });
	await expect(design).toHaveCount(0);
	await expect(layout).toHaveCount(0);
	await insertTableOfSize(page);
	const cells = editor.locator('.ProseMirror td');
	await cells.first().click();
	await design.click();
	await editor.getByRole('button', { name: 'Table Styles', exact: true }).click();
	const stylesList = editor.getByRole('listbox', { name: 'Table Styles', exact: true });
	await expect(stylesList.getByRole('option', { name: 'Normal Table', exact: true })).toBeFocused();
	await page.keyboard.press('ArrowRight');
	await expect(
		stylesList.getByRole('option', { name: 'Plain Table 1', exact: true }),
	).toBeFocused();
	await expect(stylesList.locator('[tabindex="0"]')).toHaveCount(1);
	await page.keyboard.press('ArrowDown');
	await page.keyboard.press('ArrowDown');
	await expect(
		stylesList.getByRole('option', { name: 'Grid Table 4 - Accent 1', exact: true }),
	).toBeFocused();
	await page.keyboard.press('Enter');
	await expect(cells.first()).toHaveCSS('background-color', 'rgb(68, 114, 196)');
	const headerStyle = editor.getByRole('button', { name: 'Header Row', exact: true });
	await headerStyle.click();
	await expect(headerStyle).toHaveAttribute('aria-pressed', 'false');
	await headerStyle.click();
	await expect(headerStyle).toHaveAttribute('aria-pressed', 'true');
	await layout.click();
	const repeat = editor.getByRole('button', { name: 'Repeat Header Rows', exact: true });
	await reveal(editor, repeat);
	await repeat.click();
	const align = editor.getByRole('button', { name: 'Align Bottom Right', exact: true });
	await reveal(editor, align);
	await align.click();
	await expect(cells.first()).toHaveCSS('vertical-align', 'bottom');
	await expect(cells.first().locator('p')).toHaveCSS('text-align', 'right');
	const autoFit = editor.getByRole('combobox', { name: 'AutoFit', exact: true });
	await reveal(editor, autoFit);
	await autoFit.selectOption('window');
	await expect(editor.locator('.ProseMirror table')).toHaveAttribute('style', /width:\s*100%/);
	const downloadPromise = page.waitForEvent('download');
	await saveButton(page).click();
	const download = await downloadPromise;
	await download.saveAs('test-results/table-tabs.docx');
	const zip = await JSZip.loadAsync(await readFile((await download.path())!));
	const xml = await zip.file('word/document.xml')?.async('string');
	const styles = await zip.file('word/styles.xml')?.async('string');
	expect(xml).toContain('w:val="GridTable4-Accent1"');
	expect(xml).toContain('w:firstRow="1"');
	expect(xml).toContain('<w:tblHeader');
	expect(xml).toContain('<w:vAlign w:val="bottom"');
	expect(xml).toContain('<w:tblLayout w:type="autofit"');
	expect(styles).toContain('w:styleId="GridTable4-Accent1"');
	expect(styles).toContain('w:themeFill="accent1"');
	expect(styles).toContain('w:themeColor="accent1"');
});

test('table commands without revision support are disabled during Track Changes', async ({
	page,
}) => {
	await page.goto('/?framework=vanilla');
	await newDocument(page);
	await insertTableOfSize(page);
	const editor = page.locator('docx-editor');
	await editor.locator('.ProseMirror td').first().click();
	await editor.getByRole('tab', { name: 'Review', exact: true }).click();
	await editor.getByRole('button', { name: 'Track changes', exact: true }).click();
	await editor.getByRole('tab', { name: 'Table Design', exact: true }).click();
	const reason = 'This command is unavailable while Track Changes is on.';
	for (const name of ['Table Styles', 'Header Row']) {
		const control = editor.getByRole('button', { name, exact: true });
		await expect(control).toBeDisabled();
		await expect(control).toHaveAccessibleDescription(reason);
	}
	await editor.getByRole('tab', { name: 'Table Layout', exact: true }).click();
	for (const name of ['Sort Table', 'Convert to Text', 'AutoFit', 'Row Height (inches)']) {
		const control = editor.getByLabel(name, { exact: true });
		await expect(control).toBeDisabled();
		await expect(control).toHaveAccessibleDescription(reason);
	}
});

test('Table Layout selects a column, evaluates a formula and converts with undo', async ({
	page,
}) => {
	await page.goto('/?framework=vanilla');
	await newDocument(page);
	await insertTableOfSize(page);
	const editor = page.locator('docx-editor');
	const cells = editor.locator('.ProseMirror td');
	await cells.first().click();
	await page.keyboard.type('14');
	await editor.getByRole('tab', { name: 'Table Layout', exact: true }).click();
	await editor
		.getByRole('combobox', { name: 'Select Table Parts', exact: true })
		.selectOption('column');
	await expect(editor.locator('.dve-selected-table-cell')).toHaveCount(2);
	await cells.nth(2).click();
	const repeat = editor.getByRole('button', { name: 'Repeat Header Rows', exact: true });
	await reveal(editor, repeat);
	await expect(
		editor.getByRole('button', { name: 'Repeat Header Rows', exact: true }),
	).toBeDisabled();
	await editor.getByRole('button', { name: 'Formula', exact: true }).click();
	const formula = dialogByHeading(editor, 'Formula');
	await expect(formula.getByRole('textbox', { name: 'Formula', exact: true })).toHaveValue(
		'=SUM(ABOVE)',
	);
	await formula.getByRole('button', { name: 'OK', exact: true }).click();
	await expect(cells.nth(2)).toHaveText('14');
	const convertButton = editor.getByRole('button', { name: 'Convert to Text', exact: true });
	await reveal(editor, convertButton);
	await convertButton.click();
	const convert = dialogByHeading(editor, 'Convert to Text');
	await convert
		.getByRole('combobox', { name: 'Separate text with', exact: true })
		.selectOption('commas');
	await convert.getByRole('button', { name: 'OK', exact: true }).click();
	await expect(cells).toHaveCount(0);
	await editor.locator('.ProseMirror').focus();
	await page.keyboard.press('Control+z');
	await expect(cells).toHaveCount(4);
});

test('Home bold and alignment apply to all and only the selected column', async ({ page }) => {
	await page.goto('/?framework=vanilla');
	await newDocument(page);
	await insertTableOfSize(page);
	const editor = page.locator('docx-editor');
	const cells = editor.locator('.ProseMirror td');
	for (const [index, text] of ['First', 'Neighbor', 'Second', 'Other'].entries()) {
		await cells.nth(index).click();
		await page.keyboard.type(text);
	}
	await cells.first().click();
	await editor.getByRole('tab', { name: 'Table Layout', exact: true }).click();
	await editor.getByRole('combobox', { name: 'Select Table Parts' }).selectOption('column');
	await editor.getByRole('tab', { name: 'Home', exact: true }).click();
	await editor.getByRole('button', { name: 'Bold', exact: true }).click();
	await editor.getByRole('button', { name: 'Align right', exact: true }).click();
	for (const index of [0, 2]) {
		await expect(cells.nth(index).locator('strong')).toBeVisible();
		await expect(cells.nth(index).locator('p')).toHaveCSS('text-align', 'right');
	}
	for (const index of [1, 3]) {
		await expect(cells.nth(index).locator('strong')).toHaveCount(0);
		await expect(cells.nth(index).locator('p')).not.toHaveCSS('text-align', 'right');
	}
});
