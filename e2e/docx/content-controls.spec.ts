import { expect, test } from '@playwright/test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { fileInput, newDocument } from './helpers';

const fixture = new URL('./support/content-controls.docx', import.meta.url);

async function openFixture(page: import('@playwright/test').Page) {
	await page.goto('/?framework=vanilla');
	await (
		await fileInput(page)
	).setInputFiles({
		name: 'content-controls.docx',
		mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
		buffer: await readFile(fixture),
	});
	const editor = page.locator('docx-editor');
	await expect(editor.locator('.ProseMirror .dve-sdt')).toHaveCount(5);
	return editor;
}
const savedDocument = async (page: import('@playwright/test').Page) => {
	const bytes = await page
		.locator('docx-editor')
		.evaluate(async (element) =>
			Array.from(await (element as unknown as { saveBytes(): Promise<Uint8Array> }).saveBytes()),
		);
	const zip = await JSZip.loadAsync(new Uint8Array(bytes));
	return zip.file('word/document.xml')!.async('string');
};

test('check box, drop-down list and date picker controls are operated and saved', async ({
	page,
}) => {
	const editor = await openFixture(page);
	const box = editor.locator('.dve-sdt[data-sdt-kind="checkbox"]');
	await expect(box).toHaveText('☒');
	await box.click();
	await expect(box).toHaveText('☐');

	const drop = editor.locator('.dve-sdt[data-sdt-kind="dropDownList"]');
	await expect(drop).toHaveText('Green');
	await drop.click();
	await editor.getByRole('option', { name: 'Blue', exact: true }).click();
	await expect(drop).toHaveText('Blue');

	const date = editor.locator('.dve-sdt[data-sdt-kind="date"]');
	await expect(date).toHaveAttribute('data-sdt-placeholder', '1');
	await date.click();
	const calendar = editor.locator('.sdt-calendar');
	await expect(calendar).toBeVisible();
	await calendar.locator('button[data-date]:not(.outside)').nth(14).click();
	await expect(date).not.toHaveAttribute('data-sdt-placeholder', '1');
	await expect(date).toHaveText(/^\d{2} \w+ \d{4}$/);

	const xml = await savedDocument(page);
	expect(xml).toContain('<w14:checked w14:val="0"/>');
	expect(xml).toContain('<w:t>Blue</w:t>');
	expect(xml).toMatch(/w:fullDate="\d{4}-\d{2}-15T00:00:00Z"/);
});

test('a content-locked block refuses typing and the title tab shows on hover', async ({ page }) => {
	const editor = await openFixture(page);
	const block = editor.locator('p[data-sdt-block]');
	await expect(block).toHaveAttribute('data-sdt-title', 'Notes block');
	await block.click();
	await page.keyboard.type('refused');
	await expect(block).toHaveText('Notes');
	await block.hover();
	const tab = await block.evaluate((element) => getComputedStyle(element, '::before').content);
	expect(tab).toContain('Notes block');
});

test('Insert > Content control adds a placeholder control that typing replaces', async ({
	page,
}) => {
	await page.setViewportSize({ width: 1700, height: 720 });
	await page.goto('/?framework=vanilla');
	await newDocument(page);
	const editor = page.locator('docx-editor');
	await editor.locator('.ProseMirror').click();
	await page.keyboard.type('Name: ');
	await editor.getByRole('tab', { name: 'Insert', exact: true }).click();
	await editor.getByRole('combobox', { name: 'Content control', exact: true }).selectOption('text');
	const control = editor.locator('.ProseMirror .dve-sdt[data-sdt-kind="text"]');
	await expect(control).toHaveAttribute('data-sdt-placeholder', '1');
	await page.keyboard.type('Ada');
	await expect(control).toHaveText('Ada');
	await expect(control).not.toHaveAttribute('data-sdt-placeholder', '1');
	const xml = await savedDocument(page);
	expect(xml).toMatch(/<w:sdt><w:sdtPr><w:id w:val="\d+"\/><w:text\/><\/w:sdtPr>/);
	expect(xml).toContain('<w:t>Ada</w:t>');
});
