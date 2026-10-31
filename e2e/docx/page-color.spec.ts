import { expect, test } from '@playwright/test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { newDocument, reveal, saveButton, fileInput } from './helpers';
import { createDocument, saveDocx } from '../../viewers/docx/packages/core/src';
import type { DocxEditorElement } from '../../viewers/docx/packages/web-component/src';

async function themedPage(): Promise<Buffer> {
	const zip = await JSZip.loadAsync(
		await saveDocx({
			...createDocument(),
			blocks: [{ type: 'paragraph', id: 'theme', runs: [{ text: 'Theme page' }] }],
		}),
	);
	const path = 'word/document.xml';
	zip.file(
		path,
		(await zip.file(path)!.async('string')).replace(
			'<w:body>',
			'<w:background w:color="FF0000" w:themeColor="accent2" w:themeTint="7F"/><w:body>',
		),
	);
	zip.file(
		'word/theme/theme1.xml',
		await readFile(new URL('./support/word-theme-native.xml', import.meta.url)),
	);
	zip.file(
		'word/_rels/document.xml.rels',
		(
			(await zip.file('word/_rels/document.xml.rels')?.async('string')) ??
			'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'
		).replace(
			'</Relationships>',
			'<Relationship Id="rIdNativeTheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/></Relationships>',
		),
	);
	zip.file(
		'[Content_Types].xml',
		(await zip.file('[Content_Types].xml')!.async('string')).replace(
			'</Types>',
			'<Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/></Types>',
		),
	);
	return zip.generateAsync({ type: 'nodebuffer' });
}

for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid']) {
	test(`${framework}: native theme page colour overrides its cache and survives text edits`, async ({
		page,
	}) => {
		await page.goto(`/?framework=${framework}`);
		await (
			await fileInput(page)
		).setInputFiles({
			name: 'themed.docx',
			mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
			buffer: await themedPage(),
		});
		const editor = page.locator('docx-editor');
		const paper = editor.getByLabel('Document page', { exact: true });
		await expect(paper).toHaveCSS('background-color', 'rgb(246, 190, 152)');
		await editor.locator('.ProseMirror').click();
		await page.keyboard.press('Control+End');
		await page.keyboard.type(' edited');
		await expect(paper).toHaveCSS('background-color', 'rgb(246, 190, 152)');
		const bytes = await editor.evaluate(async (element) =>
			Array.from(new Uint8Array(await (await (element as DocxEditorElement).save()).arrayBuffer())),
		);
		const zip = await JSZip.loadAsync(new Uint8Array(bytes));
		const xml = await zip.file('word/document.xml')!.async('string');
		expect(xml).toContain('w:themeColor="accent2"');
		expect(xml).toContain('w:themeTint="7F"');
		expect(xml).toContain('w:color="FF0000"');
		await editor.getByRole('tab', { name: 'Design', exact: true }).click();
		const caret = editor.locator('[aria-label="Page color"] + [data-split-caret]');
		await reveal(editor, caret);
		await caret.click();
		await editor.getByRole('menuitemradio', { name: 'Red', exact: true }).click();
		await expect(paper).toHaveCSS('background-color', 'rgb(255, 0, 0)');
		const redBytes = await editor.evaluate(async (element) =>
			Array.from(new Uint8Array(await (await (element as DocxEditorElement).save()).arrayBuffer())),
		);
		const redZip = await JSZip.loadAsync(new Uint8Array(redBytes));
		expect(await redZip.file('word/document.xml')!.async('string')).toContain(
			'<w:background w:color="FF0000"/>',
		);
	});
}

test('Design > Page Color paints the page, saves w:background and undoes', async ({ page }) => {
	await page.setViewportSize({ width: 1700, height: 800 });
	await page.goto('/?framework=vanilla');
	await newDocument(page);
	const editor = page.locator('docx-editor');
	const paper = editor.getByLabel('Document page', { exact: true });
	await editor.locator('.ProseMirror').click();
	await page.keyboard.type('Tinted page');
	await editor.getByRole('tab', { name: 'Design', exact: true }).click();
	const caret = editor.locator('[aria-label="Page color"] + [data-split-caret]');
	await reveal(editor, caret);
	await caret.click();
	await editor.getByRole('menuitemradio', { name: 'Light Green', exact: true }).click();
	await expect(paper).toHaveCSS('background-color', 'rgb(146, 208, 80)');

	const pending = page.waitForEvent('download');
	await saveButton(page).click();
	const zip = await JSZip.loadAsync(await readFile((await (await pending).path())!));
	expect(await zip.file('word/document.xml')!.async('string')).toContain(
		'<w:background w:color="92D050"/>',
	);
	expect(await zip.file('word/settings.xml')!.async('string')).toContain(
		'<w:displayBackgroundShape/>',
	);

	await editor.locator('.ProseMirror').click();
	await page.keyboard.press('Control+z');
	await expect(paper).toHaveCSS('background-color', 'rgb(255, 255, 255)');
});

test('Layout > Hyphenation turns automatic hyphenation on, saves it and undoes', async ({
	page,
}) => {
	await page.setViewportSize({ width: 1700, height: 800 });
	await page.goto('/?framework=vanilla');
	await newDocument(page);
	const editor = page.locator('docx-editor');
	const paper = editor.getByLabel('Document page', { exact: true });
	await editor.locator('.ProseMirror').click();
	await page.keyboard.type('Hyphenate me');
	await editor.locator('[role="tab"][data-tab="layout"]').click();
	const select = editor.getByLabel('Hyphenation', { exact: true });
	await reveal(editor, select);
	await select.selectOption('auto');
	await expect(paper).toHaveCSS('hyphens', 'auto');

	const pending = page.waitForEvent('download');
	await saveButton(page).click();
	const zip = await JSZip.loadAsync(await readFile((await (await pending).path())!));
	expect(await zip.file('word/settings.xml')!.async('string')).toContain('<w:autoHyphenation/>');

	await editor.locator('.ProseMirror').click();
	await page.keyboard.press('Control+z');
	await expect(paper).toHaveCSS('hyphens', 'manual');
});
