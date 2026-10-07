import { test, expect } from '@playwright/test';
import JSZip from 'jszip';
import { fileInput } from './helpers';
import type { DocxEditorElement } from '../../viewers/docx/packages/web-component/src';

for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid'])
	test(`${framework}: preserves opaque run properties through tracked typing and history`, async ({
		page,
	}) => {
		const zip = new JSZip();
		zip.file(
			'word/document.xml',
			'<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:x="urn:browser-properties"><w:body><w:p><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:hint="eastAsia" w:asciiTheme="minorAscii"/><x:property x:value="retained"/></w:rPr><w:t>Text</w:t></w:r></w:p><w:sectPr/></w:body></w:document>',
		);
		const bytes = await zip.generateAsync({ type: 'nodebuffer' });
		await page.setViewportSize({ width: 2400, height: 1000 });
		const errors: string[] = [];
		page.on('pageerror', (error) => errors.push(error.message));
		await page.goto(`/?framework=${framework}`);
		await (
			await fileInput(page)
		).setInputFiles({
			name: 'opaque-properties.docx',
			mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
			buffer: bytes,
		});
		const editor = page.locator('docx-editor');
		const body = editor.locator('.ProseMirror');
		await expect(body).toContainText('Text');
		await editor.getByRole('tab', { name: 'Review', exact: true }).click();
		await editor.getByRole('button', { name: 'Track changes', exact: true }).click();
		await body.click();
		await page.keyboard.press('Control+Home');
		await page.keyboard.press('ArrowRight');
		await page.keyboard.press('ArrowRight');
		await page.keyboard.type('!');
		await expect(body).toContainText('Te!xt');
		const exportXml = async () => {
			const data = await editor.evaluate(async (element) =>
				Array.from(await (element as DocxEditorElement).saveBytes()),
			);
			return (await JSZip.loadAsync(new Uint8Array(data)))
				.file('word/document.xml')!
				.async('string');
		};
		let xml = await exportXml();
		expect(xml.match(/x:value="retained"/g)).toHaveLength(3);
		expect(xml.match(/w:hint="eastAsia"/g)).toHaveLength(3);
		expect(xml).toContain('<w:ins ');
		await page.keyboard.press('Control+z');
		await expect(body).toHaveText('Text');
		await page.keyboard.press('Control+Shift+z');
		await expect(body).toHaveText('Te!xt');
		xml = await exportXml();
		expect(xml.match(/x:value="retained"/g)).toHaveLength(3);
		expect(errors).toEqual([]);
	});
