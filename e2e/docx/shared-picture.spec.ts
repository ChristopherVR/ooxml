import JSZip from 'jszip';
import { expect, test } from '@playwright/test';
import { readFileSync } from 'node:fs';
import { withDiagramPicture } from '../../src/core/diagram/test-support/picture-drawing';
import { fileInput } from './helpers';
const bytes = await withDiagramPicture(
	readFileSync(new URL('./support/smartart.docx', import.meta.url)),
	'word',
);
for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid']) {
	test(`${framework}: resolves and paints cropped cached SmartArt pictures`, async ({ page }) => {
		await page.goto(`/?framework=${framework}`);
		await (
			await fileInput(page)
		).setInputFiles({
			name: 'pictures.docx',
			mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
			buffer: Buffer.from(bytes),
		});
		const art = page.locator('docx-editor .ProseMirror office-ui-smartart').first();
		await expect(art.locator('pattern image')).toHaveAttribute('width', '2');
		await expect(art.locator('pattern image')).toHaveAttribute('x', '-0.5');
		await expect(art.locator('feFuncA')).toHaveAttribute('slope', '0.5');
		const width = await art.locator('pattern image').evaluate(async (element) => {
			const image = new Image();
			image.src = element.getAttribute('href')!;
			await image.decode();
			return image.naturalWidth;
		});
		expect(width).toBe(2);
		const saved = await page
			.locator('docx-editor')
			.evaluate(async (element) =>
				Array.from(await (element as unknown as { saveBytes(): Promise<Uint8Array> }).saveBytes()),
			);
		await (
			await fileInput(page)
		).setInputFiles({
			name: 'saved.docx',
			mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
			buffer: Buffer.from(saved),
		});
		await expect(
			page.locator('docx-editor .ProseMirror office-ui-smartart').first().locator('pattern image'),
		).toHaveAttribute('width', '2');
	});
}

for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid']) {
	test(`${framework}: resizes a guide-driven Word custom shape and saves its source`, async ({
		page,
	}) => {
		const zip = await JSZip.loadAsync(
			readFileSync(new URL('./support/shapes.docx', import.meta.url)),
		);
		const source = await zip.file('word/document.xml')!.async('string');
		const custom =
			'<a:custGeom><a:avLst/><a:gdLst><a:gd name="quarter" fmla="*/ w 1 4"/></a:gdLst><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="w" b="h"/><a:pathLst><a:path><a:moveTo><a:pt x="quarter" y="0"/></a:moveTo><a:lnTo><a:pt x="w" y="h"/></a:lnTo><a:lnTo><a:pt x="0" y="h"/></a:lnTo><a:close/></a:path></a:pathLst></a:custGeom>';
		zip.file(
			'word/document.xml',
			source.replace(/<a:prstGeom\b[^>]*>[\s\S]*?<\/a:prstGeom>/, custom),
		);
		await page.goto(`/?framework=${framework}`);
		await (
			await fileInput(page)
		).setInputFiles({
			name: 'custom.docx',
			mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
			buffer: Buffer.from(await zip.generateAsync({ type: 'uint8array' })),
		});
		const editor = page.locator('docx-editor');
		const shape = editor.locator('.ProseMirror .dve-shape[data-shape-preset="custom"]');
		await expect(shape.locator('svg path')).toHaveCount(1);
		await shape.click();
		const width = editor.getByRole('spinbutton', { name: 'Shape width (inches)', exact: true });
		await width.fill('3');
		await width.dispatchEvent('change');
		await expect(shape.locator('svg path')).toHaveAttribute('d', /^M72 0 L288 /);
		const saved = await editor.evaluate(async (element) =>
			Array.from(await (element as unknown as { saveBytes(): Promise<Uint8Array> }).saveBytes()),
		);
		const reopened = await JSZip.loadAsync(Buffer.from(saved));
		expect(await reopened.file('word/document.xml')!.async('string')).toContain('fmla="*/ w 1 4"');
	});
}
