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
