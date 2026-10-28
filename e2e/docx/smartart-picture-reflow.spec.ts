import { readFileSync } from 'node:fs';
import { expect, test } from '@playwright/test';
import { withNativePictureLayout } from '../../src/core/diagram/test-support/picture-layout';
import { fileInput, reveal } from './helpers';
import { savedDocument } from './saved-document';

for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid']) {
	test(`${framework}: native picture SmartArt reflows and saves its images`, async ({ page }) => {
		await page.setViewportSize({ width: 1700, height: 950 });
		await page.goto(`/?framework=${framework}`);
		const bytes = await withNativePictureLayout(
			readFileSync(new URL('../../src/core/docx/__fixtures__/smartart.docx', import.meta.url)),
			'word',
		);
		const upload = async (buffer: Uint8Array) =>
			(await fileInput(page)).setInputFiles({
				name: 'pictures.docx',
				mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
				buffer: Buffer.from(buffer),
			});
		await upload(bytes);
		const editor = page.locator('docx-editor');
		const diagram = editor.locator('.ProseMirror .dve-smartart').first();
		await expect(diagram.locator('image')).toHaveCount(3);
		await diagram.click();
		await editor.getByRole('tab', { name: 'SmartArt Design', exact: true }).click();
		const button = editor.getByRole('button', { name: 'Text Pane', exact: true });
		await reveal(editor, button);
		await button.click();
		await editor
			.locator('office-ui-smartart-text-pane input[data-node]')
			.first()
			.fill('Picture reflow caption');
		await expect(diagram.locator('image')).toHaveCount(3);
		const saved = await savedDocument(page, 'picture-reflow.docx');
		await upload(saved.bytes);
		await expect(diagram.locator('image')).toHaveCount(3);
		await expect(diagram).toContainText('Picture reflow caption');
	});
}
