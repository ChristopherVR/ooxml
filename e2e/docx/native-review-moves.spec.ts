import { fileURLToPath } from 'node:url';
import { expect, test } from '@playwright/test';
import JSZip from 'jszip';
import { fileInput } from './helpers';
import type { DocxEditorElement } from '../../viewers/docx/packages/web-component/src';

const fixture = fileURLToPath(
	new URL('../../src/core/docx/__fixtures__/review-moves/core-created.docx', import.meta.url),
);
for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid']) {
	test(`${framework}: native move acceptance, undo and export preserve the second move`, async ({
		page,
	}) => {
		const errors: string[] = [];
		page.on('pageerror', (error) => errors.push(error.message));
		await page.goto(`/?framework=${framework}`);
		await (await fileInput(page)).setInputFiles(fixture);
		const editor = page.locator('docx-editor');
		const moves = editor.locator('.ProseMirror .dve-revision-move');
		await expect(moves).toHaveCount(4);
		await editor.getByRole('tab', { name: 'Review', exact: true }).click();
		await editor.locator('.ProseMirror del[data-move="move1"]').dblclick();
		await editor.getByRole('button', { name: 'Accept', exact: true }).click();
		await expect(moves).toHaveCount(2);
		await expect(editor.locator('.ProseMirror del')).toContainText('SecondMove');
		const bytes = await editor.evaluate(async (element) =>
			Array.from(await (element as DocxEditorElement).saveBytes()),
		);
		const xml = await (
			await JSZip.loadAsync(new Uint8Array(bytes))
		)
			.file('word/document.xml')!
			.async('string');
		expect(xml).not.toContain('w:delText');
		expect(xml).not.toContain('w:name="move1"');
		expect(xml.match(/w:name="move2"/g)).toHaveLength(2);
		await editor.locator('.ProseMirror').click();
		await page.keyboard.press('Control+z');
		await expect(moves).toHaveCount(4);
		await expect(editor.locator('.ProseMirror del[data-move="move1"]')).toContainText('FirstMove');
		await page.keyboard.press('Control+Shift+z');
		await expect(moves).toHaveCount(2);
		expect(errors).toEqual([]);
	});
}
