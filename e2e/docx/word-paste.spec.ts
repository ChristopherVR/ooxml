import { expect, test } from '@playwright/test';
import { readFile } from 'node:fs/promises';
import type { DocxEditorElement } from '../../viewers/docx/packages/web-component/src';
import { newDocument } from './helpers';

/** HTML Format data Word 16 put on the clipboard (src/core/docx/__fixtures__/clipboard). */
const wordHtml = (name: string) =>
	readFile(
		new URL(`../../src/core/docx/__fixtures__/clipboard/${name}.html`, import.meta.url),
		'utf8',
	);

for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid']) {
	test(`${framework}: a Word list pastes as a list and Paste Options paste it again as text`, async ({
		page,
	}) => {
		await page.goto(`/?framework=${framework}`);
		await newDocument(page);
		const editor = page.locator('docx-editor');
		const body = editor.locator('.dve-paper > .ProseMirror');
		await body.click();
		const html = await wordHtml('word-numbered');
		await body.evaluate((element, html) => {
			const data = new DataTransfer();
			data.setData('text/html', html);
			data.setData('text/plain', 'Step one\nSub step a');
			element.dispatchEvent(
				new ClipboardEvent('paste', { clipboardData: data, bubbles: true, cancelable: true }),
			);
		}, html);
		await expect(body).toContainText('Sub sub step');
		const numbered = await editor.evaluate((element) =>
			(element as DocxEditorElement).documentModel!.blocks.flatMap((block) =>
				block.type === 'paragraph' && block.numbering
					? [[block.runs.map((run) => run.text).join(''), block.numbering.level]]
					: [],
			),
		);
		expect(numbered).toEqual([
			['Step one', 0],
			['Sub step a', 1],
			['Sub sub step', 2],
			['Sub step b', 1],
			['Step two', 0],
		]);
		// No marker text from Word's HTML reaches the document.
		await expect(body).not.toContainText('1)');

		const button = editor.getByRole('button', { name: 'Paste Options (Ctrl)' });
		await expect(button).toBeVisible();
		await button.click();
		await editor.getByRole('menuitemcheckbox', { name: 'Keep Text Only' }).click();
		await expect(body).toHaveText('Step oneSub step a');
		await page.keyboard.press('Control+z');
		await expect(body).toContainText('Sub sub step');

		// Typing ends the offer, as in Word.
		await page.keyboard.type('x');
		await expect(button).toBeHidden();
	});
}
