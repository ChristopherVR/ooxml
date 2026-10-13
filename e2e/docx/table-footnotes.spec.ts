import { expect, test } from '@playwright/test';
import { readFileSync } from 'node:fs';
import { fileInput } from './helpers';
for (const name of [
	'padded-split-early',
	'padded-split-late',
	'split-late-reference',
	'repeated-header',
	'mode14-split-long-note-columns',
])
	test(`table notes ${name} retain pagination through edit undo and reopen`, async ({ page }) => {
		await page.goto('/?framework=vanilla');
		await (
			await fileInput(page)
		).setInputFiles({
			name: name + '.docx',
			mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
			buffer: readFileSync(
				new URL(
					'../../src/core/docx/__fixtures__/table-footnotes/' + name + '.docx',
					import.meta.url,
				),
			),
		});
		const editor = page.locator('docx-editor'),
			input = editor.locator('.ProseMirror').first();
		await expect(input).toContainText('T0R0C0L00');
		const print = async () => {
			await editor.getByRole('tab', { name: 'View', exact: true }).click();
			await editor
				.getByRole('combobox', { name: 'Layout view', exact: true })
				.selectOption('print');
		};
		await print();
		const pages = editor.locator('.dve-print-page'),
			notes = editor.locator('.dve-print-footnote');
		await expect(pages).toHaveCount(
			name === 'mode14-split-long-note-columns' ? 4 : name === 'repeated-header' ? 3 : 2,
		);
		await expect(notes.first()).toContainText('N0L00');
		if (name === 'padded-split-early') {
			await expect(pages.nth(0).locator('.dve-print-cell').first()).toContainText('T0R0C0L12');
			await expect(pages.nth(0).locator('.dve-print-cell').first()).not.toContainText('T0R0C0L13');
			await expect(pages.nth(1).locator('.dve-print-cell').first()).toContainText('T0R0C0L13');
		}
		if (name === 'padded-split-late') {
			await expect(pages.nth(0).locator('.dve-print-footnote')).toHaveCount(0);
			await expect(pages.nth(0).locator('.dve-print-cell').first()).not.toContainText('T0R0C0L15');
			await expect(pages.nth(0).locator('.dve-print-cell').nth(1)).toContainText('T0R0C1L16');
			await expect(pages.nth(1).locator('.dve-print-cell').first()).toContainText('T0R0C0L15');
		}
		if (name === 'split-late-reference') {
			await expect(pages.nth(0).locator('.dve-print-footnote')).toContainText('N0L00');
			await expect(pages.nth(1).locator('.dve-print-footnote')).toContainText('N0L02');
		}
		if (name === 'repeated-header') {
			await expect(editor.locator('.dve-print-row[data-repeated="true"]')).toHaveCount(1);
			await expect(pages.nth(1).locator('.dve-print-footnote')).toHaveCount(0);
		}
		const first = editor.locator('.dve-print-cell').first();
		await first.click({ position: { x: 15, y: 8 } });
		await page.keyboard.type('Edit');
		await expect(input).toContainText('Edit');
		await page.keyboard.press('Control+z');
		await expect(input).not.toContainText('Edit');
		await page.keyboard.press('Control+y');
		await expect(input).toContainText('Edit');
		const bytes = await editor.evaluate(async (node) =>
			Array.from(await (node as unknown as { saveBytes(): Promise<Uint8Array> }).saveBytes()),
		);
		await (
			await fileInput(page)
		).setInputFiles({
			name: 'reopened.docx',
			mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
			buffer: Buffer.from(bytes),
		});
		await expect(input).toContainText('Edit');
		await print();
		await expect(notes.first()).toContainText('N0L00');
		await expect(editor.locator('.dve-print-footnote').last()).toContainText(
			name === 'mode14-split-long-note-columns' ? 'N0L79' : 'N0L02',
		);
	});
