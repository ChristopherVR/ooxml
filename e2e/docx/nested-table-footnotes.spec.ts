import { expect, test } from '@playwright/test';
import { readFileSync } from 'node:fs';
import { fileInput } from './helpers';
for (const name of ['nested-table-reference', 'nested-split-late', 'nested-second-cell-late'])
	test(`nested footnote ${name} renders read-only source through outer editing and reopen`, async ({
		page,
	}) => {
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
		await expect(input).toContainText('Outer cell');
		const print = async () => {
			await editor.getByRole('tab', { name: 'View', exact: true }).click();
			await editor
				.getByRole('combobox', { name: 'Layout view', exact: true })
				.selectOption('print');
		};
		await print();
		await expect(editor.locator('.dve-print-page')).toHaveCount(
			name === 'nested-table-reference' ? 1 : 2,
		);
		const nested = editor.locator('.dve-print-block[data-read-only="true"]').first();
		await expect(nested).toContainText('T1R0C0L00');
		await expect(editor.locator('.dve-print-footnote').first()).toContainText('N0L00');
		await expect(editor.locator('.dve-print-footnote').last()).toContainText('N0L02');
		await nested.click();
		await page.keyboard.type('Oops');
		await expect(input).not.toContainText('Oops');
		const outer = editor
			.locator('.dve-print-block:not(.dve-print-table)')
			.filter({ hasText: 'Outer cell' })
			.first();
		await outer.click({ position: { x: 15, y: 8 } });
		await page.keyboard.type('Edit');
		await expect(input).toContainText('Edit');
		await page.keyboard.press('Control+z');
		await expect(input).not.toContainText('Edit');
		await page.keyboard.press('Control+y');
		await expect(input).toContainText('Edit');
		const beforeStructureEdit = await input.textContent();
		await page.keyboard.press('End');
		await page.keyboard.press('Enter');
		await expect.poll(() => input.textContent()).toBe(beforeStructureEdit);
		await input.evaluate((node) => {
			const data = new DataTransfer();
			data.setData('text/plain', 'One\nTwo');
			node.dispatchEvent(
				new ClipboardEvent('paste', { clipboardData: data, bubbles: true, cancelable: true }),
			);
		});
		await expect.poll(() => input.textContent()).toBe(beforeStructureEdit);
		await page.keyboard.press('Shift+Enter');
		await page.keyboard.type('Line');
		await expect(input).toContainText('Line');
		await page.keyboard.press('Control+z');
		await expect(input).not.toContainText('Line');
		await page.keyboard.press('Control+y');
		await expect(input).toContainText('Line');
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
		await expect(editor.locator('.dve-print-block[data-read-only="true"]').first()).toContainText(
			'T1R0C0L00',
		);
		await expect(editor.locator('.dve-print-footnote').last()).toContainText('N0L02');
		if (name !== 'nested-table-reference') {
			const continued = editor
				.locator('.dve-print-page')
				.nth(1)
				.locator('.dve-print-block[data-read-only="true"]');
			await expect(continued.first()).toHaveCSS('left', '0px');
			await expect(continued.nth(1)).toHaveCSS('left', '106.667px');
			// Reflow can replace this block between resolving the locator and reading its style.
			// Retry the live geometry, as the CSS assertions above do, without widening the tolerance.
			await expect
				.poll(() => continued.nth(1).evaluate((node) => parseFloat(getComputedStyle(node).width)))
				.toBeCloseTo(106.6667, 1);
		}
	});
