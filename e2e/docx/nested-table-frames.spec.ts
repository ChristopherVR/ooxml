import { expect, test } from '@playwright/test';
import { readFileSync } from 'node:fs';
import { fileInput } from './helpers';
for (const name of [
	'short',
	'two-rows',
	'centered',
	'split',
	'align-top',
	'align-center',
	'align-bottom',
])
	test(`nested table frames ${name} survive save and reopen`, async ({ page }) => {
		await page.goto('/?framework=vanilla');
		const editor = page.locator('docx-editor');
		const input = editor.locator('.ProseMirror').first();
		await (
			await fileInput(page)
		).setInputFiles({
			name: name + '.docx',
			mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
			buffer: readFileSync(
				new URL(
					'../../src/core/docx/__fixtures__/nested-table-frames/nested-frame-' + name + '.docx',
					import.meta.url,
				),
			),
		});
		await expect(input).toContainText('Outer cell');
		const print = async () => {
			await editor.getByRole('tab', { name: 'View', exact: true }).click();
			await editor
				.getByRole('combobox', { name: 'Layout view', exact: true })
				.selectOption('print');
		};
		await print();
		const frames = editor.locator('.dve-print-nested-table');
		await expect(frames).toHaveCount(name === 'split' ? 2 : 1);
		const check = async () => {
			const first = frames.first().locator('.dve-print-cell').first();
			await expect(first).toHaveCSS('background-color', 'rgb(255, 255, 0)');
			await expect(first).toHaveCSS('border-left-color', 'rgb(255, 0, 0)');
			expect(
				await first.evaluate((node) => parseFloat((node as HTMLElement).style.borderLeftWidth)),
			).toBeCloseTo(4 / 3, 1);
			await expect(frames.first()).toHaveText('');
			if (name.startsWith('align-')) {
				const paragraph = editor
					.locator('.dve-print-block[data-read-only="true"]')
					.filter({ hasText: 'T1R0C0L00' });
				const paragraphY = await paragraph.evaluate((node) =>
					parseFloat((node as HTMLElement).style.top),
				);
				const frameY = await frames
					.first()
					.evaluate((node) => parseFloat((node as HTMLElement).style.top));
				expect(paragraphY - frameY).toBeCloseTo(
					name === 'align-center' ? 12 : name === 'align-bottom' ? 24 : 0,
					1,
				);
			}

			await expect(
				editor.locator('.dve-print-block[data-read-only="true"]').filter({ hasText: 'T1R0C0L00' }),
			).toHaveCount(1);
		};
		await check();
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
		await expect(input).toContainText('Outer cell');
		await print();
		await check();
	});
