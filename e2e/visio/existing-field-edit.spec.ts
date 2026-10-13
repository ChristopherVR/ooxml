import { test, expect } from '@playwright/test';
import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import type { VisioViewerElement } from 'ooxml-ui/visio';
import { openDemo } from './demo-page';
const native = readFileSync(
	new URL('../../src/core/visio/__fixtures__/field-edit-native.vsdx', import.meta.url),
);
for (const [id, formula, format, expected, timezone] of [
	['1', 'Width*2', '0.00', 'Before 8.00 | 1 After', 'Australia/Brisbane'],
	['2', 'DOCCREATION()', '{{yyyy-MM-dd}}', 'Before 2026-10-14 | 1 After', 'Australia/Brisbane'],
	['3', 'TITLE()', '@', 'Before  | 1 After', 'Australia/Brisbane'],
	['2', 'DOCCREATION()', '{{yyyy-MM-dd}}', 'Before 2026-10-13 | 1 After', 'UTC'],
]) {
	test.describe('existing field ' + id + ' in ' + timezone, () => {
		test.use({ timezoneId: timezone! });
		test(
			'Field dialog edits existing native field ' + id + ' and preserves history/export',
			async ({ page }) => {
				await openDemo(page);
				const viewer = page.locator('visio-viewer');
				const upload = async (buffer: Buffer, name: string) => {
					await page
						.locator('#file')
						.setInputFiles({ name, mimeType: 'application/vnd.ms-visio.drawing', buffer });
					await expect(page.locator('#file-name')).toHaveText(name);
				};
				await upload(native, 'Existing fields.vsdx');
				const shape = viewer.locator(`svg.paper [data-shape-id="${id}"]`).first(),
					text = shape.locator('text');
				const original = await text.textContent();
				await shape.click();
				await viewer.getByRole('tab', { name: 'Insert', exact: true }).click();
				await viewer.locator('[command="field"] button').click();
				const dialog = viewer.locator('.field-dialog');
				await dialog.locator('[name="field-target"]').selectOption('0');
				await expect(dialog.locator('[name="field-formula"]')).not.toHaveValue('');
				if (id === '2' || id === '3') {
					await dialog
						.locator('[name="field-category"]')
						.selectOption(id === '2' ? 'date' : 'document');
					await dialog.locator('[name="field-field"]').selectOption(formula!);
				} else await dialog.locator('[name="field-formula"]').fill(formula!);
				await dialog.locator('[name="field-format"]').selectOption(format!);
				await dialog.locator('[label="OK"] button').click();
				await expect(text).toHaveText(expected!);
				await viewer.evaluate((node) => (node as VisioViewerElement).undo());
				await expect(text).toHaveText(original!);
				await viewer.evaluate((node) => (node as VisioViewerElement).redo());
				await expect(text).toHaveText(expected!);
				const bytes = await viewer.evaluate((node) =>
					Array.from((node as VisioViewerElement).exportVsdx().bytes),
				);
				const originalZip = await JSZip.loadAsync(native),
					savedZip = await JSZip.loadAsync(bytes);
				for (const part of Object.keys(originalZip.files).filter(
					(part) => !originalZip.files[part]!.dir && part !== 'visio/pages/page1.xml',
				))
					expect(await savedZip.file(part)!.async('uint8array')).toEqual(
						await originalZip.file(part)!.async('uint8array'),
					);
				await upload(Buffer.from(bytes), 'Reopened fields.vsdx');
				await expect(text).toHaveText(expected!);
				const fields = await viewer.evaluate(
					(node, id) =>
						(node as VisioViewerElement).controller.state.document!.pages[0]!.shapes.find(
							(shape) => shape.id === id,
						)!.text.fields,
					id!,
				);
				expect(fields).toHaveLength(2);
				expect(fields![0]!.index).toBe('0');
				expect(fields![0]!.formula).toBe(formula);
			},
		);
	});
}
