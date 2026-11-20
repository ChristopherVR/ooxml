import { expect, test, type Locator, type Page } from '@playwright/test';
import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { fileInput } from './helpers';

// Written by `bun src/core/scripts/docx-word-cell-float-cases.ts <dir> demo` from
// src/core/docx/__fixtures__/cell-floats/cases.ts (cellFloatDemo) and opened in Word 16 without
// repair: four tables whose second cell holds a square-wrapped shape, a chart, a top-and-bottom
// shape and a shape behind its text, then a chart floating in the body.
const FIXTURE = fileURLToPath(new URL('./support/cell-floats-demo.docx', import.meta.url));

test.use({ viewport: { width: 1366, height: 1400 } });

async function open(page: Page): Promise<Locator> {
	await page.goto('/?framework=vanilla');
	await (
		await fileInput(page)
	).setInputFiles({
		name: 'cell-floats-demo.docx',
		mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
		buffer: await readFile(FIXTURE),
	});
	const editor = page.locator('docx-editor');
	await expect(editor.locator('.ProseMirror').first()).toContainText(
		'Floating objects in table cells',
	);
	await editor.getByRole('tab', { name: 'View', exact: true }).click();
	await editor.getByRole('button', { name: 'Print Layout', exact: true }).click();
	return editor;
}

test('Print Layout positions floating objects in their table cells and floating charts', async ({
	page,
}) => {
	const editor = await open(page);
	const sheet = editor.locator('.dve-print-page').first();
	const floats = sheet.locator(':scope > .dve-print-float');
	await expect(floats).toHaveCount(5);
	// Both charts are painted as SVG, as inline charts are.
	await expect(floats.nth(1).locator('svg').first()).toBeVisible();
	await expect(floats.nth(4).locator('svg').first()).toBeVisible();
	await expect(floats.nth(3)).toHaveClass(/dve-print-float-behind/);
	const tables = sheet.locator('.dve-print-table');
	await expect(tables).toHaveCount(4);
	for (let index = 0; index < 4; index++) {
		const row = tables.nth(index).locator('.dve-print-row').first();
		const cell = (await row.locator('.dve-print-cell').nth(1).boundingBox())!;
		const rowBox = (await row.boundingBox())!;
		const float = (await floats.nth(index).boundingBox())!;
		expect(float.x, `object ${index} left`).toBeGreaterThanOrEqual(cell.x - 1);
		expect(float.x + float.width, `object ${index} right`).toBeLessThanOrEqual(
			cell.x + cell.width + 1,
		);
		expect(float.y, `object ${index} top`).toBeGreaterThanOrEqual(rowBox.y - 1);
		// A wrapping object grows its row to hold it.
		if (index < 3)
			expect(float.y + float.height, `object ${index} bottom`).toBeLessThanOrEqual(
				rowBox.y + rowBox.height + 1,
			);
	}
	// The body chart is aligned right in the margins, and the body text runs left of it.
	const column = (await sheet.locator('.dve-print-column').first().boundingBox())!;
	const chart = (await floats.nth(4).boundingBox())!;
	expect(Math.abs(chart.x + chart.width - (column.x + column.width))).toBeLessThanOrEqual(1.5);
});
