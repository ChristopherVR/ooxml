import { expect, test } from '@playwright/test';
import { clickInEditor, insertTableOfSize, newDocument } from './helpers';
import { savedDocument } from './saved-document';

test('changing one column width preserves its neighbors and reports each cell size', async ({
	page,
}) => {
	await page.setViewportSize({ width: 1700, height: 960 });
	await page.goto('/?framework=vanilla');
	await newDocument(page);
	await insertTableOfSize(page, 2, 4);
	const editor = page.locator('docx-editor');
	const rows = editor.locator('.ProseMirror tr');
	const firstRow = rows.first().locator('td');
	const widths = () =>
		firstRow.evaluateAll((cells) => cells.map((cell) => (cell as HTMLElement).offsetWidth));
	const original = await widths();
	await clickInEditor(firstRow.nth(1).locator('p'));
	await editor.getByRole('tab', { name: 'Table Layout', exact: true }).click();
	const width = editor.getByRole('spinbutton', { name: 'Column Width (inches)' });
	const height = editor.getByRole('spinbutton', { name: 'Row Height (inches)' });

	const selected = original[1]!;
	await width.fill('2');
	await width.press('Enter');
	await expect.poll(async () => (await widths())[1]).toBeGreaterThan(selected + 20);
	const resized = await widths();
	for (const column of [0, 2, 3])
		expect(Math.abs(resized[column]! - original[column]!)).toBeLessThanOrEqual(1);
	await height.fill('0.75');
	await height.press('Enter');

	const secondRowCell = rows.nth(1).locator('td').first();
	await clickInEditor(secondRowCell.locator('p'));
	const secondRowSize = await secondRowCell.evaluate((cell) => ({
		width: (cell as HTMLElement).offsetWidth / 96,
		height: (cell.parentElement?.offsetHeight ?? 0) / 96,
	}));
	await expect(width).toHaveValue(String(Number(secondRowSize.width.toFixed(2))));
	await expect(height).toHaveValue(String(Number(secondRowSize.height.toFixed(2))));
	await clickInEditor(firstRow.nth(1).locator('p'));
	await expect(width).toHaveValue('2');
	await expect(height).toHaveValue('0.75');

	const { xml } = await savedDocument(page, 'sized-columns.docx');
	const grid = [...xml.matchAll(/<w:gridCol w:w="(\d+)"/g)].map((match) => Number(match[1]));
	expect(grid).toHaveLength(4);
	expect(grid[1]).toBe(2880);
	for (const column of [0, 2, 3])
		expect(Math.abs(grid[column]! / 15 - original[column]!)).toBeLessThanOrEqual(1);
});
