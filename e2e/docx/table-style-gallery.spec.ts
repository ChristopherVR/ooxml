import { expect, test } from '@playwright/test';
import { clickInEditor, insertTableOfSize, newDocument, reveal } from './helpers';

test('the table style gallery wraps fixed-width previews under their family headings', async ({
	page,
}) => {
	await page.goto('/?framework=vanilla');
	await newDocument(page);
	await insertTableOfSize(page);
	const editor = page.locator('docx-editor');
	await clickInEditor(editor.locator('.ProseMirror td').first());
	await editor.getByRole('tab', { name: 'Table Design', exact: true }).click();
	await page.setViewportSize({ width: 355, height: 850 });
	const styles = editor.locator('[data-table-command="styles"]');
	await reveal(editor, styles);
	await styles.click();
	const gallery = editor.getByRole('listbox', { name: 'Table Styles' });
	for (const family of ['Plain Tables', 'Grid Tables', 'List Tables'])
		await expect(gallery.getByRole('group', { name: family })).toBeVisible();
	const options = gallery.getByRole('option');
	await expect(options).toHaveCount(11);
	const gridRows = await gallery
		.getByRole('group', { name: 'Grid Tables' })
		.getByRole('option')
		.evaluateAll((tiles) => tiles.map((tile) => tile.getBoundingClientRect().top));
	expect(new Set(gridRows).size).toBeGreaterThan(1);
	const rectangles = await options.evaluateAll((elements) =>
		elements.map((element) => {
			const tile = element.getBoundingClientRect();
			const preview = element.querySelector('.dve-table-style-preview')?.getBoundingClientRect();
			return {
				left: tile.left,
				right: tile.right,
				top: tile.top,
				bottom: tile.bottom,
				previewWidth: preview?.width ?? 0,
			};
		}),
	);
	for (const [index, tile] of rectangles.entries()) {
		expect(tile.previewWidth).toBeGreaterThanOrEqual(60);
		for (const other of rectangles.slice(index + 1)) {
			expect(
				tile.left < other.right &&
					tile.right > other.left &&
					tile.top < other.bottom &&
					tile.bottom > other.top,
			).toBe(false);
		}
	}
});
