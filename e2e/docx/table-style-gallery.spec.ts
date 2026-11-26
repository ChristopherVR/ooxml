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

test('pressed table style options meet text and icon contrast in dark theme', async ({ page }) => {
	await page.emulateMedia({ colorScheme: 'dark' });
	await page.goto('/?framework=vanilla');
	await newDocument(page);
	await insertTableOfSize(page);
	const editor = page.locator('docx-editor');
	await editor.evaluate((element) => element.setAttribute('theme', 'dark'));
	await editor.locator('.ProseMirror td').first().click();
	await editor.getByRole('tab', { name: 'Table Design', exact: true }).click();
	const header = editor.getByRole('button', { name: 'Header Row', exact: true });
	await expect(header).toHaveAttribute('aria-pressed', 'true');
	const contrast = await header.evaluate((button) => {
		const canvas = document.createElement('canvas');
		canvas.width = canvas.height = 1;
		const ctx = canvas.getContext('2d')!;
		const paint = (color: string) => {
			ctx.fillStyle = color;
			ctx.fillRect(0, 0, 1, 1);
		};
		const luminance = () => {
			const [r, g, b] = [...ctx.getImageData(0, 0, 1, 1).data].map((channel) => {
				const value = channel / 255;
				return value <= 0.04045 ? value / 12.92 : ((value + 0.055) / 1.055) ** 2.4;
			});
			return 0.2126 * r! + 0.7152 * g! + 0.0722 * b!;
		};
		const layers = [];
		for (let element: Element | null = button; element;) {
			layers.unshift(getComputedStyle(element).backgroundColor);
			const scope = element.getRootNode();
			element = element.parentElement ?? (scope instanceof ShadowRoot ? scope.host : null);
		}
		paint('white');
		for (const color of layers) paint(color);
		const background = luminance();
		const icon = button.querySelector('.ribbon-icon');
		return [button, ...(icon ? [icon] : [])].map((element) => {
			paint(getComputedStyle(element).color);
			const foreground = luminance();
			return (Math.max(foreground, background) + 0.05) / (Math.min(foreground, background) + 0.05);
		});
	});
	for (const ratio of contrast) expect(ratio).toBeGreaterThanOrEqual(4.5);
});
