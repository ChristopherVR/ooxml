import { expect, test } from '@playwright/test';
import { insertTableOfSize, newDocument } from './helpers';

test.use({ viewport: { width: 1280, height: 900 } });

test('Table Layout folds lower-priority groups at 1280 pixels', async ({ page }) => {
	await page.goto('/?framework=vanilla');
	await newDocument(page);
	await insertTableOfSize(page);
	const editor = page.locator('docx-editor');
	await editor.locator('.ProseMirror td').first().click();
	await editor.getByRole('tab', { name: 'Table Layout', exact: true }).click();
	const panel = editor.locator('[data-panel="Table Layout"]');
	const data = panel.getByRole('button', { name: 'Table Data', exact: true });
	await expect(data).toBeVisible();
	await expect(panel.getByRole('button', { name: 'Split cells', exact: true })).toBeVisible();
	await expect(panel.getByRole('combobox', { name: 'Select Table Parts' })).toBeVisible();
	const height = panel.getByRole('spinbutton', { name: 'Row Height (inches)' });
	await expect(height).toBeVisible();
	await expect(height).toHaveAttribute('title', 'Row Height (inches)');
	// The size boxes end before the Distribute commands instead of covering their icons.
	const distribute = (await panel.getByRole('button', { name: 'Distribute Rows' }).boundingBox())!;
	for (const name of ['Row Height (inches)', 'Column Width (inches)']) {
		const box = (await panel.getByRole('spinbutton', { name }).boundingBox())!;
		expect(box.x + box.width, name).toBeLessThanOrEqual(distribute.x);
	}
	await expect
		.poll(() => panel.evaluate((element) => element.scrollWidth - element.clientWidth))
		.toBe(0);
	await data.click();
	await expect(editor.getByRole('button', { name: 'Formula', exact: true })).toBeEnabled();
});
