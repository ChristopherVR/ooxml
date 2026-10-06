import { expect, test } from '@playwright/test';
import { openSample, dialogByHeading } from './helpers';

test('View > Zoom sets a percentage and fits the page width', async ({ page }) => {
	await page.setViewportSize({ width: 1700, height: 900 });
	await openSample(page);
	const editor = page.locator('docx-editor');
	await editor.locator('[role="tab"][data-tab="view"]').click();
	await editor.getByRole('button', { name: 'Zoom dialog' }).click();
	const dialog = dialogByHeading(editor, 'Zoom');
	await expect(dialog).toBeVisible();
	await dialog.getByLabel('Percent', { exact: true }).fill('75');
	await dialog.getByRole('button', { name: 'OK' }).click();
	await expect(dialog).toBeHidden();
	await expect(editor.locator('.dve-status')).toContainText('75%');
	await editor.getByRole('button', { name: 'Zoom dialog' }).click();
	await dialog.getByLabel('Page width', { exact: true }).check();
	await dialog.getByRole('button', { name: 'OK' }).click();
	await expect(editor.locator('.dve-status')).not.toContainText('75%');
	await editor.getByRole('button', { name: 'Zoom dialog' }).click();
	await page.keyboard.press('Escape');
	await expect(dialog).toBeHidden();
});
