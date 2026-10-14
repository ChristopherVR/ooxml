import { expect, test } from '@playwright/test';
import { newDocument } from './helpers';

for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid']) {
	test(`${framework}: native Word tabs, split commands and adaptive Styles`, async ({ page }) => {
		await page.setViewportSize({ width: 2200, height: 900 });
		await page.goto(`/?framework=${framework}`);
		await newDocument(page);
		const editor = page.locator('docx-editor');
		const ribbon = editor.locator('office-ui-ribbon');
		await expect(ribbon.getByRole('tab')).toHaveText([
			'Home',
			'Insert',
			'Design',
			'Layout',
			'References',
			'Mailings',
			'Review',
			'View',
			'Help',
		]);
		const gallery = editor.locator('.ribbon-gallery');
		const wide = (await gallery.boundingBox())!.width;
		await page.setViewportSize({ width: 1600, height: 900 });
		await expect.poll(async () => (await gallery.boundingBox())!.width).toBeLessThan(wide - 400);
		const body = editor.locator('.ProseMirror');
		await body.click();
		await page.keyboard.type('Toolbar parity');
		await page.keyboard.press('Control+a');
		await ribbon
			.getByRole('combobox', { name: 'Underline style', exact: true })
			.selectOption('double');
		await expect(body.locator('[data-run-props] > span')).toHaveCSS(
			'text-decoration-style',
			'double',
		);
		await expect(body.locator('u')).toHaveCSS('text-decoration-line', 'none');
		await page.keyboard.press('Control+z');
		await expect(body.locator('u')).toHaveCount(0);
		await ribbon.getByRole('tab', { name: 'Mailings', exact: true }).click();
		await expect(
			ribbon.getByRole('button', { name: 'Start Mail Merge', exact: true }),
		).toBeDisabled();
		await ribbon.getByRole('tab', { name: 'Design', exact: true }).click();
		await expect(ribbon.getByRole('button', { name: 'Watermark', exact: true })).toBeVisible();
	});
}

test('styled underline ranges sharing a semantic mark retain their own line style', async ({
	page,
}) => {
	await page.setViewportSize({ width: 1600, height: 900 });
	await page.goto('/?framework=vanilla');
	await newDocument(page);
	const editor = page.locator('docx-editor');
	const body = editor.locator('.ProseMirror');
	await body.click();
	await page.keyboard.type('left right');
	await page.keyboard.press('Control+a');
	const styles = editor.getByRole('combobox', { name: 'Underline style', exact: true });
	await styles.selectOption('double');
	await expect(body.locator('[data-run-props]')).toHaveText('left right');
	await expect(body.locator('[data-run-props] > span')).toHaveCSS(
		'text-decoration-style',
		'double',
	);
	await body.focus();
	await page.keyboard.press('Control+Home');
	// Native selectionchange must reach the editor before the next ribbon command.
	for (let index = 0; index < 4; index++)
		await page.keyboard.press('Shift+ArrowRight', { delay: 20 });
	await expect.poll(() => page.evaluate(() => window.getSelection()?.toString())).toBe('left');
	await styles.selectOption('wave');
	await expect(body.locator('[data-run-props] > span').filter({ hasText: 'left' })).toHaveCSS(
		'text-decoration-style',
		'wavy',
	);
	await expect(body.locator('[data-run-props] > span').filter({ hasText: 'right' })).toHaveCSS(
		'text-decoration-style',
		'double',
	);
});
