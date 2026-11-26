import { expect, test } from '@playwright/test';
import { openDemo } from './demo-page';

for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid']) {
	test(`${framework}: explains unavailable local installation in the Office Add-ins dialog`, async ({
		page,
	}) => {
		await page.setViewportSize({ width: 2400, height: 900 });
		await openDemo(page, `/demo-${framework}/?sample=1`);
		const viewer = page.locator('visio-viewer');
		await viewer.getByRole('tab', { name: 'Insert', exact: true }).click();
		const get = viewer.getByRole('button', { name: 'Get Add-ins', exact: true });
		const mine = viewer.getByRole('button', { name: 'My Add-ins', exact: true });
		const dialog = viewer.locator('office-ui-add-ins-dialog office-ui-dialog[open]');
		await get.press('Enter');
		await expect(dialog.getByRole('tab', { name: 'Admin Managed' })).toHaveAttribute(
			'aria-selected',
			'true',
		);
		const search = dialog.getByRole('searchbox');
		await search.fill('notes');
		await expect(dialog).toContainText('No matching add-ins.');
		await expect(dialog.getByRole('button', { name: 'Upload My Add-in' })).toBeDisabled();
		await expect(dialog).toContainText('Local add-in installation is unavailable in Visio.');
		await page.keyboard.press('Escape');
		await expect(dialog).toBeHidden();
		await expect(get).toBeFocused();
		await mine.press('Enter');
		await expect(dialog.getByRole('tab', { name: 'My Add-ins' })).toHaveAttribute(
			'aria-selected',
			'true',
		);
		await expect(dialog.getByRole('button', { name: 'Upload My Add-in' })).toBeDisabled();
		await dialog.getByRole('button', { name: 'Close', exact: true }).click();
		await expect(dialog).toBeHidden();
		await expect(mine).toBeFocused();
	});
}
