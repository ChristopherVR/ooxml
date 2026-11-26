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
		expect((await dialog.getByRole('dialog').boundingBox())?.width).toBe(760);
		await expect(dialog.getByRole('tab', { name: 'My Add-ins', exact: true })).toHaveAttribute(
			'aria-selected',
			'true',
		);
		await expect(dialog.getByRole('tab', { name: 'Store', exact: true })).toHaveCount(0);
		const search = dialog.getByRole('searchbox');
		await search.fill('notes');
		await expect(dialog).toContainText('No matching add-ins.');
		const upload = dialog.getByRole('button', { name: 'Upload My Add-in' });
		await expect(upload).toBeDisabled();
		await expect(upload).toHaveAccessibleDescription(
			'Local add-in installation is unavailable in Visio.',
		);
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

test('shows configured Visio storefronts and explains unavailable installation', async ({
	page,
}) => {
	await page.setViewportSize({ width: 2400, height: 900 });
	await openDemo(page, '/demo-vanilla/?sample=1');
	const viewer = page.locator('visio-viewer');
	await viewer.getByRole('tab', { name: 'Insert', exact: true }).click();
	const get = viewer.getByRole('button', { name: 'Get Add-ins', exact: true });
	const dialog = viewer.locator('office-ui-add-ins-dialog office-ui-dialog[open]');
	await viewer.evaluate((element) =>
		Reflect.set(element, Symbol.for('office-ui.add-in-storefront'), { kind: 'microsoft' }),
	);
	await get.click();
	await expect(dialog.getByRole('tab', { name: 'Store', exact: true })).toHaveAttribute(
		'aria-selected',
		'true',
	);
	await expect(dialog.getByRole('link', { name: 'Browse Microsoft Marketplace' })).toBeVisible();
	await page.keyboard.press('Escape');
	await expect(dialog).toBeHidden();
	await viewer.evaluate((element) => {
		Reflect.set(element, Symbol.for('office-ui.add-in-storefront'), {
			kind: 'catalog',
			entries: [{ id: 'notes', label: 'Notes', read: () => new File([], 'notes.xml') }],
		});
		Reflect.set(element, Symbol.for('office-ui.add-in-marketplace-resolver'), async () => {
			throw new Error('Visio must not call the resolver');
		});
	});
	await get.click();
	const install = dialog.getByRole('button', { name: 'Install Notes' });
	await expect(install).toBeDisabled();
	await expect(install).toHaveAccessibleDescription(
		'Local add-in installation is unavailable in Visio.',
	);
	await dialog.getByRole('tab', { name: 'My Add-ins', exact: true }).click();
	await expect(dialog.getByLabel('Add from Marketplace link')).toHaveCount(0);
});

test('draws the add-ins glyph on the folded Visio group', async ({ page }) => {
	await page.setViewportSize({ width: 375, height: 812 });
	await openDemo(page, '/demo-vanilla/?sample=1');
	const viewer = page.locator('visio-viewer');
	await viewer.getByRole('tab', { name: 'Insert', exact: true }).click();
	const group = viewer.locator('office-ui-ribbon-group[label="Add-ins"][data-collapsed]');
	await expect(group).toHaveAttribute('icon', 'addIns');
	const face = group.getByRole('button', { name: 'Add-ins', exact: true });
	await expect(face).toBeVisible();
	await expect(face.locator('.face-icon path')).toHaveAttribute('d', /\S/u);
});
