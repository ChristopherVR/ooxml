/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright API */
import { readFile } from 'node:fs/promises';

import { expect, test } from '@playwright/test';
import JSZip from 'jszip';

import { savePptxViaBackstage } from './save-pptx';
import { loadDeck, ribbonTab } from './support/deck';
import { pressShortcut } from './support/keyboard';

test.use({ viewport: { width: 1920, height: 900 } });

test('keeps focus on Add-ins when the Insert ribbon refits', async ({ page }) => {
	await page.setViewportSize({ width: 2400, height: 900 });
	await loadDeck(page);
	await ribbonTab(page, 'Insert').click();
	const group = page.locator('pptx-ui-ribbon-insert [data-ribbon-group="insert.addIns"]');
	const get = group.getByRole('button', { name: 'Get Add-ins' });
	await get.click();
	const dialog = page.locator('office-ui-add-ins-dialog office-ui-dialog[open]');
	await expect(dialog).toBeVisible();
	expect((await dialog.getByRole('dialog').boundingBox())?.width).toBe(760);
	await page.keyboard.press('Escape');
	await expect(dialog).toBeHidden();
	await expect(get).toBeFocused();
	await page.setViewportSize({ width: 800, height: 900 });
	await expect(group).toHaveAttribute('data-collapsed', '');
	await expect(group.getByRole('button', { name: 'Add-ins', exact: true })).toBeFocused();
});

test('reviews, cancels, installs and searches local PowerPoint add-ins', async ({ page }) => {
	await page.setViewportSize({ width: 2400, height: 900 });
	await loadDeck(page);
	const insert = ribbonTab(page, 'Insert');
	await insert.click();
	const ribbon = page.locator('pptx-ui-ribbon-insert');
	const group = ribbon.locator('[data-ribbon-group="insert.addIns"]');
	const get = group.getByRole('button', { name: 'Get Add-ins' });
	const mine = group.getByRole('button', { name: 'My Add-ins' });
	const dialog = page.locator('office-ui-add-ins-dialog office-ui-dialog[open]');
	await get.press('Enter');
	await expect(dialog).toBeVisible();
	await expect(ribbon.locator('office-ui-add-ins-dialog')).toHaveCount(0);
	await page.keyboard.press('Escape');
	await expect(dialog).toBeHidden();
	await expect(get).toBeFocused();

	await mine.press('Enter');
	const search = dialog.getByRole('searchbox');
	await expect(search).toBeFocused();
	const upload = dialog.getByRole('button', { name: 'Upload My Add-in' });
	const buffer = await readFile(
		new URL('../../src/core/automation/add-ins/__fixtures__/manifest.xml', import.meta.url),
	);
	const payload = { name: 'notes.xml', mimeType: 'application/xml', buffer };
	const chooser = page.waitForEvent('filechooser');
	await upload.click();
	await (await chooser).setFiles(payload);
	const consent = dialog.getByRole('region', { name: 'Review add-in permissions' });
	await expect(consent).toContainText('Read and change this document');
	await consent.getByRole('button', { name: 'Cancel' }).click();
	await expect(search).toBeFocused();
	await expect(dialog).toContainText('No add-ins');
	await dialog.getByLabel('Choose a local add-in file').setInputFiles(payload);
	await consent.getByRole('button', { name: 'Install Notes' }).click();
	await expect(dialog).toBeHidden();
	await expect(mine).toBeFocused();

	await ribbonTab(page, 'Home').click();
	await insert.click();
	await mine.press('Enter');
	await expect(dialog).toContainText('Notes (1.0.0)');
	await expect(dialog).toContainText('Running add-ins is unavailable');
	await search.fill('unmatched');
	await expect(dialog).toContainText('No matching add-ins.');
	await search.fill('notes');
	await expect(dialog).toContainText('Notes (1.0.0)');
	await dialog.getByRole('button', { name: 'Close', exact: true }).click();
	await expect(dialog).toBeHidden();
	await expect(mine).toBeFocused();
	await mine.press('Enter');
	await dialog.getByRole('button', { name: 'Remove Notes', exact: true }).click();
	await expect(dialog).toBeHidden();
	await expect(mine).toBeFocused();
	await mine.press('Enter');
	await expect(dialog).toContainText('No add-ins');
});

test('inserts a Cameo and response link and opens the add-in dialogs', async ({
	page,
}, testInfo) => {
	await loadDeck(page);
	await ribbonTab(page, 'Insert').click();
	const ribbon = page.locator('pptx-ui-ribbon-insert');
	const command = (id: string) => ribbon.locator(`[data-ribbon-control="${id}"] button`);
	const reveal = async (id: string) => {
		const collapsed = ribbon.locator(
			`[data-ribbon-group="${id}"][data-collapsed]:not([data-open])`,
		);
		if (await collapsed.count()) await collapsed.locator('.face').click();
	};
	await reveal('insert.camera');
	await command('insert.camera.cameo').click();
	await expect(page.locator('pptx-ui-cameo-preview')).toHaveCount(1);
	await expect(page.locator('pptx-ui-cameo-preview button')).toHaveAttribute(
		'aria-pressed',
		'false',
	);
	await expect(command('insert.camera.cameo')).toBeDisabled();

	await reveal('insert.forms');
	await command('insert.forms.forms').click();
	await expect(page.getByRole('dialog', { name: 'Forms', exact: true })).toBeVisible();
	const forms = ribbon.locator('office-ui-dialog[open]');
	await forms
		.getByRole('textbox', { name: 'Microsoft Forms response link' })
		.fill('https://forms.office.com/r/Demo123');
	await forms.getByRole('button', { name: 'Insert link' }).click();
	await expect(
		page.locator('[data-pptx-viewport]').getByText('Open form', { exact: true }).first(),
	).toBeVisible();

	await reveal('insert.addIns');
	await command('insert.addIns.getAddIns').click();
	const addIns = page.locator('office-ui-add-ins-dialog office-ui-dialog[open]');
	await expect(addIns.getByRole('tab', { name: 'My Add-ins', exact: true })).toHaveAttribute(
		'aria-selected',
		'true',
	);
	await expect(addIns.getByRole('tab', { name: 'Store', exact: true })).toHaveCount(0);
	await page.keyboard.press('Escape');
	await reveal('insert.addIns');
	await command('insert.addIns.myAddIns').click();
	await expect(addIns.getByRole('tab', { name: 'My Add-ins' })).toHaveAttribute(
		'aria-selected',
		'true',
	);
	await expect(addIns).toContainText('No add-ins are available in this editor.');
	await page.keyboard.press('Escape');

	const output = testInfo.outputPath('cameo-form.pptx');
	await (await savePptxViaBackstage(page)).saveAs(output);
	const saved = await JSZip.loadAsync(await readFile(output));
	expect(await saved.file('ppt/slides/slide1.xml')?.async('string')).toContain('alf:liveFeedProps');
	expect(await saved.file('ppt/slides/_rels/slide1.xml.rels')?.async('string')).toContain(
		'https://forms.office.com/r/Demo123',
	);
	await loadDeck(page, output);
	await expect(page.locator('pptx-ui-cameo-preview')).toHaveCount(1);
	await page.route('https://forms.office.com/r/Demo123?embed=true', (route) =>
		route.fulfill({ contentType: 'text/html', body: '<p>Form response surface</p>' }),
	);
	await pressShortcut(page, 'F5');
	const show = page.locator('[data-pptx-presenting][aria-roledescription="slide"]');
	await show.getByText('Open form', { exact: true }).click();
	const response = page.locator('office-ui-dialog[open]').filter({
		has: page.getByRole('heading', { name: 'Forms', exact: true }),
	});
	await expect(response).toBeVisible();
	await expect(response.locator('iframe')).toHaveCount(0);
	await response.getByRole('button', { name: 'Load Microsoft Form' }).click();
	await expect(response.frameLocator('iframe').getByText('Form response surface')).toBeVisible();
	await response.getByRole('button', { name: 'Close', exact: true }).click();
	await expect(response).toHaveCount(0);
	await expect(show).toBeVisible();
});
