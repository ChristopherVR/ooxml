/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright API */
import { readFile } from 'node:fs/promises';

import { expect, test } from '@playwright/test';
import JSZip from 'jszip';

import { savePptxViaBackstage } from './save-pptx';
import { loadDeck, ribbonTab } from './support/deck';
import { pressShortcut } from './support/keyboard';

test.use({ viewport: { width: 1920, height: 900 } });

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
	await expect(page.getByRole('dialog', { name: 'Get Add-ins' })).toBeVisible();
	const store = ribbon.locator('office-ui-dialog[open]');
	await expect(store.getByRole('link', { name: 'Browse Microsoft Marketplace' })).toHaveAttribute(
		'href',
		'https://marketplace.microsoft.com/en-us/marketplace/apps?product=office',
	);
	await page.keyboard.press('Escape');
	await reveal('insert.addIns');
	await command('insert.addIns.myAddIns').click();
	await expect(page.getByRole('dialog', { name: 'My Add-ins' })).toBeVisible();
	await expect(ribbon.locator('office-ui-dialog[open]')).toContainText(
		'No add-ins are available in this editor.',
	);
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
