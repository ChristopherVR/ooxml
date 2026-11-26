/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright API */
import { readFile } from 'node:fs/promises';
import { expect, test } from '@playwright/test';
import JSZip from 'jszip';
import { editor as excel, newWorkbook } from './helpers';

const fixture = (name: string) =>
	readFile(
		new URL(`../../src/core/automation/add-ins/__fixtures__/${name}`, import.meta.url),
		'utf8',
	);

test('keeps the XML manifest compatibility details in Excel My Add-ins', async ({ page }) => {
	await page.setViewportSize({ width: 1600, height: 900 });
	await newWorkbook(page);
	const editor = excel(page);
	await editor.getByRole('tab', { name: 'Insert', exact: true }).click();
	await editor.getByRole('button', { name: 'Get Add-ins', exact: true }).click();
	const dialog = editor.locator('office-ui-dialog[open]');
	await dialog.getByLabel('Choose a local add-in file').setInputFiles({
		name: 'notes.xml',
		mimeType: 'application/xml',
		buffer: Buffer.from(await fixture('manifest.xml')),
	});
	await dialog.getByRole('button', { name: 'Install Notes', exact: true }).click();
	await editor.getByRole('button', { name: 'My Add-ins', exact: true }).click();
	await expect(dialog).toContainText('Unavailable requirement · AddinCommands 1.10');
	await expect(dialog).toContainText('Unsupported feature · ExecuteFunction');
});

test('reviews and installs a unified package in Excel', async ({ page }) => {
	await page.setViewportSize({ width: 1600, height: 900 });
	await newWorkbook(page);
	const editor = excel(page);
	await editor.getByRole('tab', { name: 'Insert', exact: true }).click();
	const get = editor.getByRole('button', { name: 'Get Add-ins', exact: true });
	const mine = editor.getByRole('button', { name: 'My Add-ins', exact: true });
	await get.click();
	const dialog = editor.locator('office-ui-dialog[open]');
	const manifest = (await fixture('manifest.json')).replace(
		'"lifetime": "short"',
		'"lifetime": "long"',
	);
	const zip = new JSZip()
		.file('icons/outline.png', new Uint8Array([1]))
		.file('icons/color.png', new Uint8Array([2]))
		.file('fr.json', await fixture('fr.json'));
	const input = dialog.getByLabel('Choose a local add-in file');
	await expect(input).toBeFocused();
	const marketplace = dialog.getByRole('link', { name: /Microsoft Marketplace/ });
	await expect(marketplace).toBeVisible();
	const select = async (source: string) =>
		input.setInputFiles({
			name: 'notes.zip',
			mimeType: 'application/zip',
			buffer: await zip.file('manifest.json', source).generateAsync({ type: 'nodebuffer' }),
		});
	await select(manifest);
	const consent = dialog.getByRole('region', { name: 'Review add-in permissions' });
	await expect(consent.getByRole('heading')).toHaveText('Review add-in permissions');
	await expect(consent).toContainText('This add-in requests:');
	await expect(consent).toContainText('Read and change this document');
	await expect(consent).toContainText('Lab Notes');
	await expect(consent).toContainText('https://notes.example.test');
	await expect(consent).toContainText('https://assets.example.test');
	await expect(consent).not.toContainText('Document.ReadWrite.User');
	await expect(marketplace).toBeHidden();
	await expect(dialog).toContainText('Unavailable requirement · AddinCommands 1.10');
	await expect(dialog).toContainText('Unsupported feature · shared runtime');
	await consent.getByRole('button', { name: 'Install Notes', exact: true }).click();
	await mine.click();
	const installed = editor.locator('office-ui-dialog[open]');
	await expect(installed).toContainText('Running add-ins is unavailable');
	await expect(installed).toContainText('Unavailable requirement · AddinCommands 1.10');
	await expect(installed).toContainText('Unsupported feature · shared runtime');
	await page.keyboard.press('Escape');

	await get.click();
	await select(
		manifest.replace('1.0.0', '1.0.1').replace('"short": "Notes"', '"short": "Updated Notes"'),
	);
	await expect(dialog).toContainText('Updated Notes (1.0.1)');
	await expect(dialog.getByText('Update from 1.0.0 to 1.0.1', { exact: true })).toBeVisible();
	await dialog.getByRole('button', { name: 'Update Updated Notes', exact: true }).click();

	await mine.click();
	await expect(installed).toContainText('Updated Notes (1.0.1)');
	await installed.getByRole('button', { name: 'Remove Updated Notes', exact: true }).click();
});
