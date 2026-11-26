/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright API */
import { readFile } from 'node:fs/promises';
import { expect, type Locator, test } from '@playwright/test';
import JSZip from 'jszip';
import { FRAMEWORKS, editor as excel, newWorkbook } from './helpers';

/** Clicks an Add-ins command, opening the group first when the ribbon has collapsed it. */
async function addInsCommand(editor: Locator, name: 'Get Add-ins' | 'My Add-ins') {
	const command = editor.getByRole('button', { name, exact: true });
	if (!(await command.isVisible()))
		await editor.getByRole('button', { name: 'Add-ins', exact: true }).click();
	await command.click();
}

const fixture = (name: string) =>
	readFile(
		new URL(`../../src/core/automation/add-ins/__fixtures__/${name}`, import.meta.url),
		'utf8',
	);

test('draws an icon on the folded Sparklines group', async ({ page }) => {
	await page.setViewportSize({ width: 480, height: 800 });
	await newWorkbook(page);
	const editor = excel(page);
	await editor.getByRole('tab', { name: 'Insert', exact: true }).click();
	const group = editor.locator('.ribbon-group[data-label="Sparklines"][data-collapsed]');
	const face = group.getByRole('button', { name: 'Sparklines', exact: true });
	await expect(face).toBeVisible();
	await expect(face.locator('svg[data-icon="xl-chartLine"] path')).toHaveAttribute('d', /\S/u);
});

test('lists only Excel compatibility details for a Word command manifest', async ({ page }) => {
	await page.setViewportSize({ width: 1600, height: 900 });
	await newWorkbook(page);
	const editor = excel(page);
	await editor.getByRole('tab', { name: 'Insert', exact: true }).click();
	await addInsCommand(editor, 'Get Add-ins');
	const dialog = editor.locator('office-ui-add-ins-dialog office-ui-dialog[open]');
	await dialog.getByLabel('Choose a local add-in file').setInputFiles({
		name: 'notes.xml',
		mimeType: 'application/xml',
		buffer: Buffer.from(await fixture('manifest.xml')),
	});
	await dialog.getByRole('button', { name: 'Install Notes', exact: true }).click();
	await addInsCommand(editor, 'My Add-ins');
	await expect(dialog).toContainText('Notes (1.0.0)');
	// The fixture's commands and their requirement belong to its Document (Word) override.
	await expect(dialog).not.toContainText('Unsupported feature');
	await expect(dialog).not.toContainText('Unavailable requirement');
});

test('reviews and installs a unified package in Excel', async ({ page }) => {
	await page.setViewportSize({ width: 1600, height: 900 });
	await newWorkbook(page);
	const editor = excel(page);
	await editor.getByRole('tab', { name: 'Insert', exact: true }).click();
	await addInsCommand(editor, 'Get Add-ins');
	const dialog = editor.locator('office-ui-add-ins-dialog office-ui-dialog[open]');
	const manifest = (await fixture('manifest.json')).replace(
		'"lifetime": "short"',
		'"lifetime": "long"',
	);
	const zip = new JSZip()
		.file('icons/outline.png', new Uint8Array([1]))
		.file('icons/color.png', new Uint8Array([2]))
		.file('fr.json', await fixture('fr.json'));
	const input = dialog.getByLabel('Choose a local add-in file');
	await expect(dialog.getByRole('searchbox')).toBeFocused();
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
	await expect(dialog).toContainText('Unavailable requirement · AddinCommands 1.10');
	await expect(dialog).toContainText('Unsupported feature · shared runtime');
	await consent.getByRole('button', { name: 'Install Notes', exact: true }).click();
	await addInsCommand(editor, 'My Add-ins');
	const installed = editor.locator('office-ui-add-ins-dialog office-ui-dialog[open]');
	await expect(installed).toContainText('Running add-ins is unavailable');
	await expect(installed).toContainText('Unavailable requirement · AddinCommands 1.10');
	await expect(installed).toContainText('Unsupported feature · shared runtime');
	await page.keyboard.press('Escape');

	await addInsCommand(editor, 'Get Add-ins');
	await select(
		manifest.replace('1.0.0', '1.0.1').replace('"short": "Notes"', '"short": "Updated Notes"'),
	);
	await expect(dialog).toContainText('Updated Notes (1.0.1)');
	await expect(dialog.getByText('Update from 1.0.0 to 1.0.1', { exact: true })).toBeVisible();
	await dialog.getByRole('button', { name: 'Update Updated Notes', exact: true }).click();

	await addInsCommand(editor, 'My Add-ins');
	await expect(installed).toContainText('Updated Notes (1.0.1)');
	await installed.getByRole('button', { name: 'Remove Updated Notes', exact: true }).click();
});

for (const framework of FRAMEWORKS) {
	test(`${framework}: opens, searches and reviews add-ins from either ribbon command`, async ({
		page,
	}) => {
		await page.setViewportSize({ width: 2400, height: 900 });
		await newWorkbook(page, framework);
		const editor = excel(page);
		await editor.getByRole('tab', { name: 'Insert', exact: true }).click();
		const get = editor.getByRole('button', { name: 'Get Add-ins', exact: true });
		const mine = editor.getByRole('button', { name: 'My Add-ins', exact: true });
		const dialog = editor.locator('office-ui-add-ins-dialog office-ui-dialog[open]');
		await get.click();
		await expect(dialog.getByRole('tab', { name: 'Store', exact: true })).toHaveAttribute(
			'aria-selected',
			'true',
		);
		await page.keyboard.press('Escape');
		await expect(dialog).toBeHidden();
		await expect(get).toBeFocused();

		await mine.press('Enter');
		const search = dialog.getByRole('searchbox');
		await expect(search).toBeFocused();
		const upload = dialog.getByRole('button', { name: 'Upload My Add-in' });
		const payload = {
			name: 'notes.xml',
			mimeType: 'application/xml',
			buffer: Buffer.from(await fixture('manifest.xml')),
		};
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

		await mine.press('Enter');
		await expect(dialog).toContainText('Notes (1.0.0)');
		await search.fill('unmatched');
		await expect(dialog).toContainText('No matching add-ins.');
		await search.fill('notes');
		await expect(dialog).toContainText('Notes (1.0.0)');
		await dialog.getByRole('button', { name: 'Close', exact: true }).click();
		await expect(dialog).toBeHidden();
		await expect(mine).toBeFocused();
	});
}
