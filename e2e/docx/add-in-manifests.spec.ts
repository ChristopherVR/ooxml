/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright API */
import { readFile } from 'node:fs/promises';
import { expect, test } from '@playwright/test';
import { newDocument, reveal } from './helpers';

const fixture = new URL(
	'../../src/core/automation/add-ins/__fixtures__/manifest.xml',
	import.meta.url,
);

test('consents to a local Word add-in and preserves it after an invalid update', async ({
	page,
}) => {
	await page.setViewportSize({ width: 2400, height: 900 });
	const requests: string[] = [];
	await page.route('https://*.example.test/**', async (route) => {
		requests.push(route.request().url());
		await route.abort();
	});
	await page.goto('/?framework=vanilla');
	await newDocument(page);
	const editor = page.locator('docx-editor');
	await editor.getByRole('tab', { name: 'Insert', exact: true }).click();
	const get = editor.getByRole('button', { name: 'Get Add-ins', exact: true });
	const mine = editor.getByRole('button', { name: 'My Add-ins', exact: true });
	await reveal(editor, get);
	await get.press('Enter');
	const preview = editor.locator('office-ui-dialog[open]');
	const file = preview.getByLabel('Choose a local add-in file');
	await expect(file).toBeFocused();
	const marketplace = preview.getByRole('link', { name: /Microsoft Marketplace/ });
	await expect(marketplace).toBeVisible();
	const source = (await readFile(fixture, 'utf8'))
		.replace(
			'<DefaultSettings>',
			'<Requirements><Sets><Set Name="WordApi" MinVersion="1.1" /></Sets></Requirements><DefaultSettings>',
		)
		.replace(
			'</Host>',
			`</Host><Host xsi:type="Workbook"><DesktopFormFactor>
				<ExtensionPoint xsi:type="ContextMenu" />
			</DesktopFormFactor></Host>`,
		);
	const select = (xml: string) =>
		file.setInputFiles({
			name: 'notes.xml',
			mimeType: 'application/xml',
			buffer: Buffer.from(xml),
		});
	await select(source);
	const consent = preview.getByRole('region', { name: 'Review add-in permissions' });
	await expect(consent.getByRole('heading')).toHaveText('Review add-in permissions');
	await expect(consent).toContainText('This add-in requests:');
	await expect(consent).toContainText('Read and change this document');
	await expect(consent).toContainText('Lab Notes');
	await expect(consent).toContainText('https://notes.example.test');
	await expect(consent).toContainText('https://assets.example.test');
	await expect(consent).not.toContainText('Document.ReadWrite.User');
	await expect(marketplace).toBeHidden();
	await expect(
		preview.getByText('Browse add-ins for Microsoft Office.', { exact: false }),
	).toBeHidden();
	const install = consent.getByRole('button', { name: 'Install Notes', exact: true });
	await expect(install).toBeFocused();
	await install.press('Enter');
	await expect(preview).toBeHidden();
	if (await get.isVisible()) await expect(get).toBeFocused();

	await reveal(editor, get);
	await get.click();
	await select(source.replace('https://notes.example.test/pane', 'javascript:alert(1)'));
	await expect(preview.getByRole('status')).toContainText('unsafe URL');
	await page.keyboard.press('Escape');

	await reveal(editor, mine);
	await mine.click();
	const installed = editor.locator('office-ui-dialog[open]');
	await expect(installed).toContainText('Notes (1.0.0)');
	await expect(installed).toContainText('Running add-ins is unavailable');
	await expect(installed).toContainText('Unavailable requirement · WordApi 1.1');
	await expect(installed).toContainText('Unavailable requirement · AddinCommands 1.10');
	await expect(installed).toContainText('Unsupported feature · ExecuteFunction');
	await expect(installed).not.toContainText('ContextMenu');
	await page.keyboard.press('Escape');

	await reveal(editor, get);
	await get.click();
	await select(
		source
			.replace('1.0.0', '1.0.1')
			.replace('DefaultValue="Notes"', 'DefaultValue="Updated Notes"'),
	);
	await expect(preview).toContainText('Updated Notes (1.0.1)');
	await expect(preview.getByText('Update from 1.0.0 to 1.0.1', { exact: true })).toBeVisible();
	await preview.getByRole('button', { name: 'Update Updated Notes', exact: true }).click();

	await reveal(editor, mine);
	await mine.click();
	await expect(installed).toContainText('Updated Notes (1.0.1)');
	await installed.getByRole('button', { name: 'Remove Updated Notes', exact: true }).click();
	await reveal(editor, mine);
	await mine.click();
	await expect(installed).toContainText('No add-ins');
	expect(requests).toEqual([]);
});
