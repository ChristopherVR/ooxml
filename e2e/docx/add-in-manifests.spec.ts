/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright API */
import { readFile } from 'node:fs/promises';
import { expect, test } from '@playwright/test';
import { newDocument, reveal } from './helpers';

const fixture = new URL(
	'../../src/core/automation/add-ins/__fixtures__/manifest.xml',
	import.meta.url,
);

for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid']) {
	test(`${framework}: reviews, cancels and installs an add-in from My Add-ins`, async ({
		page,
	}) => {
		await page.setViewportSize({ width: 2400, height: 900 });
		await page.goto(`/?framework=${framework}`);
		await newDocument(page);
		const editor = page.locator('docx-editor');
		await editor.getByRole('tab', { name: 'Insert', exact: true }).click();
		const mine = editor.getByRole('button', { name: 'My Add-ins', exact: true });
		const get = editor.getByRole('button', { name: 'Get Add-ins', exact: true });
		const dialog = editor.locator('office-ui-add-ins-dialog office-ui-dialog[open]');
		await reveal(editor, get);
		await get.click();
		await expect(dialog.getByRole('tab', { name: 'Store', exact: true })).toHaveAttribute(
			'aria-selected',
			'true',
		);
		await page.keyboard.press('Escape');
		await expect(dialog).toBeHidden();
		await expect(get).toBeFocused();
		await reveal(editor, mine);
		await mine.press('Enter');
		const search = dialog.getByRole('searchbox', { name: 'Search add-ins' });
		await expect(search).toBeFocused();
		const initialBox = await dialog.getByRole('dialog', { name: 'Office Add-ins' }).boundingBox();
		const upload = dialog.getByRole('button', { name: 'Upload My Add-in' });
		const chooser = page.waitForEvent('filechooser');
		await upload.click();
		await (
			await chooser
		).setFiles({ name: 'notes.xml', mimeType: 'application/xml', buffer: await readFile(fixture) });
		const consent = dialog.getByRole('region', { name: 'Review add-in permissions' });
		await expect(consent).toContainText('Read and change this document');
		expect(await dialog.getByRole('dialog', { name: 'Office Add-ins' }).boundingBox()).toEqual(
			initialBox,
		);
		await consent.getByRole('button', { name: 'Cancel' }).click();
		await expect(search).toBeFocused();
		await expect(dialog).toContainText('No add-ins');
		await dialog.getByLabel('Choose a local add-in file').setInputFiles({
			name: 'notes.xml',
			mimeType: 'application/xml',
			buffer: await readFile(fixture),
		});
		await consent.getByRole('button', { name: 'Install Notes' }).click();
		await expect(dialog).toBeHidden();
		await expect(mine).toBeFocused();
		await mine.press('Enter');
		await expect(dialog).toContainText('Notes (1.0.0)');
		await search.fill('unmatched');
		await expect(dialog).toContainText('No matching add-ins.');
		await search.fill('NOTES');
		await expect(dialog).toContainText('Notes (1.0.0)');
		await dialog.getByRole('button', { name: 'Close', exact: true }).click();
		await expect(dialog).toBeHidden();
		await expect(mine).toBeFocused();
	});
}

test('opens the Marketplace only on click and supports turning off the Store', async ({
	page,
	context,
}) => {
	await page.setViewportSize({ width: 2400, height: 900 });
	const marketplace: Array<{
		url: string;
		method: string;
		referrer?: string;
		body: string | null;
	}> = [];
	await context.route('https://marketplace.microsoft.com/**', async (route) => {
		const request = route.request();
		const referrer = request.headers().referer;
		marketplace.push({
			url: request.url(),
			method: request.method(),
			body: request.postData(),
			...(referrer && { referrer }),
		});
		await route.fulfill({ body: 'Marketplace' });
	});
	await page.goto('/?framework=vanilla');
	await newDocument(page);
	const editor = page.locator('docx-editor');
	await editor.getByRole('tab', { name: 'Insert', exact: true }).click();
	const get = editor.getByRole('button', { name: 'Get Add-ins', exact: true });
	await reveal(editor, get);
	await get.click();
	const dialog = editor.locator('office-ui-add-ins-dialog office-ui-dialog[open]');
	const link = dialog.getByRole('link', { name: 'Browse Microsoft Marketplace' });
	await expect(link).toBeVisible();
	expect(marketplace).toHaveLength(0);
	const popup = page.waitForEvent('popup');
	await link.click();
	const tab = await popup;
	await expect(tab.getByText('Marketplace', { exact: true })).toBeVisible();
	expect(marketplace).toEqual([
		{
			url: 'https://marketplace.microsoft.com/en-us/marketplace/apps?product=office',
			method: 'GET',
			body: null,
		},
	]);
	await tab.close();
	await dialog.getByRole('button', { name: 'Close', exact: true }).click();
	await expect(dialog).toBeHidden();
	await editor.evaluate((element) =>
		Reflect.set(element, Symbol.for('office-ui.add-in-storefront'), 'off'),
	);
	await get.click();
	await expect(dialog.getByRole('tab', { name: 'My Add-ins' })).toHaveAttribute(
		'aria-selected',
		'true',
	);
	await expect(dialog.getByRole('tab', { name: 'Store', exact: true })).toHaveCount(0);
	await expect(link).toHaveCount(0);
});

test('loads Office Add-ins on first open and fits the dark editor on a phone', async ({ page }) => {
	await page.setViewportSize({ width: 2400, height: 900 });
	const manifestCode = '**/assets/add-in-manifest-file-*.js';
	let requestedManifestCode = false;
	await page.route(manifestCode, (route) => {
		requestedManifestCode = true;
		return route.continue();
	});
	let release = () => {};
	let requested = false;
	const gate = new Promise<void>((resolve) => {
		release = resolve;
	});
	await page.route('**/assets/add-ins-dialog-*.js', async (route) => {
		requested = true;
		await gate;
		await route.continue();
	});
	try {
		await page.goto('/?framework=vanilla');
		await newDocument(page);
		const editor = page.locator('docx-editor');
		await editor.getByRole('tab', { name: 'Insert', exact: true }).click();
		const get = editor.getByRole('button', { name: 'Get Add-ins', exact: true });
		await reveal(editor, get);
		expect(requested).toBe(false);
		await get.press('Enter');
		await expect.poll(() => requested).toBe(true);
		const dialog = editor.getByRole('dialog', { name: 'Office Add-ins' });
		await expect(dialog).toHaveCount(0);
		release();
		await expect(dialog).toBeVisible();
		expect(requestedManifestCode).toBe(false);
		const background = () => dialog.evaluate((box) => getComputedStyle(box).backgroundColor);
		const light = await background();
		await editor.evaluate((element) => element.setAttribute('theme', 'dark'));
		await expect.poll(background).not.toBe(light);
		await page.setViewportSize({ width: 390, height: 460 });
		const box = await dialog.boundingBox();
		expect(box?.width).toBeLessThan(390);
		expect(box?.height).toBeLessThan(460);
		await expect(
			editor
				.locator('office-ui-add-ins-dialog')
				.getByRole('button', { name: 'Close', exact: true }),
		).toBeVisible();
		const validator = page.waitForRequest(manifestCode);
		const shell = editor.locator('office-ui-add-ins-dialog office-ui-dialog[open]');
		await shell.getByLabel('Choose a local add-in file').setInputFiles({
			name: 'notes.xml',
			mimeType: 'application/xml',
			buffer: await readFile(fixture),
		});
		await validator;
		await expect(shell.getByRole('region', { name: 'Review add-in permissions' })).toBeVisible();
	} finally {
		release();
	}
});

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
	const preview = editor.locator('office-ui-add-ins-dialog office-ui-dialog[open]');
	const file = preview.getByLabel('Choose a local add-in file');
	await expect(preview.getByRole('searchbox')).toBeFocused();
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
	const install = consent.getByRole('button', { name: 'Install Notes', exact: true });
	await expect(consent.getByRole('heading')).toBeFocused();
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
	const installed = editor.locator('office-ui-add-ins-dialog office-ui-dialog[open]');
	await expect(installed).toContainText('Notes (1.0.0)');
	await expect(installed).toContainText('Running add-ins is unavailable');
	await expect(installed).toContainText('Unavailable requirement · WordApi 1.1');
	await expect(installed).toContainText('Unavailable requirement · AddinCommands 1.10');
	await expect(installed).toContainText('Unsupported feature · ExecuteFunction');
	await expect(installed).not.toContainText('ContextMenu');
	await page.keyboard.press('Escape');

	await reveal(editor, get);
	await get.click();
	await select(source.replace('DefaultValue="Notes"', 'DefaultValue="Revised Notes"'));
	await expect(
		preview.getByText('Update this add-in and keep version 1.0.0', { exact: true }),
	).toBeVisible();
	await preview.getByRole('button', { name: 'Update Revised Notes', exact: true }).click();

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
