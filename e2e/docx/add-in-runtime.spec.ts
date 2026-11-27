/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright API */
import { expect, test } from '@playwright/test';
import { addInName, jsonManifest, xmlManifest } from '../add-in-fixtures/manifest';
import { newDocument, reveal } from './helpers';
import {
	addInReviewScenarios,
	checkAddInRuntime,
	enforcePublisherFramePolicy,
	prepareAddInDemo,
} from './add-in-runtime-flow';

for (const scenario of addInReviewScenarios) {
	test(`Word add-ins at ${scenario.width}px in ${scenario.scheme} mode`, async ({ page }) => {
		test.setTimeout(75_000);
		await prepareAddInDemo(page, scenario);
		await page.goto('/?framework=vanilla');
		await newDocument(page);
		const editor = page.locator('docx-editor');
		await checkAddInRuntime({
			page,
			editor,
			host: 'Word',
			async reloadEditor() {
				await page.reload();
				await newDocument(page);
			},
			async openDialog() {
				await editor.getByRole('tab', { name: 'Home', exact: true }).click();
				const home = editor.getByRole('button', { name: 'Add-ins', exact: true });
				const button = home.and(editor.locator('button[data-action]'));
				await reveal(editor, button);
				await button.click();
			},
		});
	});
}

test('reviews a Marketplace listing through the configured resolver without enabling Store', async ({
	page,
}) => {
	await page.setViewportSize({ width: 2400, height: 900 });
	await enforcePublisherFramePolicy(page);
	await page.goto('/?framework=vanilla');
	await newDocument(page);
	const editor = page.locator('docx-editor');
	await editor.evaluate((node, manifest) => {
		Reflect.set(node, Symbol.for('office-ui.add-in-marketplace-resolver'), async (id: string) => {
			node.setAttribute('data-test-asset-id', id);
			return new TextEncoder().encode(manifest);
		});
	}, xmlManifest);
	await editor.getByRole('tab', { name: 'Insert', exact: true }).click();
	const mine = editor.getByRole('button', { name: 'My Add-ins', exact: true });
	await reveal(editor, mine);
	await mine.click();
	const dialog = editor.locator('office-ui-add-ins-dialog office-ui-dialog[open]');
	const link = dialog.getByLabel('Add from Marketplace link');
	const review = dialog.getByRole('button', { name: 'Review add-in', exact: true });
	await link.fill('https://appsource.microsoft.com.example.test/en-us/product/office/WA999999999');
	await review.click();
	await expect(dialog.getByRole('alert')).toHaveText(
		'Enter a Microsoft Marketplace link or asset ID.',
	);
	await expect(editor).not.toHaveAttribute('data-test-asset-id', /.+/);
	await link.fill('https://appsource.microsoft.com/en-us/product/office/WA999999999?tab=overview');
	await review.click();
	await expect(dialog.getByRole('heading', { name: 'Review add-in permissions' })).toBeVisible();
	await expect(editor).toHaveAttribute('data-test-asset-id', 'WA999999999');
	await dialog.getByRole('button', { name: 'Cancel', exact: true }).click();
	await expect(dialog.getByRole('tab', { name: 'Store', exact: true })).toHaveCount(0);
	await expect(dialog.getByRole('button', { name: 'Open Notes', exact: true })).toHaveCount(0);
	const upload = dialog.getByRole('button', { name: 'Upload My Add-in' });
	await dialog.getByLabel('Choose a local add-in file').setInputFiles({
		name: 'broken.xml',
		mimeType: 'application/xml',
		buffer: Buffer.from('<OfficeApp>'),
	});
	const error = dialog.getByRole('alert');
	await expect(error).toHaveText('The add-in manifest is invalid.');
	await expect(error).toBeInViewport();
	await expect(upload).toHaveAttribute('aria-describedby', 'add-in-error');
});

test('explains why an OfficeTab command is disabled', async ({ page }) => {
	await page.goto('/?framework=vanilla');
	await newDocument(page);
	const editor = page.locator('docx-editor');
	const button = editor
		.getByRole('button', { name: 'Add-ins', exact: true })
		.and(editor.locator('button[data-action]'));
	await reveal(editor, button);
	await button.click();
	const dialog = editor.locator('office-ui-add-ins-dialog office-ui-dialog[open]');
	await dialog.getByLabel('Choose a local add-in file').setInputFiles({
		name: 'home.xml',
		mimeType: 'application/xml',
		buffer: Buffer.from(
			xmlManifest.replaceAll('CustomTab', 'OfficeTab').replace('id="NotesTab"', 'id="TabHome"'),
		),
	});
	const command = dialog.getByRole('button', { name: 'Open Notes', exact: true });
	await expect(command).toBeDisabled();
	await expect(command).toHaveAccessibleDescription(
		'This editor cannot add commands to built-in ribbon tabs.',
	);
});

test('a session-only add-in survives moving the editor to another parent', async ({ page }) => {
	await page.goto('/?framework=vanilla&addInsStorage=off');
	await newDocument(page);
	const editor = page.locator('docx-editor');
	const command = editor
		.getByRole('button', { name: 'Add-ins', exact: true })
		.and(editor.locator('button[data-action]'));
	await reveal(editor, command);
	await command.click();
	const dialog = editor.locator('office-ui-add-ins-dialog office-ui-dialog[open]');
	await dialog.getByLabel('Choose a local add-in file').setInputFiles({
		name: 'manifest.json',
		mimeType: 'application/json',
		buffer: Buffer.from(jsonManifest),
	});
	await dialog.getByRole('button', { name: 'Install Notes' }).click();
	const tab = editor.getByRole('tab', { name: addInName, exact: true });
	await expect(tab).toBeAttached();
	await editor.evaluate((node) => {
		const parent = node.ownerDocument.createElement('div');
		node.before(parent);
		parent.append(node);
	});
	await expect(tab).toBeAttached();
	await tab.click();
	await expect(editor.getByRole('button', { name: 'Open Notes' })).toBeVisible();
});
