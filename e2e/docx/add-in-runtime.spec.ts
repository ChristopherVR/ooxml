/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright API */
import { expect, test } from '@playwright/test';
import { xmlManifest } from '../add-in-fixtures/manifest';
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
			reloadEditor: () => newDocument(page),
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
});
