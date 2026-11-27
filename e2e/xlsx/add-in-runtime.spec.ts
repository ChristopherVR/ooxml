/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright API */
import { expect, test } from '@playwright/test';
import { jsonManifest } from '../add-in-fixtures/manifest';
import { demoReady, editor as excel, newWorkbook } from './helpers';
import {
	addInReviewScenarios,
	checkAddInRuntime,
	checkTabClipping,
	prepareAddInDemo,
} from '../docx/add-in-runtime-flow';

for (const scenario of addInReviewScenarios) {
	test(`Excel add-ins at ${scenario.width}px in ${scenario.scheme} mode`, async ({ page }) => {
		test.setTimeout(75_000);
		await prepareAddInDemo(page, scenario);
		await newWorkbook(page);
		const editor = excel(page);
		await checkAddInRuntime({
			page,
			editor,
			host: 'Excel',
			async reloadEditor() {
				await demoReady(page);
				await newWorkbook(page);
			},
			async openDialog() {
				await editor.getByRole('tab', { name: 'Home', exact: true }).click();
				const home = editor.locator('button[data-command="home.add-ins"]');
				if (!(await home.isVisible()))
					await editor.getByRole('button', { name: 'Add-ins', exact: true }).click();
				await home.click();
			},
		});
	});
}

for (const scheme of ['light', 'dark'] as const) {
	test(`Excel clips translated tab labels cleanly in ${scheme} mode`, async ({ page }) => {
		await prepareAddInDemo(page, { width: 800, height: 700, scheme });
		await newWorkbook(page);
		const editor = excel(page);
		const home = editor.locator('button[data-command="home.add-ins"]');
		if (!(await home.isVisible()))
			await editor.getByRole('button', { name: 'Add-ins', exact: true }).click();
		await home.click();
		const dialog = editor.locator('office-ui-add-ins-dialog office-ui-dialog[open]');
		await dialog.getByLabel('Choose a local add-in file').setInputFiles({
			name: 'manifest.json',
			mimeType: 'application/json',
			buffer: Buffer.from(jsonManifest),
		});
		await dialog.getByRole('button', { name: 'Install Notes' }).click();
		const notes = editor.getByRole('tab', { name: 'Notes', exact: true });
		await expect(notes).toBeFocused();
		await editor.evaluate((node) => node.setAttribute('locale', 'de-DE'));
		await expect(editor.getByRole('button', { name: 'Weitere Registerkarten' })).toBeVisible();
		await notes.focus();
		await expect(notes).toBeFocused();
		await checkTabClipping(editor, 'left');
		await editor.evaluate((node) => node.setAttribute('locale', 'fr-FR'));
		await expect(editor.getByRole('button', { name: 'Autres onglets' })).toBeVisible();
	});
}
