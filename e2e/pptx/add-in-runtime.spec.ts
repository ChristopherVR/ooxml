/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright API */
import { expect, test } from '@playwright/test';

import { jsonManifest } from '../add-in-fixtures/manifest';
import {
	addInReviewScenarios,
	checkAddInRuntime,
	prepareAddInDemo,
} from '../docx/add-in-runtime-flow';
import {
	elementsOfType,
	fixture,
	loadDeck,
	ribbonTab,
	SAMPLE_DECK,
	selectElement,
	slideStage,
} from './support/deck';

for (const scenario of addInReviewScenarios) {
	test(`PowerPoint add-ins at ${scenario.width}px in ${scenario.scheme} mode`, async ({ page }) => {
		test.setTimeout(90_000);
		await prepareAddInDemo(page, scenario);
		await loadDeck(page);
		const editor = page.locator('body');
		if (scenario.width === 1280) {
			await ribbonTab(page, 'Insert').click();
			const group = page.locator('pptx-ui-ribbon-insert [data-ribbon-group="insert.addIns"]');
			await expect(group.getByText('Get Add-ins', { exact: true })).toBeVisible();
			await expect(group.getByText('My Add-ins', { exact: true })).toBeVisible();
		}
		await checkAddInRuntime({
			page,
			editor,
			host: 'PowerPoint',
			async reloadEditor() {
				const upload = page.locator('#file-input');
				const stage = slideStage(page);
				await expect(stage.or(upload).first()).toBeAttached();
				if (await upload.count()) await upload.setInputFiles(SAMPLE_DECK);
				await expect(stage).toBeVisible();
			},
			async openDialog() {
				const homeTab = ribbonTab(page, 'Home');
				if ((await homeTab.getAttribute('aria-selected')) !== 'true') await homeTab.click();
				const group = page.locator('[data-ribbon-group="home.addIns"]:visible');
				const command = group.locator('button[data-ribbon-control="home.addIns.addIns"]');
				await expect(async () => {
					if (!(await command.isVisible())) {
						await group
							.getByRole('button', { name: 'Add-ins', exact: true })
							.click({ timeout: 1_000 });
					}
					await command.click({ timeout: 2_000 });
				}).toPass({ timeout: 10_000 });
			},
		});
	});
}

test('a picture selection takes precedence over the active add-in ribbon tab', async ({ page }) => {
	await loadDeck(page, fixture('ribbon-galleries.pptx'));
	await ribbonTab(page, 'Insert').click();
	const group = page.locator('[data-ribbon-group="insert.addIns"]');
	await group.getByText('My Add-ins', { exact: true }).click();
	const dialog = page.locator('office-ui-add-ins-dialog office-ui-dialog[open]');
	await dialog.getByLabel('Choose a local add-in file').setInputFiles({
		name: 'manifest.json',
		mimeType: 'application/json',
		buffer: Buffer.from(jsonManifest),
	});
	await dialog.getByRole('button', { name: 'Install Notes' }).click();
	const notes = ribbonTab(page, 'Notes');
	await notes.click();
	await expect(notes).toHaveAttribute('aria-selected', 'true');
	await selectElement(page, elementsOfType(page, 'image').first());
	await expect(ribbonTab(page, 'Picture Format')).toHaveAttribute('aria-selected', 'true');
	await expect(notes).toHaveAttribute('aria-selected', 'false');
});
