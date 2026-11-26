/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright API */
import { test } from '@playwright/test';

import { enforcePublisherFramePolicy } from '../add-in-fixtures/manifest';
import { checkAddInRuntime } from '../docx/add-in-runtime-flow';
import { loadDeck, ribbonTab } from './support/deck';

test('installs, runs, and removes XML, JSON and ZIP task panes in PowerPoint', async ({ page }) => {
	await enforcePublisherFramePolicy(page);
	await page.setViewportSize({ width: 2400, height: 900 });
	await loadDeck(page);
	await checkAddInRuntime({
		page,
		editor: page.locator('body'),
		host: 'PowerPoint',
		async openDialog() {
			await ribbonTab(page, 'Insert').click();
			const group = page.locator('pptx-ui-ribbon-insert [data-ribbon-group="insert.addIns"]');
			await group.getByRole('button', { name: 'My Add-ins', exact: true }).click();
		},
	});
});
