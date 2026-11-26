/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright API */
import { expect, test } from '@playwright/test';

import {
	addInReviewScenarios,
	checkAddInRuntime,
	prepareAddInDemo,
} from '../docx/add-in-runtime-flow';
import { loadDeck, ribbonTab } from './support/deck';

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
			reloadEditor: () => loadDeck(page),
			async openDialog() {
				await ribbonTab(page, 'Home').click();
				const home = page.locator('button[data-ribbon-control="home.editing.addIns"]');
				if (!(await home.isVisible()))
					await page.getByRole('button', { name: 'Editing', exact: true }).click();
				await home.click();
			},
		});
	});
}
