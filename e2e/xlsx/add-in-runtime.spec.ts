/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright API */
import { test } from '@playwright/test';
import { editor as excel, newWorkbook } from './helpers';
import {
	addInReviewScenarios,
	checkAddInRuntime,
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
			reloadEditor: () => newWorkbook(page),
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
