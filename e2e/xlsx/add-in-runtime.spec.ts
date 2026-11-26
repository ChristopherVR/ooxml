/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright API */
import { test } from '@playwright/test';
import { enforcePublisherFramePolicy } from '../add-in-fixtures/manifest';
import { editor as excel, newWorkbook } from './helpers';
import { checkAddInRuntime } from '../docx/add-in-runtime-flow';

test('installs, runs, and removes XML, JSON and ZIP task panes in Excel', async ({ page }) => {
	await enforcePublisherFramePolicy(page);
	await page.setViewportSize({ width: 2400, height: 900 });
	await newWorkbook(page);
	const editor = excel(page);
	await checkAddInRuntime({
		page,
		editor,
		host: 'Excel',
		async openDialog() {
			await editor.getByRole('tab', { name: 'Insert', exact: true }).click();
			const mine = editor.getByRole('button', { name: 'My Add-ins', exact: true });
			if (!(await mine.isVisible()))
				await editor.getByRole('button', { name: 'Add-ins', exact: true }).click();
			await mine.click();
		},
	});
});
