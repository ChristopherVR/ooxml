/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright owns test and expect */
import { readFileSync } from 'node:fs';

import { expect, test } from '@playwright/test';

import { withNativePictureLayout } from '../../src/core/diagram/test-support/picture-layout';
import { savePptxViaBackstage } from './save-pptx';
import { resetTabSession, slideStage, loadDeck } from './support/deck';

for (const layout of ['picture-accent-list--flat3', 'picture-grid--flat3']) {
	test(`${layout} keeps image fills after caption edit, reflow and save/reopen`, async ({
		page,
	}) => {
		const bytes = await withNativePictureLayout(
			readFileSync(
				new URL(
					`../../src/core/pptx/__tests__/fixtures/smartart-gallery/${layout}.pptx`,
					import.meta.url,
				),
			),
			'ppt',
			layout,
		);
		await resetTabSession(page);
		await page.goto('/');
		await page
			.locator('#file-input')
			.setInputFiles({
				name: `${layout}.pptx`,
				mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
				buffer: Buffer.from(bytes),
			});
		const stage = slideStage(page);
		await expect(stage.locator('image')).toHaveCount(3);
		const caption = stage
			.locator('[data-smartart-node-id] text')
			.filter({ hasText: /^Alpha$/ })
			.first();
		await caption.dblclick();
		const editor = page.locator('[data-pptx-viewport] textarea:visible');
		await expect(editor).toHaveCount(1);
		await editor.fill('Updated picture caption');
		await page
			.getByRole('toolbar', { name: 'Presentation toolbar' })
			.getByRole('tab', { name: 'Home', exact: true })
			.click();
		await expect(editor).toHaveCount(0);
		await expect(stage.locator('image')).toHaveCount(3);
		await expect(stage).toContainText('Updated picture caption');
		const download = await savePptxViaBackstage(page);
		const path = await download.path();
		if (!path) throw new Error('Saved deck download missing');
		await loadDeck(page, path);
		await expect(slideStage(page).locator('image')).toHaveCount(3);
		await expect(slideStage(page)).toContainText('Updated picture caption');
	});
}
