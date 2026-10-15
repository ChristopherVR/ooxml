/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec */
import { fileURLToPath } from 'node:url';

import { expect, test } from '@playwright/test';
import { PptxHandler } from 'pptx-viewer-core';

import { savePptxViaBackstage } from './save-pptx';
import { inspector, loadDeck, openRibbonTab, slideElements, slideStage } from './support/deck';
import { downloadBytes } from './support/exports';
import { chooseSelectValue } from './support/select-control';

// Share the core's small, committed reproduction deck instead of duplicating it.
const DECK = fileURLToPath(
	new URL('../../src/core/pptx/__tests__/fixtures/slide-size-centering.pptx', import.meta.url),
);

test.use({ viewport: { width: 1440, height: 900 } });

for (const { choice, frame } of [
	{ choice: 'ensure-fit', frame: { x: 340, y: 210, width: 600, height: 240 } },
	{ choice: 'maximize', frame: { x: 240, y: 160, width: 800, height: 320 } },
]) {
	test(`slide-size ${choice} keeps content visible and saves pixel coordinates`, async ({
		page,
	}, testInfo) => {
		await loadDeck(page, DECK);
		await openRibbonTab(page, 'Design');
		await page.getByRole('button', { name: 'Slide Size', exact: true }).first().click();
		await chooseSelectValue(
			page,
			inspector(page).locator('[data-pptx-slide-size-preset]'),
			'widescreen',
		);
		await page.getByTestId(`pptx-slide-size-rescale-${choice}`).click();

		const caption = slideElements(page).filter({ hasText: 'This content must stay visible' });
		await expect
			.poll(async () => {
				const content = await caption.boundingBox();
				const stage = await slideStage(page).boundingBox();
				return (
					content !== null &&
					stage !== null &&
					content.x >= stage.x - 1 &&
					content.y >= stage.y - 1 &&
					content.x + content.width <= stage.x + stage.width + 1 &&
					content.y + content.height <= stage.y + stage.height + 1
				);
			})
			.toBe(true);
		await testInfo.attach('resized-slide', {
			body: await page.screenshot(),
			contentType: 'image/png',
		});

		const bytes = await downloadBytes(await savePptxViaBackstage(page));
		const handler = new PptxHandler();
		try {
			const saved = await handler.load(Uint8Array.from(bytes).buffer);
			expect(saved.width).toBe(1280);
			expect(saved.height).toBe(720);
			const shape = saved.slides[0].elements[0];
			for (const key of ['x', 'y', 'width', 'height'] as const) {
				expect(shape[key], key).toBeCloseTo(frame[key], 3);
			}
		} finally {
			handler.dispose();
		}
	});
}
