/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec */
/**
 * Framework-neutral online video: a linked YouTube video (PowerPoint's Insert > Online Video) is
 * linked external content, so nothing reaches the provider until the user chooses Enable Content
 * (docs/powerpoint-external-content.md). After it, the provider's embedded player paints, not an
 * empty frame, in every binding.
 */
import { expect, test } from '@playwright/test';
import type { PowerPointViewerAPI } from 'pptx-viewer-shared';

declare global {
	interface Window {
		/** Public component handle exposed by each development demo. */
		__pptxViewer: PowerPointViewerAPI;
	}
}

test('a linked YouTube video paints the embedded player after Enable Content', async ({ page }) => {
	// Never reach the provider from a test: answer its player page locally, and count the requests.
	const requests: string[] = [];
	await page.route('https://www.youtube.com/**', (route) => {
		requests.push(route.request().url());
		return route.fulfill({
			status: 200,
			contentType: 'text/html',
			body: '<!doctype html><title>player</title>',
		});
	});
	await page.goto('/?sample=1');
	await expect(page.locator('[data-pptx-viewport] [data-pptx-element="true"]').first()).toBeVisible(
		{ timeout: 30_000 },
	);
	await page.evaluate(() =>
		window.__pptxViewer.addElement({
			type: 'media',
			id: 'online-video',
			mediaType: 'video',
			isLinked: true,
			mediaPath: 'https://www.youtube.com/watch?v=dQw4w9WgXcQ',
			x: 100,
			y: 100,
			width: 480,
			height: 270,
		} as never),
	);
	const player = page.locator(
		'[data-pptx-viewport] iframe[src="https://www.youtube.com/embed/dQw4w9WgXcQ"]',
	);
	const bar = page.locator('pptx-ui-external-content-bar');
	await expect(bar).toBeVisible();
	await expect(player).toHaveCount(0);
	expect(requests).toStrictEqual([]);

	await bar.getByRole('button', { name: 'Enable Content' }).click();
	await expect(player).toHaveCount(1);
	await expect(player).toBeVisible();
	await expect(player).toHaveAttribute('sandbox', /allow-scripts/u);
});
