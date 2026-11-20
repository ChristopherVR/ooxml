/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec */
/**
 * Framework-neutral live broadcast: a presenter starts a peer-to-peer broadcast from the
 * Broadcast dialog, the page address becomes the audience invitation (`?broadcast=`), and an
 * audience page opened from it shows the presenter's deck and follows the presenter's slide.
 */
import { expect, test } from '@playwright/test';
import type { Page } from '@playwright/test';
import type { PowerPointViewerAPI } from 'pptx-viewer-shared';

declare global {
	interface Window {
		/** Public component handle exposed by each development demo. */
		__pptxViewer: PowerPointViewerAPI;
	}
}

const slideElements = (page: Page) =>
	page.locator('[data-pptx-viewport] [data-pptx-element="true"]');
const activeSlide = (page: Page) => page.evaluate(() => window.__pptxViewer.getActiveSlideIndex());

/** Start a serverless broadcast of the sample deck and return the audience invitation path. */
async function startBroadcast(presenter: Page): Promise<string> {
	await presenter.goto('/?sample=1');
	await expect(slideElements(presenter).filter({ hasText: 'Product Overview' })).toBeVisible({
		timeout: 30_000,
	});
	await presenter.getByRole('tab', { name: 'Slide Show', exact: true }).click();
	await presenter.getByRole('button', { name: 'Broadcast', exact: true }).click();
	const server = presenter.getByRole('dialog').getByRole('textbox', { name: /server/iu });
	await expect(server).toBeVisible();
	await server.fill('');
	await presenter.getByRole('button', { name: /^start broadcast$/iu }).click();
	await expect
		.poll(() => new URL(presenter.url()).searchParams.get('broadcast'), { timeout: 15_000 })
		.toBeTruthy();
	const invitation = new URL(presenter.url());
	invitation.searchParams.delete('sample');
	return `${invitation.pathname}${invitation.search}${invitation.hash}`;
}

test.describe('collaboration broadcast', () => {
	test.setTimeout(120_000);

	test('an audience link shows the presenter deck and follows the presenter', async ({ page }) => {
		const invitation = await startBroadcast(page);
		const audience = await page.context().newPage();
		try {
			await audience.goto(invitation);
			await expect(slideElements(audience).filter({ hasText: 'Product Overview' })).toBeVisible({
				timeout: 30_000,
			});
			await expect.poll(() => activeSlide(audience)).toBe(0);

			await page.evaluate(() => window.__pptxViewer.goTo(2));
			await expect.poll(() => activeSlide(page)).toBe(2);
			await expect.poll(() => activeSlide(audience), { timeout: 15_000 }).toBe(2);

			await page.evaluate(() => window.__pptxViewer.goTo(1));
			await expect.poll(() => activeSlide(audience), { timeout: 15_000 }).toBe(1);
		} finally {
			await audience.close();
		}
	});
});
