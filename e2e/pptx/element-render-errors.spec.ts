/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec */
/**
 * Per-element render error containment in every binding (issue #289): when one slide element's
 * renderer throws, only that element is left out; its neighbours, the slide and the editor stay,
 * and the failure is reported (`console.error(error, context)` by default). The fault comes from
 * the demos' test-only `?elementRenderFault=<id>` hook (`demos/pptx/shared/element-render-fault.ts`),
 * which makes the shared `probeElementRenderFault` throw inside each binding's element renderer.
 */
import { expect, test } from '@playwright/test';
import type { Page } from '@playwright/test';

import { loadDeck, ribbon, SAMPLE_DECK, selectElement, slideStage } from './support/deck';

const canvasElement = (page: Page, id: string) =>
	page.locator(`[data-pptx-viewport] [data-element-id="${id}"]`);

/** Ids of the first slide's canvas elements that show text, in paint order. */
async function textElementIds(page: Page): Promise<string[]> {
	return page.locator('[data-pptx-viewport] [data-element-id]').evaluateAll((nodes) => [
		...new Set(
			nodes
				.filter((node) => (node.textContent ?? '').trim().length > 0)
				.map((node) => node.getAttribute('data-element-id') ?? '')
				.filter(Boolean),
		),
	]);
}

test.describe('element render errors', () => {
	test('a throwing element is omitted alone and reported', async ({ page }) => {
		await loadDeck(page, SAMPLE_DECK);
		const ids = await textElementIds(page);
		expect(ids.length).toBeGreaterThanOrEqual(2);
		const [healthy, broken] = [ids[0]!, ids[ids.length - 1]!];
		await expect(canvasElement(page, broken).first()).toBeVisible();

		const reports: string[] = [];
		page.on('console', (message) => {
			if (message.type() === 'error' && message.text().includes('Injected element render fault')) {
				reports.push(message.text());
			}
		});
		const pageErrors: Error[] = [];
		page.on('pageerror', (error) => pageErrors.push(error));

		await loadDeck(page, SAMPLE_DECK, `/?elementRenderFault=${encodeURIComponent(broken)}`);
		await expect(slideStage(page)).toBeVisible();
		await expect(canvasElement(page, healthy).first()).toBeVisible();
		await expect(canvasElement(page, broken)).toHaveCount(0);
		await expect.poll(() => reports.some((text) => text.includes(`(${broken})`))).toBe(true);
		// The editor survives: chrome is up and the healthy element still takes a selection.
		await expect(ribbon(page)).toBeVisible();
		await selectElement(page, canvasElement(page, healthy).first());
		await expect(canvasElement(page, broken)).toHaveCount(0);
		expect(pageErrors.filter((error) => error.message.includes('Injected element'))).toEqual([]);
	});
});
