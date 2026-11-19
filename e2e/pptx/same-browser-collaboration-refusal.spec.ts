/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec */
/**
 * A same-browser room that receives an update it cannot read stops sharing and says why, in
 * every binding, without changing the deck.
 */
import { expect, test } from '@playwright/test';
import type { Page } from '@playwright/test';

const elements = (page: Page) => page.locator('[data-pptx-viewport] [data-pptx-element="true"]');
const state = (page: Page) => page.getByLabel('Same-browser collaboration', { exact: true });

/** An update whose struct count no well-formed update of its size can hold. */
const UNREADABLE = [0, 2, 5, 1, 1, 1, 0, 0x1f];

test.use({ viewport: { width: 1440, height: 1000 } });

test('a refused update stops a same-browser room with the reason', async ({ page }, info) => {
	test.setTimeout(120_000);
	const room = `refuse-${info.project.name}-${Date.now()}`;
	await page.goto(`/?${new URLSearchParams({ sameBrowser: room, name: 'Alice' })}`);
	await expect(elements(page).first()).toBeVisible({ timeout: 30_000 });
	await expect(state(page)).toContainText('synced: true', { timeout: 30_000 });
	const before = await elements(page).allInnerTexts();

	await page.evaluate(
		({ name, bytes }) => {
			const channel = new BroadcastChannel(`ooxml-pptx:${name}`);
			channel.postMessage({ kind: 'data', from: 'other-window', data: new Uint8Array(bytes) });
			channel.close();
		},
		{ name: room, bytes: UNREADABLE },
	);

	await expect(
		page.getByText('A change from the shared session could not be read', { exact: false }).first(),
	).toBeVisible({ timeout: 10_000 });
	expect(await elements(page).allInnerTexts()).toEqual(before);
});
