/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec */
/**
 * Framework-neutral joins of a shared session: whoever joins sees the deck the sharer sees,
 * however soon after the sharer they arrive, and the sharer keeps it. A joiner opens the viewer
 * on a blank placeholder deck until the room's slides arrive; that placeholder must never become
 * the room's deck.
 */
import { expect, test } from '@playwright/test';
import type { Page } from '@playwright/test';

import { webrtcRoomPath } from './support/webrtc-room';

const slideElements = (page: Page) =>
	page.locator('[data-pptx-viewport] [data-pptx-element="true"]');
const collaborationReady = (page: Page) =>
	page.getByRole('status', { name: 'Collaboration: Connected', exact: true });
const sharerTitle = (page: Page) => slideElements(page).filter({ hasText: 'Product Overview' });

async function expectSharerDeck(sharer: Page, peer: Page): Promise<void> {
	for (const participant of [sharer, peer]) {
		await expect(collaborationReady(participant)).toBeVisible({ timeout: 30_000 });
		await expect(sharerTitle(participant)).toBeVisible({ timeout: 30_000 });
	}
	// Past the first-sync grace window, both still show the sharer's deck.
	await sharer.waitForTimeout(5_000);
	await expect(sharerTitle(sharer)).toBeVisible();
	await expect(sharerTitle(peer)).toBeVisible();
	const count = await slideElements(sharer).count();
	await expect.poll(() => slideElements(peer).count()).toBe(count);
}

test.describe('collaboration join', () => {
	test.setTimeout(90_000);

	test('a peer that reaches the room before the sharer keeps the sharer deck on both pages', async ({
		page,
	}, testInfo) => {
		const peer = await page.context().newPage();
		const room = `join-early-${testInfo.project.name}-${Date.now()}`;
		try {
			await peer.goto(webrtcRoomPath(room, 'name=peer&'));
			await page.waitForTimeout(1_000);
			await page.goto(webrtcRoomPath(room, 'sample=1&name=host&'));
			await expectSharerDeck(page, peer);
		} finally {
			await peer.close();
		}
	});

	test('a peer joining while the sharer starts keeps the sharer deck on both pages', async ({
		page,
	}, testInfo) => {
		const peer = await page.context().newPage();
		const room = `join-race-${testInfo.project.name}-${Date.now()}`;
		try {
			await Promise.all([
				page.goto(webrtcRoomPath(room, 'sample=1&name=host&')),
				peer.goto(webrtcRoomPath(room, 'name=peer&')),
			]);
			await expectSharerDeck(page, peer);
		} finally {
			await peer.close();
		}
	});
});
