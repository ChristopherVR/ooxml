/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec, `test`/`expect` come from @playwright/test */
/**
 * Enable Content is each participant's own choice. Two peers share a room whose deck links
 * pictures and a video to an external origin: both open without contacting it and both show the
 * security bar. The host's Enable Content makes only the host's browser fetch, the peer keeps its
 * bar and sends nothing, and the session keeps syncing edits both ways; the peer then enables for
 * itself. Nothing reloads, so the deck, its edits and the room stay as they were. Runs on every
 * binding, with the full viewer and the custom host shell.
 */
import { expect, test } from '@playwright/test';
import type { Page } from '@playwright/test';

import {
	EXTERNAL_ORIGIN,
	LINKED_PICTURES,
	externalContentDeck,
} from './support/external-content-deck';
import { hostOwnedSessionUrl } from './support/host-owned-session';

const BLUE_PNG = Buffer.from(
	'iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAIAAAD91JpzAAAAEklEQVR4nGNkaGBgYGBgYGAAAAN4AE0r1L6tAAAAAElFTkSuQmCC',
	'base64',
);

/** Answer the external origin offline and record the path of every request `page` sends it. */
async function recordExternalRequests(page: Page): Promise<string[]> {
	const paths: string[] = [];
	await page.route(`${EXTERNAL_ORIGIN}/**`, (route) => {
		paths.push(new URL(route.request().url()).pathname);
		return route.request().url().endsWith('.png')
			? route.fulfill({ contentType: 'image/png', body: BLUE_PNG })
			: route.fulfill({ status: 404, body: '' });
	});
	return paths;
}

const bar = (page: Page) => page.locator('pptx-ui-external-content-bar');
const sessionState = (page: Page) => page.getByLabel('Host session state', { exact: true });
const elements = (page: Page) => page.locator('[data-pptx-viewport] [data-pptx-element="true"]');

async function deleteLastElement(page: Page): Promise<void> {
	const ids = await elements(page).evaluateAll((nodes) =>
		nodes.map((node) => node.getAttribute('data-element-id')),
	);
	await page
		.locator(`[data-pptx-viewport] [data-element-id="${ids.at(-1)}"]`)
		.first()
		.click({ force: true });
	await page.keyboard.press('Delete');
}

test('Enable Content is per participant and keeps the collaboration session', async ({
	page,
	browser,
	baseURL,
}, info) => {
	test.setTimeout(180_000);
	const deck = await externalContentDeck();
	const context = await browser.newContext({ baseURL, viewport: { width: 1440, height: 1000 } });
	const peer = await context.newPage();
	try {
		for (const participant of [page, peer]) {
			await participant.route('**/sample-deck.pptx', (route) =>
				route.fulfill({ contentType: 'application/octet-stream', body: deck }),
			);
		}
		const hostRequests = await recordExternalRequests(page);
		const peerRequests = await recordExternalRequests(peer);
		const room = `linked-content-${info.project.name}-${Date.now()}`;
		await page.goto(hostOwnedSessionUrl(room, { sample: '1' }));
		await expect(sessionState(page)).toContainText('Host: connected', { timeout: 30_000 });
		await expect(elements(page).first()).toBeVisible({ timeout: 30_000 });
		await peer.goto(hostOwnedSessionUrl(room, { name: 'Peer' }));
		await expect(sessionState(peer)).toContainText('Host: connected', { timeout: 30_000 });
		const count = await elements(page).count();
		await expect(elements(peer)).toHaveCount(count, { timeout: 30_000 });

		// Both open blocked: the room carries the links, neither browser follows them.
		await expect(bar(page)).toBeVisible({ timeout: 30_000 });
		await expect(bar(peer)).toBeVisible({ timeout: 30_000 });
		expect(hostRequests).toEqual([]);
		expect(peerRequests).toEqual([]);

		await bar(page).getByRole('button', { name: 'Enable Content' }).click();
		await expect.poll(() => hostRequests).toContain(LINKED_PICTURES.shapeFill);
		await expect(bar(page)).toHaveCount(0);
		// The peer's choice is its own: its bar stays and it sends nothing.
		await page.waitForTimeout(2_000);
		await expect(bar(peer)).toBeVisible();
		expect(peerRequests).toEqual([]);
		await expect(elements(page)).toHaveCount(count);
		await expect(elements(peer)).toHaveCount(count);

		// The session still syncs after Enable Content.
		await deleteLastElement(page);
		await expect(elements(page)).toHaveCount(count - 1);
		await expect(elements(peer)).toHaveCount(count - 1, { timeout: 15_000 });
		expect(peerRequests).toEqual([]);

		await bar(peer).getByRole('button', { name: 'Enable Content' }).click();
		await expect.poll(() => peerRequests).toContain(LINKED_PICTURES.shapeFill);
		await expect(bar(peer)).toHaveCount(0);
		await expect(elements(peer)).toHaveCount(count - 1);
		await expect(elements(page)).toHaveCount(count - 1);
		await expect(bar(page)).toHaveCount(0);
	} finally {
		await context.close();
	}
});
