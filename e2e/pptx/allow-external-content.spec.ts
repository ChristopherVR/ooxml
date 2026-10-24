/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec, `test`/`expect` come from @playwright/test */
/**
 * The host `allowExternalContent` switch, driven through every binding by the
 * demos' `?allowExternalContent=0|1` query option (see
 * `demos/pptx/shared/demo-root-options.ts`). The deck links a picture and a
 * video to an external origin (answered here with `page.route`), so opening it
 * makes the browser contact that origin unless external content is blocked.
 */
import { expect, test } from '@playwright/test';
import type { Locator, Page } from '@playwright/test';

import { resetTabSession, slideStage } from './support/deck';
import { EXTERNAL_ORIGIN, externalContentDeck } from './support/external-content-deck';
import { openOptionsDialog, optionsCategory } from './support/settings-dialog';

const BLUE_PNG = Buffer.from(
	'iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAIAAAD91JpzAAAAEklEQVR4nGNkaGBgYGBgYGAAAAN4AE0r1L6tAAAAAElFTkSuQmCC',
	'base64',
);
// The label also carries an info button, so match the accessible name by its prefix.
const TOGGLE_NAME = /^Allow external content \(remote images and media\)/u;
const LOCK_HINT = 'This setting is managed by your organization';

/** Answer the external origin offline and record every request it receives. */
async function recordExternalRequests(page: Page): Promise<string[]> {
	const urls: string[] = [];
	await page.route(`${EXTERNAL_ORIGIN}/**`, (route) => {
		const url = route.request().url();
		urls.push(url);
		return url.endsWith('.png')
			? route.fulfill({ contentType: 'image/png', body: BLUE_PNG })
			: route.fulfill({ status: 404, body: '' });
	});
	return urls;
}

async function openExternalDeck(page: Page, query = ''): Promise<void> {
	await resetTabSession(page);
	await page.goto(`/${query}`);
	await page.locator('#file-input').setInputFiles({
		name: 'external-content.pptx',
		mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
		buffer: await externalContentDeck(),
	});
	await expect(slideStage(page).locator('[data-element-id]').first()).toBeVisible();
}

async function trustToggle(page: Page): Promise<Locator> {
	const dialog = await openOptionsDialog(page, ['Options']);
	await optionsCategory(dialog, 'Trust Center').click();
	return page.locator('pptx-ui-options-dialog').getByRole('checkbox', { name: TOGGLE_NAME });
}

test.describe('allowExternalContent host switch', () => {
	test('unset keeps the Trust Center option, which allows linked content by default', async ({
		page,
	}) => {
		const requests = await recordExternalRequests(page);
		await openExternalDeck(page);
		await expect.poll(() => requests.length).toBeGreaterThan(0);
		const toggle = await trustToggle(page);
		await expect(toggle).toBeChecked();
		await expect(toggle).toBeEnabled();
	});

	test('false sends no request to the external origin and locks the toggle off', async ({
		page,
	}) => {
		const requests = await recordExternalRequests(page);
		await openExternalDeck(page, '?allowExternalContent=0');
		// Give any (unwanted) image or media fetch time to start.
		await page.waitForTimeout(1000);
		expect(requests).toEqual([]);
		await expect(page.locator(`[src^="${EXTERNAL_ORIGIN}"]`)).toHaveCount(0);
		const toggle = await trustToggle(page);
		await expect(toggle).not.toBeChecked();
		await expect(toggle).toBeDisabled();
		// The shared dialog gives a host-locked row the reason as its tooltip.
		await expect(
			page.locator('pptx-ui-options-dialog label[data-key$="allowExternalContent"]'),
		).toHaveAttribute('title', LOCK_HINT);
		expect(requests).toEqual([]);
	});

	test('true fetches linked content and locks the toggle on', async ({ page }) => {
		const requests = await recordExternalRequests(page);
		await openExternalDeck(page, '?allowExternalContent=1');
		await expect.poll(() => requests.length).toBeGreaterThan(0);
		const toggle = await trustToggle(page);
		await expect(toggle).toBeChecked();
		await expect(toggle).toBeDisabled();
	});
});
