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
import {
	EXTERNAL_CONTENT_DECK_ELEMENTS,
	EXTERNAL_ORIGIN,
	LINKED_PICTURES,
	externalContentDeck,
} from './support/external-content-deck';
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
	// Every element of the slide is on the stage, so each has had its chance to reference
	// (and fetch) its linked content.
	await expect
		.poll(() =>
			slideStage(page)
				.locator('[data-element-id]')
				.evaluateAll(
					(nodes) => new Set(nodes.map((node) => node.getAttribute('data-element-id'))).size,
				),
		)
		.toBeGreaterThanOrEqual(EXTERNAL_CONTENT_DECK_ELEMENTS);
}

/** Every `src`, `poster`, `href`, `srcset` or inline style in the page, shadow roots included, that names the linked origin. */
function externalReferences(page: Page): Promise<string[]> {
	return page.evaluate((origin) => {
		const found: string[] = [];
		const visit = (root: Document | ShadowRoot) => {
			for (const element of root.querySelectorAll('*')) {
				for (const name of ['src', 'poster', 'href', 'xlink:href', 'srcset', 'style']) {
					const value = element.getAttribute(name);
					if (value?.includes(origin)) {
						found.push(`${element.tagName.toLowerCase()}[${name}]`);
					}
				}
				if (element.shadowRoot) {
					visit(element.shadowRoot);
				}
			}
		};
		visit(document);
		return found;
	}, EXTERNAL_ORIGIN);
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
		// A fetch can only start from a DOM reference to the origin, and with every element
		// rendered there is none; the network then settles with nothing sent.
		expect(await externalReferences(page)).toEqual([]);
		await page.waitForLoadState('networkidle');
		expect(requests).toEqual([]);
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
		// Each kind of linked picture is fetched: the picture and the 3D poster (both from the
		// relinked media part), the shape fill and the picture bullet.
		const requested = () => requests.map((url) => new URL(url).pathname);
		await expect.poll(requested).toContain(LINKED_PICTURES.shapeFill);
		await expect.poll(requested).toContain(LINKED_PICTURES.pictureBullet);
		const toggle = await trustToggle(page);
		await expect(toggle).toBeChecked();
		await expect(toggle).toBeDisabled();
	});
});
