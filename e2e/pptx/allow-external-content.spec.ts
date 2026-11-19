/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec, `test`/`expect` come from @playwright/test */
/**
 * Linked external content follows Microsoft PowerPoint's default: opening a deck
 * downloads none of its linked pictures and media, and a security bar offers
 * Enable Content for that presentation. The host `allowExternalContent` switch
 * (the demos' `?allowExternalContent=0|1`, see `demos/pptx/shared/demo-root-options.ts`)
 * blocks without the bar or downloads without asking. The deck links pictures
 * and a video to an external origin, answered here with `page.route`, so every
 * request it causes is recorded.
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

/** Choose the deck in the demo's file input and wait until every element of its slide is on the stage. */
async function chooseExternalDeck(page: Page): Promise<void> {
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

async function openExternalDeck(page: Page, query = ''): Promise<void> {
	await resetTabSession(page);
	await page.goto(`/${query}`);
	await chooseExternalDeck(page);
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

/** No reference to the origin is on the page and the network settles with nothing sent. */
async function expectNothingRequested(page: Page, requests: string[]): Promise<void> {
	// A fetch can only start from a DOM reference to the origin; with every element rendered
	// there is none.
	expect(await externalReferences(page)).toEqual([]);
	await page.waitForLoadState('networkidle');
	expect(requests).toEqual([]);
}

function securityBar(page: Page): Locator {
	return page.locator('pptx-ui-external-content-bar');
}

async function trustToggle(page: Page): Promise<Locator> {
	const dialog = await openOptionsDialog(page, ['Options']);
	await optionsCategory(dialog, 'Trust Center').click();
	return page.locator('pptx-ui-options-dialog').getByRole('checkbox', { name: TOGGLE_NAME });
}

const requestedPaths = (requests: string[]) => requests.map((url) => new URL(url).pathname);

test.describe('linked external content', () => {
	test('is blocked on open, with a security bar, until Enable Content', async ({ page }) => {
		const requests = await recordExternalRequests(page);
		await openExternalDeck(page);
		const bar = securityBar(page);
		await expect(bar).toBeVisible();
		await expect(bar).toContainText('SECURITY WARNING');
		await expect(bar).toContainText(
			'References to external pictures and other media have been blocked.',
		);
		await expectNothingRequested(page, requests);

		await bar.getByRole('button', { name: 'Enable Content' }).click();
		// Each kind of linked picture is fetched once enabled: the shape fill and the picture bullet.
		await expect.poll(() => requestedPaths(requests)).toContain(LINKED_PICTURES.shapeFill);
		await expect.poll(() => requestedPaths(requests)).toContain(LINKED_PICTURES.pictureBullet);
		await expect(bar).toBeHidden();
	});

	test('enabling is not kept for the next open', async ({ page }) => {
		const requests = await recordExternalRequests(page);
		await openExternalDeck(page);
		await securityBar(page).getByRole('button', { name: 'Enable Content' }).click();
		await expect.poll(() => requestedPaths(requests)).toContain(LINKED_PICTURES.shapeFill);
		requests.length = 0;
		// Opening the presentation again asks again, as PowerPoint does. (The demos offer the file
		// input only on their start page; each binding's unit tests cover a second deck opened in
		// the same viewer.)
		await openExternalDeck(page);
		await expect(securityBar(page)).toBeVisible();
		await expectNothingRequested(page, requests);
	});

	test('closing the bar keeps the content blocked', async ({ page }) => {
		const requests = await recordExternalRequests(page);
		await openExternalDeck(page);
		await securityBar(page).getByRole('button', { name: 'Close' }).click();
		await expect(securityBar(page)).toBeHidden();
		await expectNothingRequested(page, requests);
	});

	test('the Trust Center option is off by default and left to the user', async ({ page }) => {
		await recordExternalRequests(page);
		await openExternalDeck(page);
		const toggle = await trustToggle(page);
		await expect(toggle).not.toBeChecked();
		await expect(toggle).toBeEnabled();
	});

	test('host false sends no request, shows no bar and locks the toggle off', async ({ page }) => {
		const requests = await recordExternalRequests(page);
		await openExternalDeck(page, '?allowExternalContent=0');
		await expectNothingRequested(page, requests);
		await expect(securityBar(page)).toHaveCount(0);
		const toggle = await trustToggle(page);
		await expect(toggle).not.toBeChecked();
		await expect(toggle).toBeDisabled();
		// The shared dialog gives a host-locked row the reason as its tooltip.
		await expect(
			page.locator('pptx-ui-options-dialog label[data-key$="allowExternalContent"]'),
		).toHaveAttribute('title', LOCK_HINT);
		expect(requests).toEqual([]);
	});

	test('host true fetches linked content without a bar and locks the toggle on', async ({
		page,
	}) => {
		const requests = await recordExternalRequests(page);
		await openExternalDeck(page, '?allowExternalContent=1');
		await expect.poll(() => requestedPaths(requests)).toContain(LINKED_PICTURES.shapeFill);
		await expect.poll(() => requestedPaths(requests)).toContain(LINKED_PICTURES.pictureBullet);
		await expect(securityBar(page)).toHaveCount(0);
		const toggle = await trustToggle(page);
		await expect(toggle).toBeChecked();
		await expect(toggle).toBeDisabled();
	});
});
