/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec, `test`/`expect` come from @playwright/test */
/**
 * Linked external content follows Microsoft PowerPoint's default: opening a deck
 * downloads none of its linked pictures and media, and a security bar offers
 * Enable Content for that presentation, with or without the toolbar and in the
 * building blocks' host shell. Enabling only redraws: the deck, its unsaved edits
 * and undo history stay, and saves keep the pictures linked. The host `allowExternalContent` switch
 * (the demos' `?allowExternalContent=0|1`, see `demos/pptx/shared/demo-root-options.ts`)
 * blocks without the bar or downloads without asking. The deck links pictures
 * and a video to an external origin, answered here with `page.route`, so every
 * request it causes is recorded.
 */
import { expect, test } from '@playwright/test';
import type { Locator, Page } from '@playwright/test';
import JSZip from 'jszip';

import { requireBuildingBlocks } from './support/building-blocks';
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

/** The x of the first element on the first slide, through the demo's viewer handle. */
function firstElementX(page: Page): Promise<number> {
	return page.evaluate(() => {
		const viewer = (
			window as unknown as {
				__pptxViewer: { getSlides(): { elements: { x: number }[] }[] };
			}
		).__pptxViewer;
		return viewer.getSlides()[0]!.elements[0]!.x;
	});
}

/** Press the host save panel's Save and wait until the viewer reports it clean. */
async function hostSave(page: Page): Promise<number[]> {
	const log = page.getByTestId('host-save-log').locator('li');
	const done = await log.count();
	await page.getByTestId('host-save').click();
	await expect(log).toHaveCount(done + 1);
	await expect(log.last()).toHaveText(/^clean \d+$/u);
	return page.evaluate(() => {
		const snapshots = (window as unknown as { __hostSaveSnapshots: { content: Uint8Array }[] })
			.__hostSaveSnapshots;
		return [...snapshots.at(-1)!.content];
	});
}

async function slideXml(bytes: number[]): Promise<string> {
	return (await JSZip.loadAsync(Buffer.from(bytes))).file('ppt/slides/slide1.xml')!.async('string');
}

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

	test('keeps an edit saved before Enable Content, and the next save writes it', async ({
		page,
	}) => {
		const requests = await recordExternalRequests(page);
		await openExternalDeck(page, '?hostSave=1');
		const before = await firstElementX(page);
		await page.getByTestId('host-save-edit').click();
		await expect.poll(() => firstElementX(page)).toBe(before + 10);
		const saved = await slideXml(await hostSave(page));
		await expect(page.getByTestId('host-save-state')).toContainText('"dirty":false');

		await securityBar(page).getByRole('button', { name: 'Enable Content' }).click();
		await expect.poll(() => requestedPaths(requests)).toContain(LINKED_PICTURES.shapeFill);
		await expect(securityBar(page)).toBeHidden();
		// The edit is still there and still saved: nothing reloaded.
		expect(await firstElementX(page)).toBe(before + 10);
		await expect(page.getByTestId('host-save-state')).toContainText('"dirty":false');

		const next = await slideXml(await hostSave(page));
		const offsets = (xml: string) => xml.match(/<a:off [^>]*>/gu);
		expect(offsets(next)).toEqual(offsets(saved));
		// The linked pictures stay linked, not embedded.
		expect(next).toContain('r:link="rIdLinkedFill"');
	});

	test('keeps an unsaved edit unsaved and undoable through Enable Content', async ({ page }) => {
		const requests = await recordExternalRequests(page);
		await openExternalDeck(page, '?hostSave=1');
		const before = await firstElementX(page);
		await page.getByTestId('host-save-edit').click();
		await expect.poll(() => firstElementX(page)).toBe(before + 10);
		await securityBar(page).getByRole('button', { name: 'Enable Content' }).click();
		await expect.poll(() => requestedPaths(requests)).toContain(LINKED_PICTURES.shapeFill);
		expect(await firstElementX(page)).toBe(before + 10);
		await expect(page.getByTestId('host-save-state')).toContainText('"dirty":true');
		// Undo history survives: the edit made before Enable Content can be undone.
		await page.getByTestId('host-save-undo').click();
		await expect.poll(() => firstElementX(page)).toBe(before);
		// Undo never takes the linked content away again.
		await expect(securityBar(page)).toHaveCount(0);
		expect(await externalReferences(page)).not.toEqual([]);
	});

	test('shows the security bar with the toolbar off', async ({ page }) => {
		const requests = await recordExternalRequests(page);
		await openExternalDeck(page, '?showToolbar=0');
		const bar = securityBar(page);
		await expect(bar).toBeVisible();
		await expectNothingRequested(page, requests);
		await bar.getByRole('button', { name: 'Enable Content' }).click();
		await expect.poll(() => requestedPaths(requests)).toContain(LINKED_PICTURES.shapeFill);
		await expect(bar).toBeHidden();
	});

	test('the building blocks hand the security bar to the host shell', async ({ page }) => {
		requireBuildingBlocks();
		const requests = await recordExternalRequests(page);
		const deck = await externalContentDeck();
		// The blocks page opens the sample deck; answer it with the linked one.
		await page.route('**/sample-deck.pptx', (route) =>
			route.fulfill({ contentType: 'application/octet-stream', body: deck }),
		);
		await page.goto('/?blocks=1');
		await expect(page.locator('[data-building-blocks-panel]')).toBeVisible({ timeout: 30_000 });
		const bar = securityBar(page);
		await expect(bar).toHaveCount(1, { timeout: 30_000 });
		await expect(bar).toBeVisible();
		expect(requests).toEqual([]);
		await bar.getByRole('button', { name: 'Enable Content' }).click();
		await expect.poll(() => requestedPaths(requests)).toContain(LINKED_PICTURES.shapeFill);
		await expect(bar).toHaveCount(0);
	});

	test('enabling is not kept for the next open', async ({ page }) => {
		const requests = await recordExternalRequests(page);
		await openExternalDeck(page);
		await securityBar(page).getByRole('button', { name: 'Enable Content' }).click();
		await expect.poll(() => requestedPaths(requests)).toContain(LINKED_PICTURES.shapeFill);
		// Opening the presentation again asks again, as PowerPoint does. (The demos offer the file
		// input only on their start page; each binding's unit tests cover a second deck opened in
		// the same viewer.) Count from the new page on: until it replaces the enabled one, that
		// page's video can still ask for more of its file.
		await resetTabSession(page);
		await page.goto('/');
		requests.length = 0;
		await chooseExternalDeck(page);
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
