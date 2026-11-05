/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec */
/**
 * OneDrive under File > Open in every binding, against a test-only Graph stand-in: sign in from the
 * OneDrive place, open a deck from the listing, edit, and the viewer's own File > Save uploads to the
 * same item. The status bar reads "All saved" only once the upload has landed; a failed upload leaves
 * it unsaved; a file on another drive opens read-only.
 */
import { readFile } from 'node:fs/promises';

import { expect, test } from '@playwright/test';
import type { Page } from '@playwright/test';

import { loadDeck, openRibbonTab, thumbnail } from './support/deck';
import { ONEDRIVE_TEST_TOKEN, routeOneDriveGraph } from './support/onedrive-graph';

const deck = () => readFile(new URL('./fixtures/sample-deck.pptx', import.meta.url));
const unsaved = (page: Page) => page.getByText('Unsaved changes', { exact: true });

async function openBackstage(page: Page) {
	const toolbar = page.getByRole('toolbar', { name: 'Presentation toolbar' });
	await toolbar.getByRole('tab', { name: 'File', exact: true }).click();
	return page.getByRole('dialog', { name: 'File' });
}

/** File > Open > OneDrive > Sign in (demo dialog with the test token) > open `name`. */
async function openFromOneDrive(page: Page, name: string): Promise<void> {
	// The viewer (and its File tab) needs a document first.
	await loadDeck(page);
	const backstage = await openBackstage(page);
	await backstage.getByRole('button', { name: 'Open', exact: true }).click();
	await backstage.locator('[data-pptx-open-place="cloud:onedrive"]').click();
	const place = backstage.locator('[data-pptx-cloud-place="onedrive"]');
	await place.getByRole('button', { name: 'Sign in', exact: true }).click();
	const dialog = page.getByRole('dialog', { name: 'Connect OneDrive (demo)' });
	await dialog.getByLabel('Access token').fill(ONEDRIVE_TEST_TOKEN);
	await dialog.getByRole('button', { name: 'Continue', exact: true }).click();
	await place.getByRole('option', { name: new RegExp(name, 'iu') }).click();
	// The title bar names the opened file once it has replaced the deck that was showing.
	await expect(page.getByText(new RegExp(`${name}\.pptx`, 'iu')).first()).toBeVisible({
		timeout: 30_000,
	});
	await expect(page.locator('[data-pptx-viewport] [data-pptx-element="true"]').first()).toBeVisible(
		{ timeout: 30_000 },
	);
}

/** The edit `addSlide` makes, tried on a deck that may refuse it (a read-only ribbon, a no-op). */
async function tryAddSlide(page: Page): Promise<void> {
	try {
		await page
			.getByRole('toolbar', { name: 'Presentation toolbar' })
			.getByRole('tab', { name: 'Home', exact: true })
			.click({ timeout: 3000 });
		await page
			.getByRole('button', { name: /new slide/iu })
			.first()
			.click({ timeout: 3000 });
	} catch {
		// A read-only deck may not offer the command at all.
	}
	await page.waitForTimeout(1000);
}

async function addSlide(page: Page): Promise<void> {
	await openRibbonTab(page, 'Home');
	await page
		.getByRole('button', { name: /new slide/iu })
		.first()
		.click();
	await expect(unsaved(page).first()).toBeVisible();
}

async function fileSave(page: Page): Promise<void> {
	const backstage = await openBackstage(page);
	await backstage.getByRole('button', { name: 'Save', exact: true }).click();
}

test.describe('OneDrive under File > Open', () => {
	test('File > Save uploads to the same item and the status bar reads All saved', async ({
		page,
	}) => {
		const cloud = await routeOneDriveGraph(page, {
			mine: { name: 'Quarterly review.pptx', bytes: await deck() },
		});
		await openFromOneDrive(page, 'Quarterly review');
		await addSlide(page);
		await fileSave(page);
		// "All saved", or the autosave line once its recovery copy landed (it takes precedence).
		await expect(page.getByText(/^(All saved|Saved (just now|\d+ min.*))$/u).first()).toBeVisible({
			timeout: 30_000,
		});
		expect(cloud.writes).toStrictEqual(['mine']);
		await expect(unsaved(page)).toHaveCount(0);
		// Saving writes the saved slide identity back; moving off the new slide is no edit.
		await thumbnail(page, 1).click();
		await page.waitForTimeout(500);
		await expect(unsaved(page)).toHaveCount(0);
	});

	test('a failed upload keeps the deck unsaved', async ({ page }) => {
		const cloud = await routeOneDriveGraph(page, {
			mine: { name: 'Quarterly review.pptx', bytes: await deck() },
		});
		await openFromOneDrive(page, 'Quarterly review');
		await addSlide(page);
		cloud.failUploads(503);
		await fileSave(page);
		await page.waitForTimeout(1500);
		expect(cloud.writes).toStrictEqual([]);
		await expect(unsaved(page).first()).toBeVisible();
	});

	test('a file on another drive opens read-only', async ({ page }) => {
		await routeOneDriveGraph(page, {
			shared: { name: 'Shared with me.pptx', bytes: await deck(), driveId: 'colleague-drive' },
		});
		await openFromOneDrive(page, 'Shared with me');
		// The edit that marks an own-drive deck unsaved leaves this one untouched.
		await tryAddSlide(page);
		await expect(unsaved(page)).toHaveCount(0);
	});
});
