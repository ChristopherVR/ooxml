/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec */
/**
 * The OneDrive demo (`?onedrive=1`) in every binding, against a test-only Graph stand-in: open a
 * sharing link, edit, and the viewer's own File > Save uploads to the same item. The document reads
 * clean only once the upload has landed; a failed upload leaves it unsaved.
 */
import { readFile } from 'node:fs/promises';

import { expect, test } from '@playwright/test';
import type { Page } from '@playwright/test';

import { openRibbonTab, thumbnail } from './support/deck';
import { ONEDRIVE_TEST_TOKEN, routeOneDriveGraph, sharingLink } from './support/onedrive-graph';

const deck = () => readFile(new URL('./fixtures/sample-deck.pptx', import.meta.url));
const status = (page: Page) => page.locator('.od-bar').getByRole('status');

async function openFromOneDrive(page: Page, id: string): Promise<void> {
	await page.goto('/?onedrive=1');
	const bar = page.locator('.od-bar');
	await bar.getByLabel('How to sign in').selectOption('token');
	await bar.getByLabel('Access token').fill(ONEDRIVE_TEST_TOKEN);
	await bar.getByRole('button', { name: 'Use token', exact: true }).click();
	await expect(bar.getByText('Signed in: Graph Explorer token')).toBeVisible();
	await bar.getByLabel('OneDrive sharing link').fill(sharingLink(id));
	await bar.getByRole('button', { name: 'Open', exact: true }).click();
	await expect(status(page)).toContainText('Opened', { timeout: 30_000 });
	await expect(page.locator('[data-pptx-viewport] [data-pptx-element="true"]').first()).toBeVisible(
		{ timeout: 30_000 },
	);
}

async function addSlide(page: Page): Promise<void> {
	await openRibbonTab(page, 'Home');
	await page
		.getByRole('button', { name: /new slide/iu })
		.first()
		.click();
	await expect(page.getByText('Unsaved changes', { exact: true }).first()).toBeVisible();
}

async function fileSave(page: Page): Promise<void> {
	const toolbar = page.getByRole('toolbar', { name: 'Presentation toolbar' });
	await toolbar.getByRole('tab', { name: 'File', exact: true }).click();
	const backstage = page.getByRole('dialog', { name: 'File' });
	await backstage.getByRole('button', { name: 'Save', exact: true }).click();
}

test.describe('OneDrive demo', () => {
	test('File > Save uploads to the same item and the deck reads saved', async ({ page }) => {
		const cloud = await routeOneDriveGraph(page, {
			mine: { name: 'Quarterly review.pptx', bytes: await deck() },
		});
		await openFromOneDrive(page, 'mine');
		await addSlide(page);
		await fileSave(page);
		await expect(status(page)).toContainText('Saved to OneDrive', { timeout: 30_000 });
		expect(cloud.writes).toStrictEqual(['mine']);
		await expect(page.getByText('Unsaved changes', { exact: true })).toHaveCount(0);
		// Saving writes the saved slide identity back; moving off the new slide is no edit.
		await thumbnail(page, 1).click();
		await page.waitForTimeout(500);
		await expect(page.getByText('Unsaved changes', { exact: true })).toHaveCount(0);
	});

	test('a failed upload keeps the deck unsaved', async ({ page }) => {
		const cloud = await routeOneDriveGraph(page, {
			mine: { name: 'Quarterly review.pptx', bytes: await deck() },
		});
		await openFromOneDrive(page, 'mine');
		await addSlide(page);
		cloud.failUploads(503);
		await fileSave(page);
		await expect(status(page)).toContainText('Not saved', { timeout: 30_000 });
		expect(cloud.writes).toStrictEqual([]);
		await expect(page.getByText('Unsaved changes', { exact: true }).first()).toBeVisible();
	});

	test('a file on another drive opens read-only', async ({ page }) => {
		await routeOneDriveGraph(page, {
			shared: { name: 'Shared with me.pptx', bytes: await deck(), driveId: 'colleague-drive' },
		});
		await openFromOneDrive(page, 'shared');
		await expect(status(page)).toContainText('read-only');
	});
});
