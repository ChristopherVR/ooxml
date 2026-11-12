/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec */
import { writeFile } from 'node:fs/promises';

import { expect, test } from '@playwright/test';
import type { Page } from '@playwright/test';

import { loadDeck, SAMPLE_DECK } from './support/deck';

/**
 * The host-persisted Save contract (`getSaveSnapshot` / `acknowledgeSave`), driven through the
 * demos' "host save" panel (`demos/pptx/shared/host-save-panel.ts`, shown by `?hostSave=1`).
 */

type SaveWindow = Window & {
	__pptxViewer: {
		isDirty(): boolean;
		canUndo(): boolean;
		getSlides(): { elements: { id: string; x: number }[] }[];
	};
	__hostSaveSnapshots?: { revision: number; content: Uint8Array }[];
};

const panel = (page: Page, id: string) => page.getByTestId(id);
const dirty = (page: Page) => page.evaluate(() => (window as SaveWindow).__pptxViewer.isDirty());
const firstX = (page: Page) =>
	page.evaluate(() => (window as SaveWindow).__pptxViewer.getSlides()[0]?.elements[0]?.x ?? NaN);

async function setDelay(page: Page, ms: number): Promise<void> {
	await panel(page, 'host-save-delay').fill(String(ms));
}

test.beforeEach(async ({ page }) => {
	await page.setViewportSize({ width: 1500, height: 1000 });
	await loadDeck(page, SAMPLE_DECK, '/?hostSave=1');
	await panel(page, 'host-save-panel').waitFor();
	// The deck opens clean; a host edit makes it dirty with a new revision.
	expect(await dirty(page)).toBe(false);
	await panel(page, 'host-save-edit').click();
	await expect.poll(() => dirty(page)).toBe(true);
});

test('acknowledging an unchanged snapshot leaves the deck clean', async ({ page }) => {
	await setDelay(page, 150);
	await panel(page, 'host-save').click();
	await expect(panel(page, 'host-save-result')).toHaveText(/^clean \d+$/u);
	await expect.poll(() => dirty(page)).toBe(false);
});

test('an edit during persistence stays dirty', async ({ page }) => {
	await setDelay(page, 800);
	await panel(page, 'host-save').click();
	await expect(panel(page, 'host-save-inflight')).toHaveText('1');
	await panel(page, 'host-save-edit').click();
	await expect(panel(page, 'host-save-result')).toHaveText(/^pending \d+$/u);
	expect(await dirty(page)).toBe(true);
	await expect(panel(page, 'host-save-inflight')).toHaveText('0');
});

test('a failed save leaves the deck dirty', async ({ page }) => {
	await setDelay(page, 100);
	await panel(page, 'host-save-fail').check();
	await panel(page, 'host-save').click();
	await expect(panel(page, 'host-save-result')).toHaveText(/^failed /u);
	expect(await dirty(page)).toBe(true);
});

test('an older revision acknowledged after a newer one is stale and changes nothing', async ({
	page,
}) => {
	// Save A takes its snapshot now but completes late.
	await setDelay(page, 1200);
	await panel(page, 'host-save').click();
	await expect(panel(page, 'host-save-inflight')).toHaveText('1');
	await panel(page, 'host-save-edit').click();
	// Save B snapshots the newer state and completes first.
	await setDelay(page, 0);
	await panel(page, 'host-save').click();
	await expect(panel(page, 'host-save-log').locator('li').first()).toHaveText(/^clean \d+$/u);
	await expect.poll(() => dirty(page)).toBe(false);
	await expect(panel(page, 'host-save-result')).toHaveText(/^stale \d+$/u);
	expect(await dirty(page)).toBe(false);
	// A revision the viewer never issued is ignored as well.
	await panel(page, 'host-save-ack-revision').fill('99999');
	await panel(page, 'host-save-ack').click();
	await expect(panel(page, 'host-save-result')).toHaveText('unknown 99999');
	expect(await dirty(page)).toBe(false);
});

test('undo still works after a clean acknowledgement', async ({ page }) => {
	const edited = await firstX(page);
	await setDelay(page, 50);
	await panel(page, 'host-save').click();
	await expect(panel(page, 'host-save-result')).toHaveText(/^clean \d+$/u);
	const canUndo = await page.evaluate(() => (window as SaveWindow).__pptxViewer.canUndo());
	expect(canUndo).toBe(true);
	await panel(page, 'host-save-undo').click();
	await expect.poll(() => firstX(page)).toBe(edited - 10);
	// Undoing is an edit like any other: the document is unsaved again.
	await expect.poll(() => dirty(page)).toBe(true);
});

test('the snapshot bytes reopen as a valid deck', async ({ page }, testInfo) => {
	const slideCount = await page.evaluate(
		() => (window as SaveWindow).__pptxViewer.getSlides().length,
	);
	await setDelay(page, 0);
	await panel(page, 'host-save').click();
	await expect(panel(page, 'host-save-result')).toHaveText(/^clean \d+$/u);
	const bytes = await page.evaluate(() =>
		Array.from((window as SaveWindow).__hostSaveSnapshots?.[0]?.content ?? []),
	);
	expect(bytes.length).toBeGreaterThan(1000);
	// A .pptx is a zip: it starts with "PK".
	expect(bytes.slice(0, 2)).toEqual([0x50, 0x4b]);
	const file = testInfo.outputPath('host-saved.pptx');
	await writeFile(file, Buffer.from(bytes));
	await loadDeck(page, file);
	expect(await page.evaluate(() => (window as SaveWindow).__pptxViewer.getSlides().length)).toBe(
		slideCount,
	);
});
