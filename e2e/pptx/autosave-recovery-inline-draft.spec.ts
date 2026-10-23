/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec, `test`/`expect` come from @playwright/test */
/**
 * Typing that was never committed is recoverable, as PowerPoint's AutoRecover
 * keeps it: type into a text box, leave the editor open, let one autosave
 * interval pass, lose the tab, reopen the deck, and Restore brings the typed
 * text back.
 *
 * The inline editor holds typed text until it is committed (blur, click away),
 * so the slides themselves never see it. The shared autosave engine counts the
 * draft as an edit (`reportInlineTextDraft`) and writes it into the snapshot.
 *
 * Framework-neutral: the shared `[data-inline-editor]` hook, the shared slide
 * contract, the shared IndexedDB store and the recovery dialog by role.
 *
 * Run: bunx playwright test autosave-recovery-inline-draft
 */
import { expect, test } from '@playwright/test';
import type { Page } from '@playwright/test';

import { loadDeck, SAMPLE_DECK, slideElements } from './support/deck';

test.use({ viewport: { width: 1600, height: 950 } });
// The demos autosave every 2s; the budget allows for a cold dev server.
test.describe.configure({ timeout: 180_000 });

const DRAFT = 'Uncommitted draft 7319';

/** Whether the shared store holds a non-empty snapshot. */
async function hasSnapshot(page: Page): Promise<boolean> {
	return page.evaluate(
		() =>
			new Promise<boolean>((resolve) => {
				const request = indexedDB.open('pptx-viewer-autosave');
				request.onerror = () => resolve(false);
				request.onsuccess = () => {
					const db = request.result;
					if (!db.objectStoreNames.contains('recoveryVersions')) {
						db.close();
						resolve(false);
						return;
					}
					const all = db
						.transaction('recoveryVersions', 'readonly')
						.objectStore('recoveryVersions')
						.getAll();
					all.onerror = () => {
						db.close();
						resolve(false);
					};
					all.onsuccess = () => {
						db.close();
						resolve(
							(all.result as Array<{ size?: number }>).some((record) => (record.size ?? 0) > 0),
						);
					};
				};
			}),
	);
}

test('uncommitted inline typing is offered back after the tab is lost', async ({ page }) => {
	await loadDeck(page, SAMPLE_DECK);

	// Open a text box's inline editor and type, without committing.
	await slideElements(page).filter({ hasText: /\w{3}/u }).first().dblclick();
	const editor = page.locator('[data-inline-editor]').first();
	await expect(editor).toBeVisible();
	await editor.press('ControlOrMeta+A');
	await page.keyboard.type(DRAFT);
	await expect(editor).toBeVisible();

	// One interval later the draft is in a snapshot, the editor still open.
	await expect.poll(() => hasSnapshot(page), { timeout: 60_000 }).toBe(true);
	await expect(editor).toBeVisible();

	// The tab is lost: the same deck reopened with no session, the snapshot on disk.
	await page.evaluate(() => sessionStorage.clear());
	await page.goto('/');
	await page.locator('#file-input').setInputFiles(SAMPLE_DECK);
	await page.locator('[aria-roledescription="slide"]').first().waitFor({ timeout: 60_000 });

	const dialog = page.getByRole('dialog', { name: /recover unsaved changes/iu });
	await expect(dialog).toBeVisible({ timeout: 30_000 });
	await dialog.getByRole('button', { name: /^restore$/iu }).click();
	await expect(dialog).toHaveCount(0);
	await expect(slideElements(page).filter({ hasText: DRAFT })).toHaveCount(1, { timeout: 60_000 });
});
