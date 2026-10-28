/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec, `test`/`expect` come from @playwright/test */
/**
 * With a bare caret in a list paragraph, switching between Bullets and
 * Numbering converts the whole list around the caret (Word / Google Docs),
 * not only the caret's paragraph. Every binding routes its list buttons
 * through the shared `applyListStyleUpdate`; this pins that each one passes
 * the live caret through and renders the converted list.
 *
 * Fixture: `nested-list-continuation.pptx` holds one numbered list, two
 * top-level items around two nested level-1 items, then a third top-level item.
 */

import { expect, test } from '@playwright/test';
import type { Page } from '@playwright/test';

import { fixture, loadDeck, slideElements } from './support/deck';

test.describe.configure({ timeout: 120_000 });
// Home groups collapse into popups below about 1900px; this spec drives the expanded controls.
test.use({ viewport: { width: 1920, height: 1000 } });

const NESTED_LIST = fixture('nested-list-continuation.pptx');
const PARAGRAPHS = 5;

async function openRibbonTab(page: Page, name: string): Promise<void> {
	await page
		.getByRole('toolbar', { name: 'Presentation toolbar' })
		.getByRole('tab', { name, exact: true })
		.click();
}

/** Put a collapsed caret inside `text` in the open inline editor. */
async function placeCaretIn(page: Page, text: string): Promise<void> {
	await page.waitForSelector('[contenteditable="true"]');
	const placed = await page.evaluate((needle) => {
		const root = document.querySelector('[contenteditable="true"]');
		if (!root) {
			return 'no editor';
		}
		const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
		for (let node = walker.nextNode(); node; node = walker.nextNode()) {
			const at = node.textContent?.indexOf(needle) ?? -1;
			if (at >= 0) {
				const range = document.createRange();
				range.setStart(node, at + 2);
				range.collapse(true);
				const selection = window.getSelection();
				selection?.removeAllRanges();
				selection?.addRange(range);
				document.dispatchEvent(new Event('selectionchange'));
				return true;
			}
		}
		return `editor text: ${JSON.stringify(root.textContent)}`;
	}, text);
	expect(placed, `found "${text}" in the inline editor`).toBe(true);
}

const bulletCount = (text: string | null) => (text?.match(/•/gu) ?? []).length;

test('a caret in a numbered list switches the whole list to bullets', async ({ page }) => {
	await loadDeck(page, NESTED_LIST);
	await openRibbonTab(page, 'Home');
	const target = slideElements(page).filter({ hasText: 'Roman first' }).first();
	const elsewhere = slideElements(page).filter({ hasText: 'Native event probe' }).first();
	expect(bulletCount(await target.textContent())).toBe(0);

	await target.dblclick();
	// The first text node holding "Roman" is the first item's (runs may split words).
	await placeCaretIn(page, 'Roman');
	await page
		.getByRole('button', { name: /^Bullet List$/iu })
		.first()
		.click();
	await elsewhere.click();

	await expect
		.poll(async () => bulletCount(await target.textContent()), {
			message: 'every paragraph of the list carries a bullet',
		})
		.toBe(PARAGRAPHS);
	await expect(target).toContainText('Nested middle');
	await expect(target).toContainText('Roman last');
});
