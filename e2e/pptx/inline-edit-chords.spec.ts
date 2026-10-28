/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec, `test`/`expect` come from @playwright/test */
/**
 * The chords an inline text edit owns, in every binding.
 *
 * Ctrl/Cmd+Enter leaves the edit and keeps the text. React and Angular did this
 * by hand; Vanilla, Svelte and Vue let the chord through to the text box, so
 * the user stayed in it.
 *
 * Ctrl/Cmd+D is cancelled mid-edit. The editor keymap stands down while text is
 * under edit, so nothing cancelled the browser's own meaning of the chord,
 * "bookmark this page", and the bookmark dialog opened over the deck.
 *
 * Both are the shared `mapInlineEditorChord` (`render/editor-keymap-chords.ts`).
 *
 * Run: bunx playwright test inline-edit-chords
 */
import { expect, test } from '@playwright/test';
import type { Page } from '@playwright/test';

import { elementWithText, fixture, loadDeck, selectElement, slideStage } from './support/deck';

test.use({ viewport: { width: 1440, height: 900 } });

/** A one-slide deck holding exactly two shapes, "SOURCE" and "TARGET". */
const TWO_SHAPES = fixture('format-painter.pptx');

async function editSource(page: Page): Promise<void> {
	await loadDeck(page, TWO_SHAPES);
	const source = elementWithText(page, 'SOURCE');
	await selectElement(page, source);
	await source.dblclick();
	await expect(page.locator('[data-inline-editor]')).toBeVisible();
}

test.describe('inline edit chords', () => {
	test('Ctrl/Cmd+D is cancelled while editing text and the edit stays open', async ({ page }) => {
		await editSource(page);
		// Read after every listener has run: a binding may cancel it late in the dispatch.
		await page.evaluate(() => {
			document.addEventListener(
				'keydown',
				(event) => {
					if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'd') {
						setTimeout(() => {
							document.body.dataset.bookmarkChordPrevented = String(event.defaultPrevented);
						});
					}
				},
				true,
			);
		});
		await page.keyboard.press('ControlOrMeta+d');
		await expect
			.poll(() => page.evaluate(() => document.body.dataset.bookmarkChordPrevented))
			.toBe('true');
		await expect(page.locator('[data-inline-editor]')).toBeVisible();
	});

	test('Ctrl/Cmd+Enter ends the edit and keeps the typed text', async ({ page }) => {
		await editSource(page);
		await page.keyboard.press('End');
		await page.keyboard.type(' DONE');
		await page.keyboard.press('ControlOrMeta+Enter');
		await expect(page.locator('[data-inline-editor]')).toBeHidden();
		await expect(slideStage(page)).toContainText('SOURCE DONE');
	});

	test('Escape ends the edit and keeps the typed text through the next edit', async ({ page }) => {
		await editSource(page);
		await page.keyboard.press('End');
		await page.keyboard.type(' FIRST');
		await page.keyboard.press('Escape');
		await expect(page.locator('[data-inline-editor]')).toBeHidden();
		await expect(slideStage(page)).toContainText('SOURCE FIRST');
		await elementWithText(page, 'SOURCE FIRST').dblclick();
		await expect(page.locator('[data-inline-editor]')).toBeVisible();
		await page.keyboard.press('End');
		await page.keyboard.type(' SECOND');
		await page.keyboard.press('Escape');
		await expect(slideStage(page)).toContainText('SOURCE FIRST SECOND');
	});
});
