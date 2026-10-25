/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec, `test`/`expect` come from @playwright/test */
/**
 * Keys typed into a field stay in the field, in every binding.
 *
 * Space, the arrows, Home, End and the Page keys also page the slides, so the viewer root listens
 * for them. Svelte's root handler took them from every descendant, a comment box or the notes
 * editor included, and `preventDefault()`ed them: a comment could not hold a space, and ArrowDown
 * in one changed the slide. Vanilla skipped inputs and textareas by tag, which missed the rich
 * notes editor (an editable region). Every other spec fills these fields with `fill()`, which
 * sends no key events, so nothing noticed.
 */
import { resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

import { expect, test } from '@playwright/test';
import type { Locator, Page } from '@playwright/test';

import { COMPOSE_BOX_SELECTOR, openCommentsThread } from './support/comments';
import { loadDeck, resetTabSession, SAMPLE_DECK, slideStage } from './support/deck';
import { slidePosition } from './support/keyboard';

test.use({ viewport: { width: 1440, height: 900 } });

/** One slide holding one comment, so every binding's comment UI has a thread to open. */
const COMMENTED_DECK = resolve(
	fileURLToPath(new URL('./fixtures/comment-mentions.pptx', import.meta.url)),
);
const COMMENT_TEXT = 'Please check this with @Bob Grant before Friday';

/** The keys a viewer root also uses for paging, pressed inside a field. */
const PAGING_KEYS = ['ArrowDown', 'ArrowUp', 'PageDown', 'PageUp', 'End'];

async function openNotesEditor(page: Page): Promise<Locator> {
	await page
		.getByRole('button', { name: /^toggle notes$/iu })
		.first()
		.focus();
	await page.keyboard.press('Enter');
	const editor = page.locator('#slide-notes-content [contenteditable="true"]').first();
	await editor.click();
	// Start from an empty editor: the slide may already have notes.
	await editor.press('ControlOrMeta+A');
	await editor.press('Backspace');
	return editor;
}

test.describe('keys typed into a text field', () => {
	test('a comment box keeps its spaces and its caret keys', async ({ page }) => {
		await resetTabSession(page);
		await page.goto('/');
		await page.locator('#file-input').setInputFiles(COMMENTED_DECK);
		await page.locator('[data-element-id]').first().waitFor();
		await openCommentsThread(page, COMMENT_TEXT);
		const compose = page.locator(COMPOSE_BOX_SELECTOR).first();
		await compose.click();

		await compose.pressSequentially('one two');
		await expect(compose).toHaveValue('one two');
		for (const key of PAGING_KEYS) {
			await compose.press(key);
		}
		// Home must have moved the caret to the start, or the X lands at the end.
		await compose.press('Home');
		await compose.pressSequentially('X ');
		await expect(compose).toHaveValue('X one two');
	});

	test('the notes editor keeps its spaces and does not change slide', async ({ page }) => {
		await loadDeck(page, SAMPLE_DECK);
		await slideStage(page).waitFor();
		expect(await slidePosition(page)).toBe('Slide 1 of 7');
		const editor = await openNotesEditor(page);

		await editor.pressSequentially('one two');
		for (const key of PAGING_KEYS) {
			await editor.press(key);
		}
		await editor.press('Home');
		await editor.pressSequentially('X ');
		await expect(editor).toHaveText('X one two');
		expect(await slidePosition(page), 'no key typed into the notes paged the slides').toBe(
			'Slide 1 of 7',
		);
	});
});
