/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec, `test`/`expect` come from @playwright/test */
/**
 * `@`-mentions in a MODERN (`p188`) threaded comment, across every binding.
 *
 * Before this, mentions had no model at all: `PptxComment` carried no mentions
 * field and no binding had any mention UI, so an `@Bob Grant` in a real deck
 * rendered as inert plain text indistinguishable from the rest of the body.
 * The offsets were also never re-based when the body was edited, so a surviving
 * `p188:mentionLst` could end up pointing at the wrong characters.
 *
 * The fixture is a PowerPoint-authored modern comment (COM), with only the
 * second `p188:author` and the `p188:mentionLst` span injected afterwards,
 * because a mention needs a real M365/AAD identity and cannot be authored
 * through COM. PowerPoint reopens it cleanly.
 *
 * The neutral contract every binding stamps is `data-pptx-comment-mention`
 * (`COMMENT_MENTION_ATTRIBUTE` in `pptx-viewer-shared`), carrying the
 * mentioned person's `p188:author` id.
 */
import { resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

import { expect, test } from '@playwright/test';
import type { Page } from '@playwright/test';
import { PptxHandler } from 'pptx-viewer-core';

import { savePptxViaBackstage } from './save-pptx';
import { COMPOSE_BOX_SELECTOR, commentTextVisible, openCommentsThread } from './support/comments';
import { resetTabSession } from './support/deck';
import { downloadBytes } from './support/exports';

test.use({ viewport: { width: 1440, height: 900 } });

const fixturePath = resolve(
	fileURLToPath(new URL('./fixtures/comment-mentions.pptx', import.meta.url)),
);

/** The whole comment body, and the span the `p188:mention` covers. */
const COMMENT_TEXT = 'Please check this with @Bob Grant before Friday';
const MENTION_TEXT = '@Bob Grant';
const MENTION_PERSON_ID = '{7C9E4A21-5F30-4C88-9B1E-2D6A0F4B7E15}';

async function openDeck(page: Page): Promise<void> {
	await resetTabSession(page);
	await page.goto('/');
	await page.locator('#file-input').setInputFiles(fixturePath);
	await page.locator('[data-element-id]').first().waitFor();
}

test.describe('modern comment mentions', () => {
	test('renders the mention span distinctly, carrying the author id', async ({ page }) => {
		await openDeck(page);
		await openCommentsThread(page, COMMENT_TEXT);

		const mention = page.locator('[data-pptx-comment-mention]').first();
		await expect(mention).toBeVisible();
		await expect(mention).toHaveText(MENTION_TEXT);
		await expect(mention).toHaveAttribute('data-pptx-comment-mention', MENTION_PERSON_ID);

		// The rest of the body must stay outside the mention span: a segmenter
		// that swallowed the whole string would still satisfy a text assertion
		// on the mention alone.
		const inside = await mention.evaluate((el) => el.textContent ?? '');
		expect(inside).not.toContain('Please check');
		expect(inside).not.toContain('before Friday');
	});

	test('shows the full comment body around the mention', async ({ page }) => {
		await openDeck(page);
		await openCommentsThread(page, COMMENT_TEXT);

		// Read through the segment spans rather than asserting on a particular
		// DOM shape: the five bindings nest the mention differently (a <p>, a
		// <span>, a component host), so a parent-relative locator tests the
		// binding's markup instead of the behaviour. What must hold everywhere
		// is that segmenting the body did not LOSE any of it.
		await expect.poll(() => commentTextVisible(page, COMMENT_TEXT)).toBe(true);
	});

	test('a new comment keeps its mention on the name through the trim and the save', async ({
		page,
	}) => {
		await openDeck(page);
		await openCommentsThread(page, COMMENT_TEXT);
		// Leading spaces are trimmed off when the comment is stored; the mention must move with them.
		const compose = page.locator(COMPOSE_BOX_SELECTOR).first();
		await compose.click();
		await compose.pressSequentially('  Ask @bo');
		await page.locator('[data-testid="pptx-comment-mention-option"]:visible').first().click();
		await expect(compose).toHaveValue(`  Ask ${MENTION_TEXT} `);
		await compose.pressSequentially('now');
		await page.locator('button:visible:text-is("Add Comment")').last().click();

		const added = `Ask ${MENTION_TEXT} now`;
		await expect.poll(() => commentTextVisible(page, added)).toBe(true);
		const mentions = page.locator('[data-pptx-comment-mention]');
		await expect(mentions).toHaveCount(2);
		await expect(mentions).toHaveText([MENTION_TEXT, MENTION_TEXT]);

		const saved = await new PptxHandler().load(
			(await downloadBytes(await savePptxViaBackstage(page))).slice().buffer,
		);
		const comment = saved.slides
			.flatMap((slide) => slide.comments ?? [])
			.find((entry) => entry.text === added);
		expect(
			comment?.mentions?.map((m) => comment.text.slice(m.startIndex, m.startIndex + m.length)),
		).toStrictEqual([MENTION_TEXT]);
	});
});
