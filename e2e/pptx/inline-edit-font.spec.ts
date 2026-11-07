/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec, `test`/`expect` come from @playwright/test */
/**
 * Double-clicking a text box keeps the face view mode painted it in, in every
 * binding.
 *
 * Vanilla and Svelte gave the editor surface the body's RAW `fontFamily`
 * (`"ADLaM Display"`), overriding the block style's substitution chain
 * (`"ADLaM Display", ..., sans-serif`). When the deck font was not painted (a
 * webfont still downloading or blocked, offline, or the download option off)
 * view mode fell back to the chain's stand-in and the editor to the browser's
 * default serif, so the text visibly changed font on double-click. Svelte's
 * plain editor also dropped the alignment and colour, and Angular edited in an
 * unstyled <textarea> on white paper.
 *
 * The Google Fonts endpoints are blocked so the deck font is never downloaded
 * and the result does not depend on the network. The check compares computed
 * styles, so it holds whether or not this machine has the font installed.
 *
 * Run: bunx playwright test inline-edit-font
 */
import { expect, test } from '@playwright/test';
import type { Page } from '@playwright/test';

import { elementWithText, fixture, loadDeck, selectElement } from './support/deck';

test.use({ viewport: { width: 1440, height: 900 } });

/** Every run names `ADLaM Display`, a family served by Google Fonts. */
const WEBFONT_DECK = fixture('adlam-webfont.pptx');
const PROBE_TEXT = 'Box A';
/** View mode splits a run into one span per word, so the style is read from the first word. */
const PROBE_WORD = 'Box';

interface PaintedText {
	fontFamily: string;
	fontSize: string;
	fontWeight: string;
	color: string;
	textAlign: string;
}

/** The computed style of the innermost element holding `needle` under `root`. */
async function paintedText(page: Page, rootSelector: string, needle: string): Promise<PaintedText> {
	return page.evaluate(
		([selector, text]) => {
			const roots = [...document.querySelectorAll<HTMLElement>(selector)].filter((root) =>
				root.textContent?.includes(text),
			);
			// The canvas copy is the widest; thumbnails hold the same element ids.
			roots.sort((a, b) => b.getBoundingClientRect().width - a.getBoundingClientRect().width);
			const root = roots[0];
			if (!root) {
				throw new Error(`no ${selector} holds "${text}"`);
			}
			const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
			let holder: HTMLElement | null = null;
			for (let node = walker.nextNode(); node; node = walker.nextNode()) {
				if (node.textContent?.includes(text)) {
					holder = node.parentElement;
					break;
				}
			}
			if (!holder) {
				throw new Error(`"${text}" is not one text node under ${selector}`);
			}
			const style = getComputedStyle(holder);
			// Alignment lives on the paragraph, not the run.
			let block: HTMLElement | null = holder;
			while (block && getComputedStyle(block).display === 'inline') {
				block = block.parentElement;
			}
			return {
				fontFamily: style.fontFamily,
				fontSize: style.fontSize,
				fontWeight: style.fontWeight,
				color: style.color,
				textAlign: block ? getComputedStyle(block).textAlign : '',
			};
		},
		[rootSelector, needle] as const,
	);
}

test.describe('inline edit font', () => {
	test('the editor paints the text in the face, size and colour view mode used', async ({
		page,
	}) => {
		await page.route(/^https?:\/\/fonts\.(googleapis|gstatic)\.com\//u, (route) => route.abort());
		await loadDeck(page, WEBFONT_DECK);
		const target = elementWithText(page, PROBE_TEXT);
		const id = await target.getAttribute('data-element-id');
		const view = await paintedText(page, `[data-element-id="${id}"]`, PROBE_WORD);
		expect(view.fontFamily).toContain('ADLaM Display');

		await selectElement(page, target);
		await target.dblclick();
		await expect(page.locator('[data-inline-editor]')).toBeVisible();
		const edit = await paintedText(page, '[data-inline-editor]', PROBE_WORD);

		expect(edit).toEqual(view);
	});
});
