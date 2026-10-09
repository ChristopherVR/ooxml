/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright */
/** Reporter IHAGI-c's reproduction for #37, shared by every binding.
 * Fixture: https://raw.githubusercontent.com/IHAGI-c/ooxml/pr-assets/decks/slide-jump-links.pptx
 */
import { fileURLToPath } from 'node:url';

import { expect, test } from '@playwright/test';
import type { Page } from '@playwright/test';
import type { PowerPointViewerAPI } from 'ooxml-ui/pptx';

import { loadDeck } from './support/deck';

const fixture = fileURLToPath(new URL('./fixtures/slide-jump-links.pptx', import.meta.url));
type ViewerWindow = Window & { __pptxViewer: PowerPointViewerAPI };

// A show overlay follows the editing canvas in DOM order, even when the
// covered editing canvas is still mounted.
const stage = (page: Page) => page.locator('[aria-roledescription="slide"]:visible').last();

function linkedElement(page: Page, text: string) {
	// Some bindings split a linked run into words for text layout. Click
	// the link itself, rather than the centre of its wider text box.
	return text.startsWith('A.')
		? stage(page).getByRole('link', { name: /^A\./u }).first()
		: stage(page).getByText(text, { exact: true });
}

async function setMode(page: Page, mode: 'preview' | 'edit' | 'present'): Promise<void> {
	await page.waitForFunction(() => Boolean((window as ViewerWindow).__pptxViewer));
	await page.evaluate((next) => (window as ViewerWindow).__pptxViewer.setMode(next), mode);
}

async function activeSlide(page: Page): Promise<number> {
	return page.evaluate(() => (window as ViewerWindow).__pptxViewer.getActiveSlideIndex());
}

for (const text of ['A. Text link to slide 3', 'B. Shape link to slide 3']) {
	for (const mode of ['preview', 'present'] as const) {
		test(`${mode}: ${text} jumps directly to slide 3 without a new tab`, async ({
			page,
			context,
		}) => {
			await loadDeck(page, fixture);
			await setMode(page, mode);
			const pages = context.pages().length;
			await linkedElement(page, text).click();
			await expect(stage(page).getByText('Slide 3 (link target)', { exact: true })).toBeVisible();
			expect(context.pages()).toHaveLength(pages);
		});
	}
}

test('preview hides authoring badges and lets Enter follow an internal text link', async ({
	page,
}) => {
	await loadDeck(page, fixture);
	await setMode(page, 'preview');
	await expect(stage(page).locator('.pptx-action-indicator')).toBeHidden();
	const link = stage(page).getByRole('link', { name: /^A\./u }).first();
	await expect(link).not.toHaveAttribute('href', /slide3\.xml/u);
	await link.press('Enter');
	await expect.poll(() => activeSlide(page)).toBe(2);
});

for (const text of ['A. Text link to slide 3', 'B. Shape link to slide 3']) {
	test(`editing: ${text} requires Ctrl/Cmd+Click to navigate`, async ({ page }) => {
		await loadDeck(page, fixture);
		await setMode(page, 'edit');
		const link = linkedElement(page, text);
		await expect(link).not.toHaveAttribute('href', /./u);
		await link.click();
		expect(await activeSlide(page)).toBe(0);
		await link.click({ modifiers: ['ControlOrMeta'] });
		await expect.poll(() => activeSlide(page)).toBe(2);
	});
}
