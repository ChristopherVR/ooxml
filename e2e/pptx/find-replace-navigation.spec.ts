/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright owns the test globals */
/**
 * Find and Replace navigation, in every binding.
 *
 * ## The drift this pins
 *
 * Each binding kept its own copy of the Find and Replace session around the
 * shared search, and the copies had drifted: one had no Next or Previous at
 * all, so its status always read "1 of N" and Replace always rewrote the first
 * match; two went to the matched slide without selecting the matched element;
 * and only one dropped its matches when the bar closed. The session now lives
 * once in `src/ui/src/pptx/render/find-replace-session.ts`, and this spec is
 * the framework-neutral proof that every binding drives it the same way.
 *
 * The sample deck has three matches of "quarter" on two slides: "Quarterly
 * growth" in the agenda (slide 2), and the "Quarterly Growth" title and the
 * "by quarter" subtitle of the chart slide (slide 3). The status is read the
 * way the other find specs read it, as "n of N" or "n / N".
 */
import { expect, test } from '@playwright/test';
import type { Locator, Page } from '@playwright/test';

import { loadDeck } from './support/deck';
import { armKeyboard } from './support/keyboard';

test.use({ viewport: { width: 1440, height: 900 } });

/** The slice of the development handle (`window.__pptxViewer`) this spec reads. */
interface HandleWindow {
	__pptxViewer: {
		getActiveSlideIndex(): number;
		getSelectedElementIds(): string[];
		getSlides(): Array<{ elements: Array<{ id: string; text?: string }> }>;
	};
}

/** The bar's "n of N" status line. */
const status = (page: Page, current: number, total: number): Locator =>
	page
		.getByText(new RegExp(`^\\s*${current}\\s+(?:of|/)\\s+${total}\\s*$`, 'u'))
		.filter({ visible: true });

/** Every element whose text holds "quarter" (any case), by id, with the slide it is on. */
function quarterElements(page: Page) {
	return page.evaluate(() =>
		(window as unknown as HandleWindow).__pptxViewer
			.getSlides()
			.flatMap((slide, slideIndex) =>
				slide.elements
					.filter((element) => /quarter/iu.test(element.text ?? ''))
					.map((element) => ({ slideIndex, id: element.id, text: element.text ?? '' })),
			),
	);
}

/** The active slide and the text of each selected element on it. */
function selection(page: Page) {
	return page.evaluate(() => {
		const api = (window as unknown as HandleWindow).__pptxViewer;
		const active = api.getActiveSlideIndex();
		const elements = api.getSlides()[active]?.elements ?? [];
		return {
			active,
			selected: api
				.getSelectedElementIds()
				.map((id) => elements.find((element) => element.id === id)?.text ?? id),
		};
	});
}

async function openFindReplace(page: Page): Promise<void> {
	await armKeyboard(page);
	await page.keyboard.press('Control+h');
	await expect(page.locator('input[placeholder*="Find" i]').first()).toBeVisible();
}

test('next selects the second match, replace changes only that match, closing clears the status', async ({
	page,
}) => {
	await loadDeck(page);
	const before = await quarterElements(page);
	expect(before.map(({ slideIndex }) => slideIndex)).toEqual([1, 2, 2]);

	await openFindReplace(page);
	await page.locator('input[placeholder*="Find" i]').first().fill('quarter');
	await expect(status(page, 1, 3)).toBeVisible();

	await page.getByRole('button', { name: 'Next match', exact: true }).first().click();
	await expect(status(page, 2, 3)).toBeVisible();
	await expect.poll(() => selection(page)).toEqual({ active: 2, selected: [before[1]!.text] });

	await page.locator('input[placeholder*="Replace with" i]').first().fill('Annual');
	await page.getByRole('button', { name: /^Replace(?: current match)?$/u }).click();
	await expect
		.poll(async () => {
			const slides = await page.evaluate(() =>
				(window as unknown as HandleWindow).__pptxViewer.getSlides(),
			);
			return before.map(
				({ slideIndex, id }) =>
					slides[slideIndex]?.elements.find((element) => element.id === id)?.text,
			);
		})
		.toEqual([before[0]!.text, before[1]!.text.replace(/quarter/iu, 'Annual'), before[2]!.text]);
	// Replace moves on to the match after the replaced one.
	await expect(status(page, 2, 2)).toBeVisible();

	await page.getByRole('button', { name: /^Close find (?:and replace|bar)$/u }).click();
	await expect(status(page, 2, 2)).toHaveCount(0);

	// Reopened, the bar searches the query it kept again, from the first match.
	await openFindReplace(page);
	await expect(status(page, 1, 2)).toBeVisible();
});
