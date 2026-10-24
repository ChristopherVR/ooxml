/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec, `test`/`expect` come from @playwright/test */
/**
 * The inspector never scrolls sideways.
 *
 * Vanilla's inspector scrolled 10 px sideways at 1440x900 as soon as a text box was selected: the
 * Accessibility section's title field put its caption and a text input of the browser's default
 * width side by side in an inline box, which grows with its content instead of the panel, so the row
 * ran past the panel's edge. Run on every binding: with a text box selected, no box of the visible
 * inspector that scrolls (the panel itself or anything inside it) is wider than its own content box.
 */
import { expect, test } from '@playwright/test';
import type { Locator } from '@playwright/test';

import { loadDeck, SAMPLE_DECK } from './support/deck';

test.use({ viewport: { width: 1440, height: 900 } });

/** Every scrolling box in `inspector` whose content is wider than it, as `tag.class scroll>client`. */
function sidewaysScrollers(inspector: Locator): Promise<string[]> {
	return inspector.evaluate((root) =>
		[root, ...root.querySelectorAll('*')].flatMap((node) => {
			const overflowX = getComputedStyle(node).overflowX;
			const scrolls = overflowX === 'auto' || overflowX === 'scroll';
			return scrolls && node.clientWidth > 0 && node.scrollWidth > node.clientWidth + 1
				? [
						`${node.localName}.${[...node.classList].join('.')} ${node.scrollWidth}>${node.clientWidth}`,
					]
				: [];
		}),
	);
}

test('the inspector does not scroll sideways with a text box selected', async ({ page }) => {
	await loadDeck(page, SAMPLE_DECK);
	await page.locator('[data-pptx-element="true"]').filter({ hasText: 'Project' }).first().click();
	const inspector = page.locator('[data-pptx-inspector]:visible').first();
	// The selection's own sections are on screen (the Accessibility section is the one that ran wide).
	await expect(inspector.getByLabel('Accessibility title', { exact: true }).first()).toBeAttached();
	await expect.poll(() => sidewaysScrollers(inspector)).toStrictEqual([]);
});
