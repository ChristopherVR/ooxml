/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec */
import { expect, test } from '@playwright/test';

import { loadDeck, SAMPLE_DECK } from './support/deck';

/**
 * The viewer stylesheet's base reset is scoped to the viewer, so host text keeps its own styles.
 * A viewer mounted inside host text the browser styles (here a heading) must still start from
 * the page's base text, as it did when the reset flattened every heading on the page.
 */
test('a viewer inside a host heading starts from the base text', async ({ page }) => {
	await loadDeck(page, SAMPLE_DECK);
	const styles = await page.evaluate(() => {
		const scopes = '[data-pptx-viewer],[data-pptx-viewport],[data-pptx-scope]';
		let root = document.querySelector('[data-pptx-viewport]');
		while (root?.parentElement?.closest(scopes)) {
			root = root.parentElement.closest(scopes);
		}
		if (!root?.parentElement) {
			throw new Error('viewer root not found');
		}
		const heading = document.createElement('h1');
		root.parentElement.insertBefore(heading, root);
		heading.append(root);
		const read = (element: Element) => {
			const style = getComputedStyle(element);
			return {
				fontSize: style.fontSize,
				fontWeight: style.fontWeight,
				fontStyle: style.fontStyle,
			};
		};
		return { root: read(root), base: getComputedStyle(document.documentElement).fontSize };
	});
	// Whether the host heading itself is reset is the host page's business: some demo pages load
	// their own Tailwind base styles.
	expect(styles.root).toStrictEqual({
		fontSize: styles.base,
		fontWeight: '400',
		fontStyle: 'normal',
	});
});
