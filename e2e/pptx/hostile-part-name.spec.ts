/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec, `test`/`expect` come from @playwright/test */
/**
 * A package controls its part names, and element ids used to be built from them. A name holding
 * markup must stay inert in every binding: the deck opens, and nothing from the name joins the page.
 */
import { expect, test } from '@playwright/test';

import { resetTabSession, slideStage } from './support/deck';
import { INJECTED_ATTRIBUTE, hostilePartNameDeck } from './support/hostile-part-name-deck';

test('a slide part name holding markup adds nothing to the page', async ({ page }) => {
	await resetTabSession(page);
	await page.goto('/');
	await page.locator('#file-input').setInputFiles({
		name: 'hostile-part-name.pptx',
		mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
		buffer: await hostilePartNameDeck(),
	});

	const shape = slideStage(page).locator('[data-element-id*="%3C"]').first();
	await expect(shape).toBeAttached();
	await expect(slideStage(page)).toContainText('Glow');

	const injected = await page.evaluate((attribute) => {
		let count = 0;
		const visit = (root: Document | ShadowRoot) => {
			for (const element of root.querySelectorAll('*')) {
				if (element.hasAttribute(attribute)) {
					count += 1;
				}
				if (element.shadowRoot) {
					visit(element.shadowRoot);
				}
			}
		};
		visit(document);
		return count;
	}, INJECTED_ATTRIBUTE);
	expect(injected).toBe(0);
});
