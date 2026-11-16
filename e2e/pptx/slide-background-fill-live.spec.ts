/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec, `test`/`expect` come from @playwright/test */
/**
 * Choosing "Slide background fill" in the Fill section paints the shape with
 * the slide background at once, without a reopen: the slide stage publishes
 * its resolved background as `--pptx-slide-bg-*` custom properties and the
 * shared fill resolver paints the shape with them (`slide-background-fill.ts`).
 *
 * Run: bunx playwright test slide-background-fill-live
 */
import { expect, test } from '@playwright/test';
import type { Locator, Page } from '@playwright/test';

import { stageElements } from './support/context-menu';
import { inspector, loadDeck, selectElement } from './support/deck';
import { openFormatSection } from './support/format-pane';
import { chooseSelectValue, expectSelectValue } from './support/select-control';

test.use({ viewport: { width: 1440, height: 900 } });

function canvasShape(page: Page, needle: string): Locator {
	return stageElements(page).filter({ hasText: needle }).first();
}

/** The painted fill of `shape` (the node carrying the background-fill paint), if any. */
function backgroundFillPaint(shape: Locator): Promise<{ color: string; published: string } | null> {
	return shape.evaluate((root) => {
		const nodes = [root, ...root.querySelectorAll('*')];
		const painted = nodes.find((node) =>
			(node.getAttribute('style') ?? '').includes('var(--pptx-slide-bg-color'),
		);
		if (!painted) {
			return null;
		}
		const computed = getComputedStyle(painted);
		return {
			color: computed.backgroundColor,
			published: computed.getPropertyValue('--pptx-slide-bg-color').trim(),
		};
	});
}

test('a shape switched to the slide background fill paints it live', async ({ page }) => {
	await loadDeck(page);
	const shape = canvasShape(page, 'Product Overview');
	await selectElement(page, shape);
	await expect(inspector(page)).toBeVisible();
	const fill = await openFormatSection(page, 'fill');
	const kind = fill.locator('[data-testid="fill-mode"]').first();
	await chooseSelectValue(page, kind, 'background');
	await expectSelectValue(kind, 'background');

	await expect.poll(() => backgroundFillPaint(shape)).not.toBeNull();
	const paint = (await backgroundFillPaint(shape))!;
	// The stage published the slide background, and the shape paints it.
	expect(paint.published).not.toBe('');
	expect(paint.color).not.toBe('rgba(0, 0, 0, 0)');
});
