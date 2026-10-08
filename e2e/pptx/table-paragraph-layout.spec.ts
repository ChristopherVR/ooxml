/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec, `test`/`expect` come from @playwright/test */
/**
 * Paragraphs inside a table cell keep their own layout, in every framework
 * demo.
 *
 * FIXTURE (`src/core/pptx/__tests__/fixtures/table-paragraph-layout.pptx`,
 * shared with the core parsing test) is saved by PowerPoint for Mac 16. Each
 * row's second cell sets one paragraph property: per paragraph alignment,
 * exact and proportional line spacing, space after, or a hanging indent.
 * Before, every cell drew its paragraphs as one run stream with the cell's
 * default spacing.
 */
import { fileURLToPath } from 'node:url';

import { expect, test } from '@playwright/test';

import { loadDeck, slideStage } from './support/deck';

const FIXTURE = fileURLToPath(
	new URL('../../src/core/pptx/__tests__/fixtures/table-paragraph-layout.pptx', import.meta.url),
);

test('table cell paragraphs keep their own alignment and spacing', async ({ page }) => {
	await page.setViewportSize({ width: 1600, height: 1000 });
	await loadDeck(page, FIXTURE);
	const stage = slideStage(page);

	// Alignment per paragraph: the first paragraph ends at the cell's right
	// edge, the second sits in its middle.
	const right = await stage
		.getByText('Right aligned first paragraph', { exact: true })
		.boundingBox();
	const centred = await stage.getByText('Centered second paragraph', { exact: true }).boundingBox();
	expect(right).not.toBeNull();
	expect(centred).not.toBeNull();
	expect(right!.x + right!.width).toBeGreaterThan(centred!.x + centred!.width + 20);

	// Space after 12pt: consecutive 11pt paragraphs sit about two lines apart
	// (13.2pt line plus 12pt), not one.
	const first = await stage.getByText('First paragraph', { exact: true }).boundingBox();
	const second = await stage.getByText('Second paragraph', { exact: true }).boundingBox();
	expect(first).not.toBeNull();
	expect(second).not.toBeNull();
	const pitch = second!.y - first!.y;
	expect(pitch / first!.height).toBeGreaterThan(1.6);
});
