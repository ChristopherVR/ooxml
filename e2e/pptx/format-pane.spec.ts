/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec, `test`/`expect` come from @playwright/test */
/**
 * The inspector's Properties tab is a selection-aware Format pane: collapsible
 * sections chosen by the shared model for what is selected, open state kept
 * while the user moves between elements of the same kind, and "Format Shape"
 * from the context menu revealing (expanding and focusing) the Fill section.
 * Hosts hide sections through `inspector.hiddenSections`.
 *
 * Run: bunx playwright test format-pane
 */
import { expect, test } from '@playwright/test';
import type { Locator, Page } from '@playwright/test';

import { chooseCommand, openMenuOn, stageElements } from './support/context-menu';
import { inspector, loadDeck, SAMPLE_DECK, selectElement, slideStage } from './support/deck';
import { formatSection } from './support/format-pane';

test.use({ viewport: { width: 1440, height: 900 } });

function canvasShape(page: Page, needle: string): Locator {
	return stageElements(page).filter({ hasText: needle }).first();
}

function withCustomization(customization: object): string {
	return `/?customization=${encodeURIComponent(JSON.stringify(customization))}`;
}

/** The sections the shared model opens first for a text box or a shape with text. */
const FIRST_OPEN: Record<string, string[]> = {
	text: ['font', 'paragraph'],
	shape: ['fill', 'outline'],
};

test('the Properties tab follows the selection with collapsible sections', async ({ page }) => {
	await loadDeck(page);
	const sections = inspector(page).locator('[data-pptx-format-sections]').first();
	await expect(sections).toHaveAttribute('data-pptx-format-sections', 'slide');
	await expect(formatSection(page, 'slide')).toHaveAttribute('open', '');

	await selectElement(page, canvasShape(page, 'Product Overview'));
	await expect(sections).toHaveAttribute('data-pptx-format-sections', /^(text|shape)$/u);
	const category = (await sections.getAttribute('data-pptx-format-sections'))!;
	for (const id of FIRST_OPEN[category]!) {
		await expect(formatSection(page, id)).toHaveAttribute('open', '');
	}
	const size = formatSection(page, 'sizePosition');
	await expect(size).not.toHaveAttribute('open', '');

	// Open Size and position: it stays open for the next element of the same kind.
	await size.locator('summary').click();
	await expect(size).toHaveAttribute('open', '');
	await selectElement(page, canvasShape(page, 'Q2 2026'));
	await expect(sections).toHaveAttribute('data-pptx-format-sections', category);
	await expect(formatSection(page, 'sizePosition')).toHaveAttribute('open', '');
});

test('the Format pane edits each Summary Zoom tile independently', async ({ page }) => {
	await loadDeck(page);
	await page.evaluate(() => {
		window.__pptxViewer.addElement({
			type: 'zoom',
			id: 'summary-format-probe',
			zoomType: 'summary',
			targetSlideIndex: 1,
			x: 250,
			y: 260,
			width: 400,
			height: 100,
			summaryTargets: [
				{
					title: 'Introduction',
					sectionId: 'intro',
					targetSlideIndex: 1,
					x: 250,
					y: 260,
					width: 200,
					height: 100,
				},
				{
					title: 'Details',
					sectionId: 'details',
					targetSlideIndex: 2,
					x: 450,
					y: 260,
					width: 200,
					height: 100,
				},
			],
		});
	});

	const zoom = slideStage(page).locator('[data-zoom-type="summary"]');
	await expect(zoom).toBeVisible();
	await selectElement(page, zoom);
	const size = formatSection(page, 'sizePosition');
	await expect(size).toBeVisible();
	if ((await size.getAttribute('open')) === null) await size.locator('summary').click();
	const tiles = size.getByRole('combobox', { name: 'Summary Zoom', exact: true });
	await tiles.click();
	await expect(size.getByRole('option', { name: 'Introduction', exact: true })).toBeVisible();
	await expect(size.getByRole('option', { name: 'Details', exact: true })).toBeVisible();
	await tiles.press('Escape');
	const duration = size.getByRole('spinbutton', { name: 'Duration (ms)', exact: true }).nth(1);
	await duration.fill('750');
	await duration.press('Tab');

	await expect
		.poll(() =>
			page.evaluate(() => {
				const element = window.__pptxViewer
					.getElements(0)
					.find((candidate) => candidate.type === 'zoom' && candidate.zoomType === 'summary');
				return element?.type === 'zoom'
					? element.summaryTargets?.map((tile) => tile.transitionDurationMs ?? 0)
					: null;
			}),
		)
		.toEqual([0, 750]);
});

test('Format Shape reveals and focuses the Fill section', async ({ page }) => {
	await loadDeck(page);
	const shape = canvasShape(page, 'Product Overview');
	await selectElement(page, shape);
	const fill = formatSection(page, 'fill');
	// Close Fill first, so the launcher has to open it again.
	if ((await fill.getAttribute('open')) !== null) {
		await fill.locator('summary').click();
	}
	await expect(fill).not.toHaveAttribute('open', '');
	await openMenuOn(page, shape);
	// `chooseCommand` matches a pattern: escape the label's ellipsis.
	await chooseCommand(page, 'Format Shape\\.\\.\\.');
	await expect(fill).toHaveAttribute('open', '');
	await expect
		.poll(() =>
			page.evaluate(() => document.activeElement?.getAttribute('data-pptx-format-section') ?? null),
		)
		.toBe('fill');
});

test('a host hides and reorders Format pane sections through customization', async ({ page }) => {
	await loadDeck(
		page,
		SAMPLE_DECK,
		withCustomization({
			inspector: { hiddenSections: ['effects'], sectionOrder: ['sizePosition'] },
		}),
	);
	await selectElement(page, canvasShape(page, 'Product Overview'));
	await expect(formatSection(page, 'arrange')).toBeVisible();
	await expect(formatSection(page, 'effects')).toBeHidden();
	// Size and position is drawn first, above the sections the selection opens.
	const top = async (id: string) => (await formatSection(page, id).boundingBox())!.y;
	await expect.poll(async () => (await top('sizePosition')) < (await top('arrange'))).toBe(true);
	const first = await inspector(page)
		.locator('[data-pptx-format-sections] > [data-pptx-format-section]:visible')
		.evaluateAll(
			(nodes) =>
				nodes
					.map((node) => ({
						id: node.getAttribute('data-pptx-format-section'),
						y: node.getBoundingClientRect().top,
					}))
					.sort((a, b) => a.y - b.y)[0]?.id,
		);
	expect(first).toBe('sizePosition');
});
