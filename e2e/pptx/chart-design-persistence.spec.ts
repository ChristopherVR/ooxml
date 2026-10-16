/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec, `test`/`expect` come from @playwright/test */
/**
 * Chart Design styling survives a save and reopen, in every binding.
 *
 * Selecting a chart brings up the Chart Design tab; its Change Colors and
 * Chart Styles galleries recolour the plotted marks. Those edits must reach
 * the saved `.pptx`, so reopening the downloaded file paints the same colour.
 * The OOXML Office suite (which embeds the vanilla binding) runs the same
 * flow in `scripts/suite-pptx-chart-check.mjs`, together with the properties
 * panel layout checks that need a host without a CSS reset.
 *
 * Run: bunx playwright test chart-design-persistence
 */
import { writeFile } from 'node:fs/promises';

import { expect, test } from '@playwright/test';
import type { Page } from '@playwright/test';

import { savePptxViaBackstage } from './save-pptx';
import { elementsOfType, fixture, inspector, loadDeck, selectElement } from './support/deck';
import { downloadBytes } from './support/exports';

test.use({ viewport: { width: 1440, height: 900 } });

/** The fill of the first plotted mark (series 0) of the chart on the main slide. */
async function firstMarkFill(page: Page): Promise<string | null> {
	return page.evaluate(
		() =>
			document
				.querySelector('[aria-roledescription="slide"] [aria-roledescription="chart"] svg')
				?.querySelector('[data-chart-part="dataPoint"][data-chart-series="0"]')
				?.getAttribute('fill')
				?.toLowerCase() ?? null,
	);
}

async function openChartDesign(page: Page): Promise<void> {
	const chart = elementsOfType(page, 'chart').first();
	await chart.waitFor();
	await selectElement(page, chart);
	await page.locator('[data-ribbon-contextual-tab="chartDesign"]').first().click();
}

const gallery = (page: Page, name: string) =>
	page.locator(`pptx-ui-ribbon-gallery:has([data-ribbon-gallery="${name}"])`).first();

test('chart colours and styles apply and persist through save and reopen', async ({
	page,
}, testInfo) => {
	await loadDeck(page, fixture('chart-gallery.pptx'));
	await openChartDesign(page);
	const authored = await firstMarkFill(page);
	expect(authored).not.toBeNull();

	// Chart Styles (inline gallery): a different style restyles the series.
	await gallery(page, 'chartStyles').locator('[data-gallery-item]').nth(4).click();
	await expect.poll(() => firstMarkFill(page)).not.toBe(authored);

	// Change Colors (dropdown gallery): a monochromatic palette recolours it again.
	const styled = await firstMarkFill(page);
	await gallery(page, 'chartColors').locator('[data-ribbon-gallery]').click();
	await page.locator('[data-ribbon-gallery-popup] [data-gallery-item]').nth(5).click();
	await expect.poll(() => firstMarkFill(page)).not.toBe(styled);
	const recoloured = await firstMarkFill(page);

	const saved = testInfo.outputPath('chart-design-saved.pptx');
	await writeFile(saved, await downloadBytes(await savePptxViaBackstage(page)));
	await loadDeck(page, saved);
	await elementsOfType(page, 'chart').first().waitFor();
	await expect.poll(() => firstMarkFill(page)).toBe(recoloured);
});

test('the chart properties lead with Title, Chart type and Grouping, one per row', async ({
	page,
}) => {
	await loadDeck(page, fixture('chart-gallery.pptx'));
	await openChartDesign(page);
	const panel = inspector(page);
	const rows = [
		panel.getByLabel('Title', { exact: true }).first(),
		panel.getByLabel('Chart type', { exact: true }).first(),
		panel.getByLabel('Grouping', { exact: true }).first(),
	];
	await expect(rows[0]).toBeVisible();
	const boxes = await Promise.all(rows.map((row) => row.boundingBox()));
	// Every binding orders the three the same way, and none shares a line with the next.
	for (let index = 1; index < boxes.length; index++) {
		const upper = boxes[index - 1];
		const lower = boxes[index];
		expect(upper).not.toBeNull();
		expect(lower).not.toBeNull();
		expect(lower!.y).toBeGreaterThanOrEqual(upper!.y + upper!.height - 1);
	}
	await expect(panel).not.toContainText(/\{\{\w+\}\}/u);
});
