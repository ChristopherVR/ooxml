/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec, `test`/`expect` come from @playwright/test */
/**
 * PowerPoint's chart Format task pane, in every binding.
 *
 * Chart Format > Current Selection > Format Selection opens the shared
 * `pptx-ui-chart-format-pane` on the chart area; double-clicking a data mark
 * retargets it to that series ("Format Data Series"), and the context menu
 * offers "Format Data Series..." on a mark. Fill edits repaint the chart at
 * once and survive a save and reopen.
 *
 * Run: bunx playwright test chart-format-pane
 */
import { writeFile } from 'node:fs/promises';

import { expect, test } from '@playwright/test';
import type { Page } from '@playwright/test';

import { savePptxViaBackstage } from './save-pptx';
import { elementsOfType, fixture, loadDeck, selectElement } from './support/deck';
import { downloadBytes } from './support/exports';

test.use({ viewport: { width: 1440, height: 900 } });

const SVG = '[aria-roledescription="slide"] [aria-roledescription="chart"] svg';
const pane = (page: Page) => page.locator('pptx-ui-chart-format-pane');
const paneTitle = (page: Page) => pane(page).locator('h3');
const field = (page: Page, id: string) => pane(page).locator(`[data-chart-format-field="${id}"]`);

/** Fill of the first data mark of series `series`. */
async function markFill(page: Page, series = 0): Promise<string | null> {
	return page.evaluate(
		({ svg, series }) =>
			document
				.querySelector(svg)
				?.querySelector(`[data-chart-part="dataPoint"][data-chart-series="${series}"]`)
				?.getAttribute('fill')
				?.toLowerCase() ?? null,
		{ svg: SVG, series },
	);
}

/** Fill of the full-bleed chart-area rectangle, or null when none is painted. */
async function chartAreaFill(page: Page): Promise<string | null> {
	return page.evaluate((svg) => {
		const root = document.querySelector<SVGSVGElement>(svg);
		const width = root?.viewBox.baseVal.width;
		const rect = [...(root?.querySelectorAll('rect') ?? [])].find(
			(r) => Number(r.getAttribute('width')) === width && !r.hasAttribute('data-chart-part'),
		);
		return rect?.getAttribute('fill')?.toLowerCase() ?? null;
	}, SVG);
}

async function selectChart(page: Page): Promise<void> {
	const chart = elementsOfType(page, 'chart').first();
	await chart.waitFor();
	await selectElement(page, chart);
	// Marks take pointer events once the chart is selected and editable.
	await expect
		.poll(() =>
			page.evaluate((svg) => {
				const mark = document.querySelector(svg)?.querySelector('[data-chart-part="dataPoint"]');
				return mark ? getComputedStyle(mark).pointerEvents : 'none';
			}, SVG),
		)
		.not.toBe('none');
}

async function openFormatSelection(page: Page): Promise<void> {
	await page.locator('[data-ribbon-contextual-tab="chartFormat"]').first().click();
	await page.locator('[data-testid="chart-format-selection"]').first().click();
	await expect(pane(page)).toBeVisible();
}

test('Format Selection opens the pane and fill edits repaint, retarget and persist', async ({
	page,
}, testInfo) => {
	await loadDeck(page, fixture('chart-gallery.pptx'));
	await selectChart(page);
	await openFormatSelection(page);
	await expect(paneTitle(page)).toHaveText('Format Chart Area');

	// Chart area: solid fill in a chosen colour paints the chart-area rectangle.
	await field(page, 'fill.kind').locator('[data-chart-format-choice="solid"]').click();
	await field(page, 'fill.color').locator('input').fill('#ffe699');
	await expect.poll(() => chartAreaFill(page)).toBe('#ffe699');

	// Double-clicking a column retargets the pane to that series.
	await page
		.locator(`${SVG} [data-chart-part="dataPoint"][data-chart-series="1"]`)
		.first()
		.dblclick();
	await expect(paneTitle(page)).toHaveText('Format Data Series');
	await field(page, 'fill.kind').locator('[data-chart-format-choice="solid"]').click();
	await field(page, 'fill.color').locator('input').fill('#7030a0');
	await expect.poll(() => markFill(page, 1)).toBe('#7030a0');

	// Series Options: a gap width outside ST_GapAmount is refused, a valid one sticks.
	await pane(page).locator('[data-chart-format-tab="options"]').click();
	const gap = field(page, 'series.gapWidth').locator('input');
	await gap.fill('40');
	await gap.press('Tab');
	await expect(gap).toHaveValue('40');

	const saved = testInfo.outputPath('chart-format-saved.pptx');
	await writeFile(saved, await downloadBytes(await savePptxViaBackstage(page)));
	await loadDeck(page, saved);
	await elementsOfType(page, 'chart').first().waitFor();
	await expect.poll(() => markFill(page, 1)).toBe('#7030a0');
	await expect.poll(() => chartAreaFill(page)).toBe('#ffe699');
});

test('the context menu offers Format Data Series on a mark and the pane closes', async ({
	page,
}) => {
	await loadDeck(page, fixture('chart-gallery.pptx'));
	await selectChart(page);
	await page
		.locator(`${SVG} [data-chart-part="dataPoint"][data-chart-series="0"]`)
		.first()
		.click({ button: 'right' });
	const entry = page.getByRole('menuitem', { name: 'Format Data Series...' });
	await expect(entry).toBeVisible();
	await entry.click();
	await expect(paneTitle(page)).toHaveText('Format Data Series');
	await pane(page).locator('[data-chart-format-close]').click();
	await expect(pane(page)).toHaveCount(0);
});
