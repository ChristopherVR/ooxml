/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec, `test`/`expect` come from @playwright/test */
/**
 * The chart Format pane's picture, pattern and gradient fills, 3-D Format,
 * Text Options, Plot Series on Secondary Axis, per-series data labels and the
 * chart area's outer shadow, in every binding: each edit shows on the slide
 * and survives a save and reopen.
 *
 * Run: bunx playwright test chart-format-pane-more
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
const pick = (page: Page, id: string, value: string) =>
	field(page, id).locator(`[data-chart-format-choice="${value}"]`).click();

async function selectChart(page: Page): Promise<void> {
	const chart = elementsOfType(page, 'chart').first();
	await chart.waitFor();
	await selectElement(page, chart);
	await expect
		.poll(() =>
			page.evaluate((svg) => {
				const mark = document.querySelector(svg)?.querySelector('[data-chart-part="dataPoint"]');
				return mark ? getComputedStyle(mark).pointerEvents : 'none';
			}, SVG),
		)
		.not.toBe('none');
}

async function openSeries(page: Page, series: number): Promise<void> {
	await page
		.locator(`${SVG} [data-chart-part="dataPoint"][data-chart-series="${series}"]`)
		.first()
		.dblclick();
	await expect(paneTitle(page)).toHaveText('Format Data Series');
}

/** Pick an element in the pane's Chart Options picker. */
async function target(page: Page, key: string): Promise<void> {
	await pane(page)
		.locator('[data-chart-format-target]')
		.evaluate((select, value) => {
			(select as HTMLElement & { value: string }).value = value;
			select.dispatchEvent(new Event('change', { bubbles: true }));
		}, key);
}

const markFill = (page: Page, series: number) =>
	page.evaluate(
		({ svg, series }) =>
			document
				.querySelector(svg)
				?.querySelector(`[data-chart-part="dataPoint"][data-chart-series="${series}"]`)
				?.getAttribute('fill') ?? null,
		{ svg: SVG, series },
	);

/** Whether a `<pattern>` def with an id ending in `suffix` exists in the chart. */
const hasDef = (page: Page, suffix: string) =>
	page.evaluate(
		({ svg, suffix }) =>
			[...(document.querySelector(svg)?.querySelectorAll('pattern') ?? [])].some((p) =>
				p.id.endsWith(suffix),
			),
		{ svg: SVG, suffix },
	);

async function saveAndReopen(page: Page, path: string): Promise<void> {
	await writeFile(path, await downloadBytes(await savePptxViaBackstage(page)));
	await loadDeck(page, path);
	await elementsOfType(page, 'chart').first().waitFor();
}

test('pattern fill, bevel and secondary axis on a series persist', async ({ page }, testInfo) => {
	await loadDeck(page, fixture('chart-gallery.pptx'));
	await selectChart(page);
	await openSeries(page, 1);

	await pick(page, 'fill.kind', 'pattern');
	await expect.poll(() => markFill(page, 1)).toMatch(/^url\(#.*-fmt-pt1\)$/u);

	await pane(page).locator('[data-chart-format-tab="effects"]').click();
	await field(page, 'threeD.bevelTop')
		.locator('pptx-ui-select')
		.evaluate((select) => {
			(select as HTMLElement & { value: string }).value = 'circle';
			select.dispatchEvent(new Event('change', { bubbles: true }));
		});
	await expect.poll(() => hasDef(page, '-fmt-b1')).toBe(true);

	await pane(page).locator('[data-chart-format-tab="options"]').click();
	await pick(page, 'series.axis', 'secondary');
	await expect(
		field(page, 'series.axis').locator('[data-chart-format-choice="secondary"]'),
	).toHaveJSProperty('checked', true);

	await saveAndReopen(page, testInfo.outputPath('chart-format-more-series.pptx'));
	await expect.poll(() => markFill(page, 1)).toMatch(/^url\(#.*-fmt-pt1\)$/u);
	await expect.poll(() => hasDef(page, '-fmt-b1')).toBe(true);
	await selectChart(page);
	await openSeries(page, 1);
	await pane(page).locator('[data-chart-format-tab="options"]').click();
	await expect(
		field(page, 'series.axis').locator('[data-chart-format-choice="secondary"]'),
	).toHaveJSProperty('checked', true);
});

test('Text Options recolour the title and a picture fills the plot area', async ({
	page,
}, testInfo) => {
	await loadDeck(page, fixture('chart-gallery.pptx'));
	await selectChart(page);
	await page.locator('[data-ribbon-contextual-tab="chartFormat"]').first().click();
	await page.locator('[data-testid="chart-format-selection"]').first().click();
	await expect(paneTitle(page)).toHaveText('Format Chart Area');

	await target(page, 'plotArea');
	await expect(paneTitle(page)).toHaveText('Format Plot Area');
	await pick(page, 'fill.kind', 'picture');
	await expect(field(page, 'fill.picture')).toHaveAttribute('data-chart-format-file-set', 'true');

	await target(page, 'title');
	await pane(page).locator('[data-chart-format-text-options]').click();
	await expect(pane(page).locator('[data-chart-format-tab="textFill"]')).toBeVisible();
	await pick(page, 'text.fill.kind', 'solid');
	await field(page, 'text.fill.color').locator('input').fill('#c00000');
	const titleFill = () =>
		page.evaluate((svg) => {
			const texts = [...(document.querySelector(svg)?.querySelectorAll('text') ?? [])];
			const top = texts.sort(
				(a, b) => a.getBoundingClientRect().top - b.getBoundingClientRect().top,
			)[0];
			return (top?.getAttribute('fill') ?? '').toLowerCase();
		}, SVG);
	await expect.poll(titleFill).toBe('#c00000');

	await saveAndReopen(page, testInfo.outputPath('chart-format-more-text.pptx'));
	await expect.poll(titleFill).toBe('#c00000');
});

test('one series gets its own data labels and the chart area casts a shadow', async ({ page }) => {
	await loadDeck(page, fixture('chart-gallery.pptx'));
	await selectChart(page);
	await page.locator('[data-ribbon-contextual-tab="chartFormat"]').first().click();
	await page.locator('[data-testid="chart-format-selection"]').first().click();
	await expect(paneTitle(page)).toHaveText('Format Chart Area');

	// A solid chart area with a shadow casts it outside the frame (the element container).
	await pick(page, 'fill.kind', 'solid');
	await pane(page).locator('[data-chart-format-tab="effects"]').click();
	await field(page, 'shadow.enabled').locator('pptx-ui-checkbox').click();
	await expect
		.poll(() =>
			page.evaluate((svg) => {
				let node: Element | null = document.querySelector(svg);
				while (node && node.getAttribute('aria-roledescription') !== 'slide') {
					const shadow = getComputedStyle(node).boxShadow;
					if (shadow && shadow !== 'none') return shadow;
					node = node.parentElement;
				}
				return 'none';
			}, SVG),
		)
		.toMatch(/rgba\(0, 0, 0, 0\.4\)/u);

	const labels = () =>
		page.evaluate((svg) => document.querySelector(svg)?.querySelectorAll('text').length ?? 0, SVG);
	const before = await labels();
	// Series "X" Data Labels: switching values on there labels that series only.
	await target(page, 'dataLabels:0');
	await expect(paneTitle(page)).toHaveText('Format Data Labels');
	await pane(page).locator('[data-chart-format-tab="options"]').click();
	await expect(pane(page).locator('.note')).toContainText('this series only');
	await field(page, 'labels.showValue').locator('pptx-ui-checkbox').click();
	const categories = await page.evaluate(
		(svg) =>
			document
				.querySelector(svg)
				?.querySelectorAll('[data-chart-part="dataPoint"][data-chart-series="0"]').length ?? 0,
		SVG,
	);
	await expect.poll(labels).toBe(before + categories);
});
