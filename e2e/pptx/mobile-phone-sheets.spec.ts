/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec, `test`/`expect` come from @playwright/test */
/**
 * Phone sheets (412x915, touch): the Format sheet is one column with a compact tab row and
 * nothing clipped sideways, a sheet stacks above the selected shape's handles, the comment
 * box fills the sheet, and the slide is inset from both screen edges.
 */
import { resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

import { test, expect } from '@playwright/test';
import type { Page } from '@playwright/test';

import { resetTabSession } from './support/deck';

const deck = resolve(
	fileURLToPath(new URL('../../viewers/pptx/.github/assets/sample-deck.pptx', import.meta.url)),
);

test.use({ viewport: { width: 412, height: 915 }, hasTouch: true, isMobile: true });

async function load(page: Page): Promise<void> {
	await resetTabSession(page);
	await page.goto('/');
	await page.locator('#file-input').setInputFiles(deck);
	await page.locator('[data-pptx-element="true"]').first().waitFor();
	await page.waitForTimeout(500);
}

const bar = (page: Page) => page.getByRole('navigation', { name: 'Editor actions' });

test('the slide is inset from both screen edges', async ({ page }) => {
	await load(page);
	const box = await page.locator('[aria-roledescription="slide"]').first().boundingBox();
	expect(box).not.toBeNull();
	expect(box!.x).toBeGreaterThanOrEqual(8);
	expect(box!.x + box!.width).toBeLessThanOrEqual(412 - 8);
});

test('the Format sheet is one column and nothing is clipped sideways', async ({ page }) => {
	await load(page);
	await bar(page).getByRole('button', { name: 'Format' }).tap();
	const tabs = page.getByRole('tab');
	await expect(tabs.first()).toBeVisible();
	const boxes = await tabs.evaluateAll((items) =>
		items.map((item) => {
			const { x, y, width, height } = item.getBoundingClientRect();
			return { x, y, width, height };
		}),
	);
	expect(boxes.length).toBeGreaterThanOrEqual(3);
	// A compact row: one line, none of them a tall tile.
	expect(Math.max(...boxes.map((b) => b.y)) - Math.min(...boxes.map((b) => b.y))).toBeLessThan(8);
	expect(Math.max(...boxes.map((b) => b.height))).toBeLessThanOrEqual(64);
	// Properties sit below the tabs, full width, within the screen.
	const clipped = await page.evaluate(() => {
		const tab = document.querySelector('[role="tab"]');
		const root = tab?.closest('aside, section, [role="dialog"]') ?? document.body;
		return [...root.querySelectorAll('input, select, button, label')]
			.map((element) => element.getBoundingClientRect())
			.filter((rect) => rect.width > 0 && (rect.right > innerWidth + 1 || rect.left < -1)).length;
	});
	expect(clipped).toBe(0);
	const scroll = await page.evaluate(() => document.documentElement.scrollWidth);
	expect(scroll).toBeLessThanOrEqual(413);
});

test('Insert opens the phone sheet and a Text Box tile inserts and closes it', async ({ page }) => {
	await load(page);
	const elements = page.locator('[data-pptx-element="true"]');
	const before = await elements.count();
	await bar(page).getByRole('button', { name: 'Insert' }).tap();
	for (const name of [
		'Text Box',
		'Shapes',
		'Pictures',
		'Table',
		'Chart',
		'SmartArt',
		'Equation',
		'Video',
		'New Slide',
	]) {
		await expect(page.getByRole('button', { name, exact: true })).toBeVisible();
	}
	await page.getByRole('button', { name: 'Text Box', exact: true }).tap();
	await expect(page.getByRole('button', { name: 'Text Box', exact: true })).toHaveCount(0);
	await expect(elements).toHaveCount(before + 1);
});

test('the comment box fills the sheet width', async ({ page }) => {
	await load(page);
	await bar(page).getByRole('button', { name: 'Comments' }).tap();
	await page.waitForTimeout(500);
	// The compose box is the widest visible textarea in the lower half, where the sheet sits.
	const widest = await page.evaluate(() =>
		Math.max(
			0,
			...[...document.querySelectorAll('textarea')]
				.map((field) => field.getBoundingClientRect())
				.filter((rect) => rect.width > 0 && rect.top > innerHeight * 0.4)
				.map((rect) => rect.width),
		),
	);
	expect(widest).toBeGreaterThan(412 * 0.8);
});
