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

test('the Slides sheet lists non-overlapping thumbnails across the width and ends above the bar', async ({
	page,
}) => {
	await load(page);
	await bar(page).getByRole('button', { name: 'Slides' }).tap();
	await page.waitForTimeout(600);
	const result = await page.evaluate(() => {
		const frames = [
			...document.querySelectorAll<HTMLElement>('[data-pptx-chrome="slide-frame"]'),
		].filter(
			(frame) => frame.getBoundingClientRect().width > 0 && frame.getBoundingClientRect().top > 100,
		);
		const rects = frames.map((frame) => frame.getBoundingClientRect());
		let overlaps = 0;
		for (let a = 0; a < rects.length; a += 1) {
			for (let b = a + 1; b < rects.length; b += 1) {
				const x =
					Math.min(rects[a]!.right, rects[b]!.right) - Math.max(rects[a]!.left, rects[b]!.left);
				const y =
					Math.min(rects[a]!.bottom, rects[b]!.bottom) - Math.max(rects[a]!.top, rects[b]!.top);
				if (x > 1 && y > 1) {
					overlaps += 1;
				}
			}
		}
		let scroller: HTMLElement | null = frames[0]?.parentElement ?? null;
		while (scroller && !/(auto|scroll)/u.test(getComputedStyle(scroller).overflowY)) {
			scroller = scroller.parentElement;
		}
		const area = scroller?.getBoundingClientRect();
		const navTop = document.querySelector('pptx-ui-mobile-bar')!.getBoundingClientRect().top;
		return {
			count: frames.length,
			overlaps,
			width: area?.width ?? 0,
			bottom: area?.bottom ?? 0,
			navTop,
		};
	});
	expect(result.count).toBeGreaterThanOrEqual(3);
	expect(result.overlaps).toBe(0);
	expect(result.width).toBeGreaterThan(412 * 0.9);
	expect(result.bottom).toBeLessThanOrEqual(result.navTop + 1);
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

/**
 * Picks on the phone Insert sheet that open a dialog, by tile (and sub-row for New Slide). The
 * dialog must stay up after the sheet closes: it used to live inside the sheet's body, which
 * unmounts with it. `instant` marks a pick a binding may also answer by inserting at once
 * (Date & Time in the bindings without a format picker).
 */
const DIALOG_PICKS: {
	tile: string;
	row?: string;
	dialog: RegExp;
	select?: boolean;
	instant?: boolean;
}[] = [
	{ tile: 'Date & Time', dialog: /date\/time|date & time/i, instant: true },
	{ tile: 'SmartArt', dialog: /smartart/i },
	{ tile: 'Equation', dialog: /equation/i },
	{ tile: 'Link', dialog: /hyperlink/i, select: true },
	{ tile: 'New Slide', row: 'Slide Templates...', dialog: /template/i },
];

for (const pick of DIALOG_PICKS) {
	test(`the phone Insert sheet's ${pick.tile}${pick.row ? ` > ${pick.row}` : ''} opens its dialog after the sheet closes`, async ({
		page,
	}) => {
		await load(page);
		const elements = page.locator('[data-pptx-element="true"]');
		if (pick.select) {
			// A link needs a target to attach to.
			await elements.last().tap();
		}
		const before = await elements.count();
		await bar(page).getByRole('button', { name: 'Insert' }).tap();
		await page.getByRole('button', { name: pick.tile, exact: true }).tap();
		if (pick.row) {
			await page.getByRole('button', { name: pick.row, exact: true }).tap();
		}
		// The sheet closes (its tiles are gone) and the dialog it asked for stays open. Some
		// bindings draw these as plain panels, so a dialog role is not required, but the text must
		// be visible: the desktop ribbon's hidden copy of the same label does not count.
		await expect(page.getByRole('button', { name: 'Text Box', exact: true })).toHaveCount(0);
		const shown = page
			.getByRole('dialog')
			.filter({ hasText: pick.dialog })
			.or(page.getByText(pick.dialog))
			.filter({ visible: true });
		if (pick.instant) {
			await expect
				.poll(async () => (await shown.count()) > 0 || (await elements.count()) > before)
				.toBe(true);
		} else {
			await expect(shown.first()).toBeVisible();
		}
	});
}
