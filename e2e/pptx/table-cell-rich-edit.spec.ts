/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec, `test`/`expect` come from @playwright/test */
/**
 * Rich table cell editing, run identically against every framework demo.
 *
 * PowerPoint edits a table cell as rich text: typing into a word keeps the
 * runs around it formatted as they were, Tab moves to the next cell (and Tab
 * in the last cell adds a row), Escape leaves the cell keeping what was typed,
 * and tab-separated text pasted into a cell fills the cells from it. Every
 * binding mounts the same shared editor (`openTableCellTextEditor`), so one
 * spec covers the behaviour for all five.
 *
 * FIXTURE: slide 4 of `e2e/fixtures/table-styling.pptx` (PowerPoint's own
 * output) holds a cell "Revenue grew 42%" whose second run is formatted.
 */
import { resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

import { test, expect } from '@playwright/test';
import type { Locator, Page } from '@playwright/test';

import { centreOf, selectTableCell } from './support/context-menu';
import { resetTabSession } from './support/deck';
import { nextPaint, selectRibbonTab, selectSlide } from './support/render-ready';

const fixturePath = resolve(
	fileURLToPath(new URL('./fixtures/table-styling.pptx', import.meta.url)),
);

async function loadDeck(page: Page): Promise<void> {
	await page.setViewportSize({ width: 1600, height: 1000 });
	await resetTabSession(page);
	await page.goto('/');
	await page.locator('#file-input').setInputFiles(fixturePath);
	await page.locator('[aria-label="Go to slide 5"]').first().waitFor({ timeout: 60_000 });
	await page.waitForFunction(() => document.fonts.status === 'loaded');
	await nextPaint(page);
	await selectSlide(page, 4);
}

function slideTable(page: Page): Locator {
	return page.locator('[aria-roledescription="slide"]').first().locator('table').first();
}

function cellWithText(page: Page, text: string): Locator {
	return slideTable(page)
		.locator('td')
		.filter({ hasText: new RegExp(`^\\s*${text}\\s*$`, 'u') })
		.first();
}

async function editCell(page: Page, cell: Locator): Promise<Locator> {
	const box = (await cell.boundingBox())!;
	await page.mouse.dblclick(box.x + box.width / 2, box.y + box.height / 2);
	const editor = slideTable(page).locator('[data-table-cell-editor]');
	await expect(editor).toBeVisible();
	return editor;
}

test.describe('rich table cell editing', () => {
	test('typing into one run keeps the other run formatted', async ({ page }) => {
		await loadDeck(page);
		const editor = await editCell(page, cellWithText(page, 'Revenue grew 42%'));
		await expect(editor.locator('span')).toHaveCount(2);
		await page.keyboard.press('Control+Home');
		await page.keyboard.type('Net ');
		await page.keyboard.press('Escape');
		await expect(editor).toBeHidden();
		const cell = cellWithText(page, 'Net Revenue grew 42%');
		await expect(cell).toBeVisible();
		expect(await cell.locator('span').allTextContents()).toEqual(['Net Revenue ', 'grew 42%']);
	});

	test('Tab moves to the next cell, and Tab in the last cell adds a row', async ({ page }) => {
		await loadDeck(page);
		const rowsBefore = await slideTable(page).locator('tr').count();
		const cells = slideTable(page).locator('td');
		const last = cells.nth((await cells.count()) - 1);
		const editor = await editCell(page, last);
		await expect(editor).toBeVisible();
		await page.keyboard.press('Tab');
		await expect(slideTable(page).locator('tr')).toHaveCount(rowsBefore + 1);
		const newRow = slideTable(page).locator('tr').nth(rowsBefore);
		await expect(newRow.locator('[data-table-cell-editor]')).toBeVisible();
		await page.keyboard.type('Added');
		await page.keyboard.press('Shift+Tab');
		await expect(slideTable(page).locator('[data-table-cell-editor]')).toBeVisible();
		await page.keyboard.press('Escape');
		await expect(cellWithText(page, 'Added')).toBeVisible();
	});

	test('pasting tab-separated text fills cells from the edited cell', async ({ page }) => {
		await loadDeck(page);
		const editor = await editCell(page, slideTable(page).locator('td').first());
		await editor.evaluate((node) => {
			const data = new DataTransfer();
			data.setData('text/plain', 'P1\tP2\r\nP3\tP4\r\n');
			node.dispatchEvent(
				new ClipboardEvent('paste', { clipboardData: data, bubbles: true, cancelable: true }),
			);
		});
		for (const text of ['P1', 'P2', 'P3', 'P4']) {
			await expect(cellWithText(page, text)).toBeVisible();
		}
	});

	test('Ctrl+V with a cell selected fills cells from it', async ({ page, context }) => {
		await context.grantPermissions(['clipboard-read', 'clipboard-write']);
		await loadDeck(page);
		await page.evaluate(() => navigator.clipboard.writeText('V1\tV2\r\nV3\tV4\r\n'));
		await selectTableCell(page, slideTable(page).locator('td').first());
		await page.keyboard.press('Control+V');
		for (const text of ['V1', 'V2', 'V3', 'V4']) {
			await expect(cellWithText(page, text)).toBeVisible();
		}
	});

	test('Bold with a cell range selected bolds every selected cell', async ({ page }) => {
		await loadDeck(page);
		const cells = slideTable(page).locator('tr').first().locator('td');
		await selectTableCell(page, cells.nth(0));
		const second = await centreOf(cells.nth(1));
		await page.keyboard.down('Shift');
		await page.mouse.click(second.x, second.y);
		await page.keyboard.up('Shift');
		await selectRibbonTab(page, 'Home');
		await page.getByRole('button', { name: 'Bold', exact: true }).first().click();
		const weight = (cell: Locator) =>
			cell.evaluate((td) => {
				const text = [...td.querySelectorAll('span')].find((span) => span.textContent?.trim());
				return Number.parseInt(getComputedStyle(text ?? td).fontWeight, 10);
			});
		await expect.poll(() => weight(cells.nth(0))).toBeGreaterThanOrEqual(600);
		await expect.poll(() => weight(cells.nth(1))).toBeGreaterThanOrEqual(600);
	});
});
