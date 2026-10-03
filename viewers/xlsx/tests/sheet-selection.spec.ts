import { expect, test, type Page } from '@playwright/test';
import {
	editor,
	goToCell,
	grid,
	nameBoxValue,
	newWorkbook,
	openSample,
	pageErrors,
	sheetTabs,
} from './helpers';

const selection = (page: Page) =>
	editor(page).evaluate((node) => (node as unknown as { getSelection(): string }).getSelection());
const activeSheet = (page: Page) =>
	editor(page).evaluate((node) => (node as unknown as { activeSheet: number }).activeSheet);

/** Records every selection-change and sheet-change the element dispatches from now on. */
async function recordEvents(page: Page) {
	await editor(page).evaluate((node) => {
		const log: { type: string; detail: unknown }[] = [];
		(window as unknown as { xlsxEvents: typeof log }).xlsxEvents = log;
		for (const type of ['selection-change', 'sheet-change'])
			node.addEventListener(type, (event) =>
				log.push({ type, detail: (event as CustomEvent).detail }),
			);
	});
	return (type: string) =>
		page.evaluate(
			(name) =>
				(
					window as unknown as { xlsxEvents: { type: string; detail: unknown }[] }
				).xlsxEvents.filter((event) => event.type === name),
			type,
		);
}

test('switching sheet tabs comes back to each sheet’s last selection', async ({ page }) => {
	const errors = pageErrors(page);
	await openSample(page);
	await goToCell(page, 'C5');
	await sheetTabs(page).getByRole('tab', { name: 'Budget' }).click();
	await expect.poll(() => activeSheet(page)).toBe(1);
	await goToCell(page, 'B3');
	await sheetTabs(page).getByRole('tab', { name: 'Sales' }).click();
	await expect.poll(() => nameBoxValue(page)).toBe('C5');
	expect(await selection(page)).toBe('C5');
	await sheetTabs(page).getByRole('tab', { name: 'Budget' }).click();
	await expect.poll(() => nameBoxValue(page)).toBe('B3');
	expect(errors).toEqual([]);
});

test('a column header selection is reported once, as B:B, everywhere', async ({ page }) => {
	await newWorkbook(page);
	const events = await recordEvents(page);
	await grid(page).locator('.xg-hdr-col .xg-hd', { hasText: /^B$/u }).click();
	await expect.poll(() => selection(page)).toBe('B:B');
	await expect(page.locator('#build-stamp')).toHaveAttribute('data-selection', 'B:B');
	expect(await events('selection-change')).toEqual([
		{ type: 'selection-change', detail: { sheet: 0, ref: 'B:B', active: 'B1' } },
	]);
});

test('deleting the active sheet announces sheet-change and resets the selection', async ({
	page,
}) => {
	const errors = pageErrors(page);
	await newWorkbook(page);
	await sheetTabs(page).getByRole('button', { name: 'New sheet' }).click();
	await expect(sheetTabs(page)).toContainText('Sheet2');
	await sheetTabs(page).getByRole('tab', { name: 'Sheet2' }).click();
	await goToCell(page, 'D7');
	await sheetTabs(page).getByRole('tab', { name: 'Sheet1' }).click();
	await expect.poll(() => activeSheet(page)).toBe(0);
	await goToCell(page, 'C3');
	const events = await recordEvents(page);
	await sheetTabs(page).getByRole('tab', { name: 'Sheet1' }).click({ button: 'right' });
	await editor(page).getByRole('menuitem', { name: 'Delete' }).click();
	await expect(sheetTabs(page).getByRole('tab')).toHaveCount(1);
	expect(await activeSheet(page)).toBe(0);
	expect(await events('sheet-change')).toEqual([
		{ type: 'sheet-change', detail: { index: 0, name: 'Sheet2' } },
	]);
	await expect.poll(() => nameBoxValue(page)).toBe('D7');
	expect(await selection(page)).toBe('D7');
	expect(errors).toEqual([]);
});
