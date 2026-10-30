import { expect, test } from '@playwright/test';
import { clickRibbonControl, editor, goToCell, newWorkbook, pageErrors, ribbon } from './helpers';

interface SplitView {
	split?: { x: number; y: number; topLeft?: { row: number; col: number } };
	topLeft?: { row: number; col: number };
	freeze?: unknown;
}

const view = (page: import('@playwright/test').Page) =>
	editor(page).evaluate(
		(node) =>
			(node as unknown as { workbook: { sheets: { view: SplitView }[] } }).workbook.sheets[0]!.view,
	);

test('View > Split splits at the active cell, scrolls panes on their own, drags and removes bars', async ({
	page,
}) => {
	const errors = pageErrors(page);
	await newWorkbook(page);
	await goToCell(page, 'D6');
	await ribbon(page).getByRole('tab', { name: 'View', exact: true }).click();
	const split = ribbon(page).locator('[data-command="view.split"]');
	await clickRibbonControl(page, split);

	// Four panes and a bar on each axis.
	for (const name of ['corner', 'top', 'left', 'main'])
		await expect(editor(page).locator(`.xg-q-${name}`)).toBeVisible();
	await expect(editor(page).locator('.xg-split-bar:visible')).toHaveCount(2);
	const created = (await view(page)).split!;
	expect(created.x).toBeGreaterThan(0);
	expect(created.y).toBeGreaterThan(0);
	expect((await view(page)).freeze).toBeUndefined();

	// Wheel over the top-left pane scrolls only that pane; the saved first cells follow.
	const corner = (await editor(page).locator('.xg-q-corner').boundingBox())!;
	await page.mouse.move(corner.x + corner.width / 2, corner.y + corner.height / 2);
	await page.mouse.wheel(0, 200);
	await expect.poll(async () => (await view(page)).split?.topLeft?.row ?? 0).toBeGreaterThan(0);
	expect((await view(page)).topLeft?.row ?? 0).toBe(0);

	// Wheel over the bottom-right pane scrolls the sheet's own scroller.
	const main = (await editor(page).locator('.xg-q-main').boundingBox())!;
	await page.mouse.move(main.x + main.width / 2, main.y + main.height / 2);
	await page.mouse.wheel(0, 400);
	await expect.poll(async () => (await view(page)).topLeft?.row ?? 0).toBeGreaterThan(0);

	// Dragging the vertical bar moves the split.
	const bar = (await editor(page).locator('.xg-freeze-v').boundingBox())!;
	const startX = bar.x + bar.width / 2;
	const startY = bar.y + 200;
	await page.mouse.move(startX, startY);
	await page.mouse.down();
	await page.mouse.move(startX + 80, startY, { steps: 5 });
	await page.mouse.up();
	await expect.poll(async () => (await view(page)).split?.x ?? 0).toBeGreaterThan(created.x + 600);

	// Double click removes a bar, then the other one, which removes the split.
	const horizontal = (await editor(page).locator('.xg-freeze-h').boundingBox())!;
	await page.mouse.dblclick(horizontal.x + horizontal.width - 40, horizontal.y + 2);
	await expect.poll(async () => (await view(page)).split?.y).toBe(0);
	const vertical = (await editor(page).locator('.xg-freeze-v').boundingBox())!;
	await page.mouse.dblclick(vertical.x + 2, vertical.y + vertical.height - 40);
	await expect.poll(async () => (await view(page)).split).toBeUndefined();
	await expect(editor(page).locator('.xg-split-bar:visible')).toHaveCount(0);
	expect(errors).toEqual([]);
});
