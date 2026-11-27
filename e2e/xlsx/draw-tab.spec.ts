import { expect, test, type Page } from '@playwright/test';
import JSZip from 'jszip';
import { openFixture } from './fixture-files';
import {
	FRAMEWORKS,
	editor,
	formulaBarValue,
	goToCell,
	grid,
	nameBoxValue,
	ribbon,
	typeInActiveCell,
} from './helpers';

const inks = (page: Page) => grid(page).locator('.xg-obj[data-ink]');

async function saved(page: Page): Promise<JSZip> {
	const list = await editor(page).evaluate(async (node) =>
		Array.from(await (node as unknown as { saveBytes(): Promise<Uint8Array> }).saveBytes()),
	);
	return JSZip.loadAsync(new Uint8Array(list));
}

/** Drags across the grid from (x, y), offsets from the grid's top-left corner. */
async function drawOnGrid(page: Page, x: number, y: number, points: [number, number][]) {
	const box = (await grid(page).boundingBox())!;
	await page.mouse.move(box.x + x, box.y + y);
	await page.mouse.down();
	for (const [dx, dy] of points)
		await page.mouse.move(box.x + x + dx, box.y + y + dy, { steps: 3 });
	await page.mouse.up();
}

const wave: [number, number][] = Array.from({ length: 16 }, (_, i) => [
	i * 10,
	Math.round(Math.sin(i / 2) * 20),
]);

for (const framework of FRAMEWORKS)
	test(`${framework}: Draw tab strokes become Excel ink that saves as xdr:contentPart`, async ({
		page,
	}) => {
		await page.setViewportSize({ width: 1600, height: 900 });
		await openFixture(page, framework, 'excel-ink.xlsx');
		await expect(inks(page)).toHaveCount(1);
		await ribbon(page).getByRole('tab', { name: 'Draw', exact: true }).click();
		const panel = ribbon(page);
		await expect(panel.getByRole('button', { name: 'Ink to Math', exact: true })).toBeDisabled();
		const red = panel.getByRole('button', { name: 'Pen: 0.5 mm, Red' });
		await red.click();
		await expect(red).toHaveAttribute('aria-pressed', 'true');
		await drawOnGrid(page, 200, 350, wave);
		await expect(inks(page)).toHaveCount(2);
		// A stroke right after joins the same drawing.
		await drawOnGrid(page, 200, 380, [
			[60, 0],
			[120, 4],
		]);
		await expect(inks(page)).toHaveCount(2);
		await expect(inks(page).nth(1).locator('path')).toHaveCount(2);

		const zip = await saved(page);
		const drawing = await zip.file('xl/drawings/drawing1.xml')!.async('string');
		expect(drawing.match(/<xdr:contentPart/g)).toHaveLength(2);
		const ink = await zip.file('xl/ink/ink2.xml')!.async('string');
		expect(ink).toContain('value="#E81123"');
		// Excel's own ink part is kept as it was.
		expect(await zip.file('xl/ink/ink1.xml')!.async('string')).toContain('inkml:ink');

		// The eraser removes what it touches.
		await panel.getByRole('button', { name: 'Eraser', exact: true }).click();
		const target = (await inks(page).nth(1).boundingBox())!;
		for (let pass = 0; pass < 5 && (await inks(page).count()) > 1; pass++) {
			await page.mouse.move(target.x - 4, target.y + (target.height * pass) / 4);
			await page.mouse.down();
			await page.mouse.move(
				target.x + target.width + 4,
				target.y + target.height * (1 - pass / 4),
				{
					steps: 30,
				},
			);
			await page.mouse.up();
		}
		await expect(inks(page)).toHaveCount(1);
	});

test('Excel Draw: lasso, Ink to Shape, pen settings and Ink Replay', async ({ page }) => {
	await page.setViewportSize({ width: 1600, height: 900 });
	await openFixture(page, 'vanilla', 'excel-ink.xlsx');
	await ribbon(page).getByRole('tab', { name: 'Draw', exact: true }).click();
	const panel = ribbon(page);
	const black = panel.getByRole('button', { name: 'Pen: 0.5 mm, Black' });
	await black.click();
	await black.click();
	const menu = panel.getByRole('dialog', { name: 'Pen: 0.5 mm, Black' });
	await expect(menu).toBeVisible();
	// Office's ink colours come first, then the standard colours.
	await menu.getByRole('menuitemradio', { name: 'Green', exact: true }).first().click();
	await expect(panel.getByRole('button', { name: 'Pen: 0.5 mm, Green' })).toBeVisible();
	await page.keyboard.press('Escape');
	await expect(menu).toHaveCount(0);

	// A straight stroke, lassoed and turned into a line shape.
	await drawOnGrid(page, 200, 420, [
		[80, 0],
		[160, 0],
		[220, 0],
	]);
	await expect(inks(page)).toHaveCount(2);
	await panel.getByRole('button', { name: 'Lasso Select', exact: true }).click();
	await drawOnGrid(page, 180, 400, [
		[260, 0],
		[260, 40],
		[0, 40],
		[0, 0],
	]);
	await expect(grid(page).locator('.xg-ink-lasso')).toHaveCount(1);
	await panel.getByRole('button', { name: 'Ink to Shape', exact: true }).click();
	await expect(inks(page)).toHaveCount(1);
	await expect(grid(page).locator('.xg-shape')).toHaveCount(1);

	await panel.getByRole('button', { name: 'Ink Replay', exact: true }).click();
	await expect(grid(page).locator('.xg-ink-replay path').first()).toBeAttached();
	await page.keyboard.press('Escape');
	await expect(grid(page).locator('.xg-ink-replay')).toHaveCount(0);
});

test('Excel Draw: a pen press never selects cells; one eraser stroke is one undo step', async ({
	page,
}) => {
	await page.setViewportSize({ width: 1600, height: 900 });
	await openFixture(page, 'vanilla', 'excel-ink.xlsx');
	await goToCell(page, 'B2');
	await ribbon(page).getByRole('tab', { name: 'Draw', exact: true }).click();
	const panel = ribbon(page);
	await panel.getByRole('button', { name: 'Pen: 0.5 mm, Red' }).click();
	// Two strokes far apart, drawn at once: each is ink, and the cell selection stays.
	await drawOnGrid(page, 300, 300, wave);
	await drawOnGrid(page, 300, 340, [
		[200, 0],
		[400, 4],
	]);
	expect(await nameBoxValue(page)).toBe('B2');
	await expect(inks(page)).toHaveCount(2);
	const drawn = inks(page).nth(1);
	await expect(drawn.locator('path')).toHaveCount(2);
	const before = (await drawn.boundingBox())!;

	// One sweep over the long stroke erases it; the box shrinks to the wave that is left.
	await panel.getByRole('button', { name: 'Eraser', exact: true }).click();
	const box = (await grid(page).boundingBox())!;
	await page.mouse.move(box.x + 650, box.y + 330);
	await page.mouse.down();
	await page.mouse.move(box.x + 650, box.y + 360, { steps: 20 });
	await page.mouse.up();
	await expect(drawn.locator('path')).toHaveCount(1);
	const after = (await drawn.boundingBox())!;
	expect(after.width).toBeLessThan(before.width - 100);
	// Undo brings the whole stroke back in one step.
	await page.keyboard.press('Control+z');
	await expect(drawn.locator('path')).toHaveCount(2);
});

test('Excel Draw: strokes drawn while the Draw surface still loads are kept, in order', async ({
	page,
}) => {
	await page.setViewportSize({ width: 1280, height: 800 });
	await openFixture(page, 'vanilla', 'excel-ink.xlsx');
	await ribbon(page).getByRole('tab', { name: 'Draw', exact: true }).click();
	// Right after the pen is picked, on a cold page: the first stroke can land before the surface
	// is over the cells, and a quick second one must not race it.
	await ribbon(page).getByRole('button', { name: 'Pen: 0.5 mm, Red' }).click();
	await drawOnGrid(page, 300, 300, wave);
	await drawOnGrid(page, 300, 330, [
		[100, 0],
		[200, 6],
	]);
	await expect(inks(page)).toHaveCount(2);
	await expect(inks(page).nth(1).locator('path')).toHaveCount(2);
	expect(await nameBoxValue(page)).toBe('A1');
});

test('Excel: Undo and Redo on the title bar give focus back to the sheet', async ({ page }) => {
	await page.setViewportSize({ width: 1280, height: 800 });
	await openFixture(page, 'vanilla', 'excel-ink.xlsx');
	await goToCell(page, 'C3');
	await typeInActiveCell(page, '5');
	await goToCell(page, 'C4');
	await typeInActiveCell(page, '6');
	await editor(page).getByRole('button', { name: 'Undo', exact: true }).click();
	// Ctrl+Z goes on working without a click in the sheet first.
	await page.keyboard.press('Control+z');
	await goToCell(page, 'C3');
	expect(await formulaBarValue(page)).toBe('');
	await editor(page).getByRole('button', { name: 'Redo', exact: true }).click();
	await page.keyboard.press('Control+y');
	await goToCell(page, 'C4');
	expect(await formulaBarValue(page)).toBe('6');
});
