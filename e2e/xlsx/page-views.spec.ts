import { expect, test } from '@playwright/test';
import {
	editor,
	goToCell,
	grid,
	newWorkbook,
	ribbon,
	statusBar,
	typeInActiveCell,
} from './helpers';

const view = (page: Parameters<typeof newWorkbook>[0], id: 'normal' | 'pageLayout' | 'pageBreak') =>
	statusBar(page).locator(`[data-id="${id}"]`);

async function seed(page: Parameters<typeof newWorkbook>[0]) {
	await newWorkbook(page);
	await typeInActiveCell(page, 'First');
	await goToCell(page, 'A12');
	await typeInActiveCell(page, 'Last');
}

test('status bar and View tab switch between Normal, Page Layout and Page Break Preview', async ({
	page,
}) => {
	await seed(page);
	await expect(view(page, 'normal')).toHaveAttribute('aria-pressed', 'true');
	await view(page, 'pageLayout').click();
	await expect(view(page, 'pageLayout')).toHaveAttribute('aria-pressed', 'true');
	await expect(editor(page).locator('.xpl-page').first()).toBeVisible();
	await expect(editor(page).locator('.xpl-header .xpl-placeholder')).toHaveText('Add header');
	await expect(editor(page).locator('.xg-root')).toBeHidden();

	await ribbon(page).getByRole('tab', { name: 'View', exact: true }).click();
	await ribbon(page).locator('[data-command="view.page-break-preview"]').click();
	await expect(view(page, 'pageBreak')).toHaveAttribute('aria-pressed', 'true');
	await expect(editor(page).locator('.xpl-root')).toBeHidden();
	await expect(editor(page).locator('.xgpb-mark').first()).toHaveText('Page 1');
	await expect(ribbon(page).locator('[data-command="view.custom-views"]')).toBeDisabled();

	await ribbon(page).locator('[data-command="view.normal"]').click();
	await expect(view(page, 'normal')).toHaveAttribute('aria-pressed', 'true');
	await expect(editor(page).locator('.xgpb-layer')).toBeHidden();

	// The view is part of the sheet and survives a save and reopen.
	await view(page, 'pageLayout').click();
	const bytes = await editor(page).evaluate(async (node) => [
		...new Uint8Array(
			await (await (node as unknown as { save(): Promise<Blob> }).save()).arrayBuffer(),
		),
	]);
	await editor(page).evaluate(
		async (node, data) =>
			(node as unknown as { load(bytes: Uint8Array, name: string): Promise<void> }).load(
				new Uint8Array(data),
				'views.xlsx',
			),
		bytes,
	);
	await expect(view(page, 'pageLayout')).toHaveAttribute('aria-pressed', 'true');
	await expect(editor(page).locator('.xpl-page').first()).toBeVisible();
});

test('Page Layout view edits a header and a cell', async ({ page }) => {
	await seed(page);
	await view(page, 'pageLayout').click();
	const pageNode = editor(page).locator('.xpl-page').first();
	await pageNode.locator('.xpl-header .xpl-placeholder').click();
	const inputs = editor(page).locator('.xpl-band-input');
	await expect(inputs).toHaveCount(3);
	await inputs.nth(1).fill('Page &P of &N');
	await inputs.nth(1).press('Enter');
	await expect(pageNode.locator('.xpl-header')).toContainText('Page 1 of 1');
	await pageNode.locator('.xpl-inner td', { hasText: 'First' }).dblclick();
	const box = editor(page).locator('.xpl-edit');
	await expect(box).toHaveValue('First');
	await box.fill('Changed');
	await box.press('Enter');
	await expect(pageNode.locator('.xpl-inner')).toContainText('Changed');
	await view(page, 'normal').click();
	await expect(grid(page)).toContainText('Changed');
});

test('Page Break Preview moves a manual break by dragging', async ({ page }) => {
	await seed(page);
	await goToCell(page, 'A6');
	await ribbon(page).getByRole('tab', { name: 'Page Layout', exact: true }).click();
	await ribbon(page).getByRole('button', { name: 'Breaks', exact: true }).click();
	await editor(page).getByRole('menuitem', { name: 'Insert Row Page Break', exact: true }).click();
	await view(page, 'pageBreak').click();
	await expect(editor(page).locator('.xgpb-line.xgpb-manual')).toHaveCount(1);
	const line = editor(page).locator('.xgpb-line.xgpb-manual');
	const top = () => line.evaluate((node) => parseFloat((node as HTMLElement).style.top));
	const before = await top();
	const handle = editor(page).locator('.xgpb-grab.xgpb-row').nth(2);
	const box = (await handle.boundingBox())!;
	await page.mouse.move(box.x + 40, box.y + box.height / 2);
	await page.mouse.down();
	await page.mouse.move(box.x + 40, box.y + box.height / 2 + 60, { steps: 6 });
	await page.mouse.up();
	await expect.poll(top).toBeGreaterThan(before + 40);
	await page.keyboard.press('Control+z');
	await expect.poll(top).toBeLessThan(before + 5);
});

test('Page Layout view fills the page, aligns the rulers and fits narrow screens', async ({
	page,
}) => {
	await seed(page);
	await view(page, 'pageLayout').click();
	const first = editor(page).locator('.xpl-page').first();
	// Default Letter page: ten default columns across, blank cells included.
	await expect(first.locator('.xpl-colhdr')).toHaveCount(10);
	await expect(first.locator('.xpl-inner tr').first().locator('td')).toHaveCount(10);
	const pageBox = (await first.boundingBox())!;
	const ruler = (await editor(page).locator('.xpl-ruler-h .xpl-ruler-area').first().boundingBox())!;
	expect(Math.abs(ruler.x - pageBox.x)).toBeLessThan(1.5);
	expect(Math.abs(ruler.width - pageBox.width)).toBeLessThan(1.5);
	const rulerV = (await editor(page)
		.locator('.xpl-ruler-v .xpl-ruler-area')
		.first()
		.boundingBox())!;
	expect(Math.abs(rulerV.width - 20)).toBeLessThan(1.5);
	await page.setViewportSize({ width: 375, height: 812 });
	await expect
		.poll(() =>
			editor(page)
				.locator('.xpl-root')
				.evaluate((n) => n.scrollWidth - n.clientWidth),
		)
		.toBeLessThanOrEqual(1);
	const narrow = (await editor(page).locator('.xpl-page').first().boundingBox())!;
	const root = (await editor(page).locator('.xpl-root').boundingBox())!;
	expect(narrow.x + narrow.width).toBeLessThanOrEqual(root.x + root.width + 1);
});
