import { expect, test, type Page } from '@playwright/test';
import { editor, goToCell, grid, newWorkbook, ribbon, typeInActiveCell } from './helpers';

const home = (page: Page) => ribbon(page).locator('.ribbon-panel[data-tab="home"]');

/** Each Home group's label and width, and whether it is folded into its dropdown button. */
const homeGroups = (page: Page) =>
	home(page).evaluate((panel) =>
		[...panel.querySelectorAll<HTMLElement>(':scope > .ribbon-group')].map((group) => ({
			label: group.dataset.label,
			folded: group.hasAttribute('data-collapsed'),
			width: group.offsetWidth,
		})),
	);

/** Lets the ribbon's resize observers fit it to the window. */
const settle = (page: Page) =>
	page.evaluate(
		() => new Promise((done) => requestAnimationFrame(() => requestAnimationFrame(done))),
	);

test('a folded group unfolds to its own width, without its dropdown button', async ({ page }) => {
	await page.setViewportSize({ width: 2200, height: 900 });
	await newWorkbook(page);
	const styles = home(page).locator('.ribbon-group[data-label="Styles"]');
	const gallery = styles.locator('.xve-cell-styles');
	// The Cell Styles previews show inline once the ribbon is fitted to the wide window.
	await expect(gallery).toHaveAttribute('mode', 'inline');
	await settle(page);
	const unfolded = await homeGroups(page);
	expect(unfolded.filter((group) => group.folded)).toEqual([]);

	// A narrow window folds Styles into its dropdown button...
	await page.setViewportSize({ width: 1000, height: 900 });
	await expect(styles).toHaveAttribute('data-collapsed', '');
	await expect(styles.locator('.ribbon-overflow-button')).toBeVisible();

	// ...and widening it again gives every group back the width it had, with no dropdown button
	// left in an unfolded group to push the groups after it into a fold.
	await page.setViewportSize({ width: 2200, height: 900 });
	await expect(gallery).toHaveAttribute('mode', 'inline');
	await expect.poll(() => homeGroups(page)).toEqual(unfolded);
	await expect(styles.locator('.ribbon-overflow-button')).toBeHidden();
});

test.describe('on a phone', () => {
	test.use({ viewport: { width: 320, height: 720 }, hasTouch: true, isMobile: true });

	test('reaches the right side of Function Library and commits a formula from its menu', async ({
		page,
	}) => {
		await newWorkbook(page);
		const workbook = editor(page);
		await ribbon(page).getByRole('tab', { name: 'Formulas', exact: true }).tap();
		await ribbon(page).getByRole('button', { name: 'Function Library', exact: true }).tap();
		const group = workbook.locator('.ribbon-overflow-panel');
		await expect(group).toBeInViewport({ ratio: 1 });
		await group.getByRole('button', { name: 'More', exact: true }).tap();
		const moreFunctions = group.getByRole('button', { name: 'More Functions', exact: true });
		await expect(moreFunctions).toBeInViewport({ ratio: 1 });
		await moreFunctions.tap();
		const average = workbook.getByRole('menuitem', { name: 'AVERAGE', exact: true });
		await expect(average).toBeVisible();
		await page.keyboard.press('Escape');
		await expect(average).toBeHidden();
		await expect(group).toBeVisible();
		await moreFunctions.tap();
		await average.tap();
		await page.keyboard.type('2,8)');
		await page.keyboard.press('Enter');
		await expect(grid(page)).toContainText('5');
	});

	test('toggles a data filter, undoes and redoes it, and saves from the quick access toolbar', async ({
		page,
	}) => {
		await newWorkbook(page);
		await typeInActiveCell(page, 'Product');
		await typeInActiveCell(page, 'Apples');
		await goToCell(page, 'A1');
		const commands = ribbon(page);
		await commands.getByRole('tab', { name: 'Data', exact: true }).tap();
		const group = commands.getByRole('button', { name: 'Sort & Filter', exact: true });
		if (await group.isVisible()) await group.tap();
		await commands.getByRole('button', { name: 'Filter', exact: true }).tap();
		const filter = grid(page).locator('.xg-fb');
		await expect(filter).toHaveCount(1);

		const quick = editor(page).locator('office-ui-title-bar');
		await quick.getByRole('button', { name: 'Undo', exact: true }).tap();
		await expect(filter).toHaveCount(0);
		await quick.getByRole('button', { name: 'Redo', exact: true }).tap();
		await expect(filter).toHaveCount(1);
		const download = page.waitForEvent('download');
		await quick.getByRole('button', { name: 'Save', exact: true }).tap();
		expect((await download).suggestedFilename()).toMatch(/\.xlsx$/);
	});
});
