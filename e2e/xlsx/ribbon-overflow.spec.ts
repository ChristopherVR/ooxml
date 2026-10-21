import { expect, test, type Page } from '@playwright/test';
import { newWorkbook, ribbon } from './helpers';

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
